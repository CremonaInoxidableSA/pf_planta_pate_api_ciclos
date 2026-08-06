from sqlalchemy.orm import Session, sessionmaker
from datetime import datetime, timedelta
from collections import defaultdict, deque
from contextlib import contextmanager
from statistics import mean
from models.ciclo import Ciclo
from models.receta import Receta
from models.sensoresAA import SensoresAA
from models.sensoresIO import SensoresIO
from models.estadoCiclo import EstadoCiclo

import os
import json
import logging
import asyncio
import math
import threading

# Se importa el modulo completo, no solamente el engine, para poder leer
# config.db.engine en cada operacion. Despues de una reconexion de MySQL el
# modulo publica un engine nuevo y las sesiones deben vincularse a ese.
from config import db as config_db

logger = logging.getLogger("uvicorn")

EQUIPOS_MAP = {
    1:  "PF-L1-COCINA-1",
    2:  "PF-L1-COCINA-2",
    3:  "PF-L1-COCINA-3",
    4:  "PF-L2-COCINA-4",
    5:  "PF-L2-COCINA-5",
    6:  "PF-L2-COCINA-6",
    7:  "PF-L1-ENFRIADOR-1",
    8:  "PF-L1-ENFRIADOR-2",
    9:  "PF-L1-ENFRIADOR-3",
    10: "PF-L1-ENFRIADOR-4",
    11: "PF-L2-ENFRIADOR-5",
    12: "PF-L2-ENFRIADOR-6",
    13: "PF-L2-ENFRIADOR-7",
    14: "PF-L2-ENFRIADOR-8",
}

CAMPOS_UNIDAD = [
    "ESTADO_EQUIPO", "NUMERO_EQUIPO", "NUMERO_RECETA", "PASO_ACTUAL",
    "TEMP_AGUA", "TEMP_INGRESO", "TEMP_PRODUCTO", "NIVEL_AGUA",
    "CANTIDAD_TORRES", "PESO_PRODUCTO", "LOTE_CICLO", "CICLO_TIPO_FIN",
    "FILTRO_SUCCION_AGUA", "CARGA_AGUA", "BOMBA_CENTRIFUGA", "TIEMPO_TRANS",
    # Cocina
    "VAPOR_SERPENTINA", "VAPOR_SERPENTINA_ACC", "VAPOR_VIVO", "VAPOR_VIVO_ACC",
    # Enfriador
    "AMONIACO", "AMONIACO_ACC", "VAPOR_LIMPIEZA", "VAPOR_LIMPIEZA_ACC",
]

# -----------------------------------------------------------------------
# Mapeo OPC campo --> id sensor booleano (tabla `sensores`)
#
#   6  Bomba centrifuga         ENTRADA
#   8  Valvula amoniaco         ENTRADA  <- AMONIACO_ACC
#  10  Vapor vivo limpieza      ENTRADA  <- VAPOR_LIMPIEZA_ACC
#  13  Vapor serpentina accion  SALIDA   <- VAPOR_SERPENTINA_ACC
#  14  Vapor vivo accionamiento SALIDA   <- VAPOR_VIVO_ACC
#  15  Agua toma de filtro      SALIDA   <- FILTRO_SUCCION_AGUA
#  16  Carga de agua            SALIDA   <- CARGA_AGUA
#
# CICLO_TIPO_FIN no tiene sensor propio en la tabla; se omite.
# -----------------------------------------------------------------------

IO_SENSOR_MAP_COCINA: dict[str, int] = {
    "FILTRO_SUCCION_AGUA":  15,
    "CARGA_AGUA":           16,
    "BOMBA_CENTRIFUGA":     6,
    "VAPOR_SERPENTINA_ACC": 13,
    "VAPOR_SERPENTINA":     7,
    "VAPOR_VIVO_ACC":       14,
    "VAPOR_VIVO":           9
}

IO_SENSOR_MAP_ENFRIADOR: dict[str, int] = {
    "FILTRO_SUCCION_AGUA": 15,
    "CARGA_AGUA":          16,
    "BOMBA_CENTRIFUGA":    6,
    "AMONIACO_ACC":        19,
    "VAPOR_LIMPIEZA_ACC":  18,
    "VAPOR_LIMPIEZA":      10,
    "AMONIACO":            8
}

# Mapeo campo JSON --> id sensor anal�gico (SensoresAA)
AA_SENSOR_MAP: dict[str, int] = {
    "temp_agua":    1,
    "temp_ingreso": 2,
    "temp_prod":    3,
    "niv_agua":     5,
}

# Los cuatro sensores permanecen juntos en cada registro JSONL. El filtrado
# termico solo usa estos tres nombres, sin modificar AA_SENSOR_MAP.
CAMPOS_TEMPERATURA = ("temp_agua", "temp_ingreso", "temp_prod")
CAMPOS_HISTORIAL = (
    "id_historial", "tiempo", "estado", "idCiclo", "lote",
    "tiempo_trans_opc", "temp_agua", "temp_ingreso", "temp_prod", "niv_agua",
)

UMBRAL_CAMBIO = 0.5
VENTANA_PROMEDIO_SEGUNDOS = 10
PERSISTENCIA_REQUERIDA = 3
TIEMPO_MAXIMO_SEGUNDOS = 300
HISTORIAL_PREVIEW_MAX = 100

# Un hueco mayor a este umbral entre dos capturas consecutivas no se atribuye
# al ultimo estado conocido: se registra como un tramo propio. La captura
# normal es de aproximadamente una muestra por segundo.
HUECO_SIN_DATOS_SEGUNDOS = 5
ESTADO_SIN_DATOS = "SIN DATOS / DESCONEXION"


# -----------------------------------------------------------------------
# Helpers puros (sin estado de clase)
# -----------------------------------------------------------------------

def datetime_to_string(obj):
    if isinstance(obj, datetime):
        return obj.strftime("%Y-%m-%d %H:%M:%S")
    raise TypeError(f"Object of type {type(obj)} is not JSON serializable")


def _numero_finito_o_none(valor):
    """Normaliza lecturas OPC numericas; cero siempre es un valor valido."""
    if valor is None or isinstance(valor, bool):
        return valor
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(numero):
        return None
    if isinstance(valor, int):
        return valor
    return numero


def _redondear_para_presentacion(valor):
    """Formato visual compatible con el frontend; nunca se guarda en JSONL."""
    numero = _numero_finito_o_none(valor)
    if numero is None or isinstance(numero, bool):
        return "Sin registro"
    return round(float(numero), 1)


def _parse_tiempo(s: str) -> datetime:
    return datetime.strptime(s, "%Y-%m-%d %H:%M:%S")


def _delta_str(inicio: datetime, fin: datetime) -> str:
    segundos = max(0, int((fin - inicio).total_seconds()))
    h, r = divmod(segundos, 3600)
    m, s = divmod(r, 60)
    return f"{h:02}:{m:02}:{s:02}"


def _tramo(nombre: str, inicio: datetime, fin: datetime) -> dict:
    return {
        "nombre":             nombre,
        "fechaInicio":        inicio,
        "fechaFin":           fin,
        "tiempoTranscurrido": _delta_str(inicio, fin),
    }


def _construir_tramos_estado(
    historial: list,
    umbral_hueco_segundos: int = HUECO_SIN_DATOS_SEGUNDOS,
) -> list[dict]:
    """
    Agrupa filas consecutivas del mismo estado en tramos e intercala tramos
    ESTADO_SIN_DATOS cuando la distancia entre dos capturas consecutivas supera
    umbral_hueco_segundos.

    Un hueco NO se atribuye al ultimo estado conocido: el intervalo faltante se
    registra como periodo sin datos, de modo que una desconexion no se informe
    como tiempo productivo ni se confunda con PAUSADO.

    Maneja:
    - Historial vacio         -> []
    - Una sola muestra        -> un tramo con fechaInicio == fechaFin
    - Timestamps repetidos    -> el tramo se extiende sin error
    - Timestamps fuera orden  -> se ordenan antes de procesar
    - Huecos largos           -> tramo ESTADO_SIN_DATOS intercalado

    Los tramos generados no se solapan y son contiguos a ambos lados de cada
    hueco: el tramo sin datos empieza exactamente en la ultima captura valida y
    termina exactamente en la primera captura posterior.

    Retorna:
        [
          {
            "nombre":             str,
            "fechaInicio":        datetime,
            "fechaFin":           datetime,
            "tiempoTranscurrido": "HH:MM:SS",
          },
          ...
        ]
    """
    if not historial:
        return []

    ordenado = sorted(
        historial,
        key=lambda r: (
            r.get("tiempo", ""),
            int(r.get("id_historial") or 0),
        ),
    )

    nombre_actual = ordenado[0]["estado"]
    inicio_actual = _parse_tiempo(ordenado[0]["tiempo"])
    fin_actual    = inicio_actual
    tramos        = []

    for fila in ordenado[1:]:
        t      = _parse_tiempo(fila["tiempo"])
        nombre = fila["estado"]

        # El historial ya esta ordenado, pero un timestamp repetido produce
        # hueco 0 y nunca un intervalo negativo.
        hueco = (t - fin_actual).total_seconds()

        if hueco > umbral_hueco_segundos:
            # Cerrar el estado observado y registrar el intervalo faltante.
            tramos.append(_tramo(nombre_actual, inicio_actual, fin_actual))
            tramos.append(_tramo(ESTADO_SIN_DATOS, fin_actual, t))
            nombre_actual = nombre
            inicio_actual = t
            fin_actual    = t
            continue

        if nombre == nombre_actual:
            if t > fin_actual:
                fin_actual = t
        else:
            tramos.append(_tramo(nombre_actual, inicio_actual, fin_actual))
            nombre_actual = nombre
            inicio_actual = t
            fin_actual    = t

    # Ultimo tramo (siempre existe)
    tramos.append(_tramo(nombre_actual, inicio_actual, fin_actual))
    return tramos


def _cerrar_tramos_hasta(
    tramos: list,
    fecha_fin: datetime,
    umbral_hueco_segundos: int = HUECO_SIN_DATOS_SEGUNDOS,
) -> list:
    """
    Lleva la lista de tramos hasta la fecha de cierre oficial del ciclo.

    Si entre la ultima captura y el cierre pasaron mas de umbral_hueco_segundos,
    ese intervalo se registra como ESTADO_SIN_DATOS en vez de estirar el ultimo
    estado observado. Si la diferencia es normal (el historial termina un par de
    segundos antes del cierre), el ultimo tramo simplemente se extiende.
    """
    if not tramos:
        return tramos

    ultimo = tramos[-1]
    if fecha_fin <= ultimo["fechaFin"]:
        return tramos

    hueco = (fecha_fin - ultimo["fechaFin"]).total_seconds()
    if hueco > umbral_hueco_segundos:
        tramos.append(
            _tramo(ESTADO_SIN_DATOS, ultimo["fechaFin"], fecha_fin)
        )
    else:
        ultimo["fechaFin"] = fecha_fin
        ultimo["tiempoTranscurrido"] = _delta_str(
            ultimo["fechaInicio"], fecha_fin
        )
    return tramos


# -----------------------------------------------------------------------
# Clase principal
# -----------------------------------------------------------------------

class ObtenerNodosOpcUA:

    ESTADOS_EQUIPO_MAP = {
        1: "PRE OPERACIONAL",
        2: "OPERACIONAL",
        3: "PAUSADO",
        4: "INACTIVO",
        5: "CANCELADO",
        6: "FINALIZADO",
        7: "LIMPIEZA",
    }

    ESTADOS_CONTINUOS = {"PRE OPERACIONAL", "OPERACIONAL", "PAUSADO"}
    ESTADOS_FIN       = {"FINALIZADO", "CANCELADO", "INACTIVO"}

    RECON_SIN_HISTORIAL = "SIN_HISTORIAL"
    RECON_MISMO_CICLO   = "MISMO_CICLO"
    RECON_NUEVO_CICLO   = "NUEVO_CICLO"
    RECON_CONFIRMAR_TT   = "CONFIRMAR_TT"
    RECON_INDETERMINADA  = "INDETERMINADA"

    def __init__(self, conexion_servidor):
        self.conexion_servidor  = conexion_servidor

        # ---------------------------------------------------------------
        # Sesiones SQLAlchemy
        #
        # Ya no existe una sesion unica de larga vida: se abre una por
        # operacion y se cierra siempre. La fabrica se reconstruye cuando
        # config.db publica un engine nuevo (reconexion de MySQL), de modo
        # que la adquisicion se recupera sin reiniciar el contenedor.
        # ---------------------------------------------------------------
        self._session_factory = None
        self._session_engine = None
        self._session_factory_lock = threading.RLock()

        self._equipos: list[dict] = []
        self._recetario_cache: dict = {}
        self._pf_l1_node = None

        # ---------------------------------------------------------------
        # Trazabilidad IO en tiempo real
        #
        # _io_state[equipo_key][campo_opc] = {
        #     "valor":    bool,
        #     "inicio":   datetime (sin microsegundos),
        #     "id_ciclo": int,
        # }
        #
        # Cuando una se�al cambia de valor el tramo anterior se persiste
        # en BD inmediatamente y se abre uno nuevo en memoria.
        # Al cierre del ciclo se persisten los tramos a�n abiertos.
        # ---------------------------------------------------------------
        self._io_state: dict[str, dict[str, dict]] = {}

        # Estado JSONL. Cada archivo tiene su propio lock, contador y preview.
        # El directorio puede montarse en un volumen persistente de Docker con
        # OPC_HISTORIAL_DIR=/ruta/del/volumen.
        self._jsonl_dir = os.path.abspath(
            os.getenv("OPC_HISTORIAL_DIR", "data/opc_historial")
        )
        self._jsonl_fsync = os.getenv("OPC_JSONL_FSYNC", "0") == "1"
        self._jsonl_locks: dict[str, threading.RLock] = defaultdict(
            threading.RLock
        )
        self._jsonl_meta_lock = threading.RLock()
        self._jsonl_loaded: set[str] = set()
        self._jsonl_next_id: dict[str, int] = {}
        self._jsonl_preview: dict[str, deque] = {}
        self._jsonl_first_time: dict[str, str] = {}
        self._jsonl_last_time: dict[str, str] = {}
        self._jsonl_last_record: dict[str, dict] = {}

        # Reconciliacion por TIEMPO_TRANS. El numero de generacion cambia al
        # crear cada suscripcion y permite distinguir los valores recibidos
        # despues de una reconexion de los que pertenecian a la anterior.
        self._subscription_generation = 0
        self._equipos_pendientes_reconciliacion: set[str] = set()
        self._tt_retrocesos: dict[str, dict] = {}
        self._advertencias_reconciliacion: set[tuple] = set()

        os.makedirs(self._jsonl_dir, exist_ok=True)
        self._informar_archivos_pendientes()

    # -----------------------------------------------------------------------
    # Sesiones SQLAlchemy
    # -----------------------------------------------------------------------

    def _fabrica_de_sesiones(self):
        """
        Devuelve una sessionmaker vinculada al engine vigente.

        Se compara la identidad del engine publicado por config.db en lugar de
        cachearlo para siempre. Cuando el monitor de conexion reemplaza el
        engine despues de una caida de MySQL, la fabrica se reconstruye sola en
        la siguiente operacion.
        """
        engine_actual = getattr(config_db, "engine", None)
        if engine_actual is None:
            raise RuntimeError(
                "El engine de MySQL todavia no esta disponible"
            )

        with self._session_factory_lock:
            if (
                self._session_factory is None
                or self._session_engine is not engine_actual
            ):
                self._session_factory = sessionmaker(
                    autocommit=False,
                    autoflush=False,
                    bind=engine_actual,
                )
                self._session_engine = engine_actual
                logger.info(
                    "Fabrica de sesiones SQLAlchemy vinculada al engine "
                    "vigente de config.db."
                )
            return self._session_factory

    @contextmanager
    def _sesion_bd(self):
        """
        Abre una sesion para una operacion y la cierra siempre.

        Ante una excepcion se ejecuta rollback antes de cerrar, de modo que
        ninguna transaccion queda abierta ocupando una conexion del pool.
        """
        fabrica = self._fabrica_de_sesiones()
        session: Session = fabrica()
        try:
            yield session
        except Exception:
            try:
                session.rollback()
            except Exception:
                pass
            raise
        finally:
            try:
                session.close()
            except Exception:
                pass

    # -----------------------------------------------------------------------
    # Navegaciion del Arbol OPC
    # -----------------------------------------------------------------------

    def _sorted_children(self, node):
        def sort_key(n):
            name = n.get_browse_name().Name
            if name.startswith("[") and name.endswith("]"):
                try:
                    return int(name[1:-1])
                except Exception:
                    return name
            return name
        return sorted(node.get_children(), key=sort_key)

    def _children_dict(self, node):
        return {child.get_browse_name().Name: child for child in node.get_children()}

    def _read_array_object(self, array_obj_node):
        valores = []
        for child in self._sorted_children(array_obj_node):
            try:
                valores.append(child.get_value())
            except Exception:
                pass
        return valores

    def _navegar_unidades(self, grupo_node, linea: str, tipo: str, offset_id: int):
        nodos_a_suscribir = []
        for idx_local, unit_node in enumerate(self._sorted_children(grupo_node)):
            numero_local = idx_local + 1
            id_equipo    = numero_local + offset_id
            node_map: dict[str, str] = {}
            hijos = self._children_dict(unit_node)
            for campo in CAMPOS_UNIDAD:
                if campo in hijos:
                    node_id = hijos[campo].nodeid.to_string()
                    node_map[campo] = node_id
                    nodos_a_suscribir.append(hijos[campo])
            self._equipos.append({
                "linea":        linea,
                "tipo":         tipo,
                "numero_local": numero_local,
                "id_equipo":    id_equipo,
                "node_map":     node_map,
            })
            logger.info(
                f"  Registrado: {linea}/{tipo}/[{idx_local}] "
                f"-> id_equipo={id_equipo}, {len(node_map)} nodos"
            )
        return nodos_a_suscribir

    def _navegar_arbol_sync(self, root_node):
        """
        Recorre el arbol OPC una unica vez y devuelve la lista de nodos
        a suscribir.

        Objects / Server interfaces
          PF-L1 / COCINA L1    [0..2]
          PF-L1 / ENFRIADOR L1 [0..3]
          PF-L2 / COCINA L2    [0..2]
          PF-L2 / ENFRIADOR L2 [0..3]
        """
        nodos_suscritos  = []
        objects_node     = root_node.get_child(["0:Objects"])
        server_ifaces    = objects_node.get_child(["2:ServerInterfaces"])
        pf_l1            = server_ifaces.get_child(["2:PF-L1"])
        pf_l2            = server_ifaces.get_child(["2:PF-L2"])
        self._pf_l1_node = pf_l1

        grupos = [
            (pf_l1.get_child(["2:COCINA L1"]),    "PF-L1", "COCINA",    0),
            (pf_l1.get_child(["2:ENFRIADOR L1"]), "PF-L1", "ENFRIADOR", 6),
            (pf_l2.get_child(["2:COCINA L2"]),    "PF-L2", "COCINA",    3),
            (pf_l2.get_child(["2:ENFRIADOR L2"]), "PF-L2", "ENFRIADOR", 10),
        ]
        for grupo_node, linea, tipo, offset_id in grupos:
            logger.info(f"Navegando {linea} / {tipo}...")
            nodos_suscritos.extend(self._navegar_unidades(grupo_node, linea, tipo, offset_id))

        logger.info(
            f"Arbol OPC navegado: {len(self._equipos)} equipos, "
            f"{len(nodos_suscritos)} nodos a suscribir."
        )
        return nodos_suscritos

    # -----------------------------------------------------------------------
    # Recetario
    # -----------------------------------------------------------------------

    def _obtener_recetario_desde_opc_sync(self, pf_l1_node):
        recetario_cache = {}
        try:
            recetario_node = pf_l1_node.get_child(["4:RECETARIO"])
            for item_node in self._sorted_children(recetario_node):
                item_name = item_node.get_browse_name().Name
                try:
                    numero_receta = int(item_name[1:-1])
                except Exception:
                    continue
                hijos = self._children_dict(item_node)
                recetario_cache[numero_receta] = {
                    "NOMBRE":                hijos["NOMBRE"].get_value()                     if "NOMBRE"                in hijos else f"RECETA_{numero_receta:02}",
                    "PASOS":                 hijos["PASOS"].get_value()                      if "PASOS"                 in hijos else 0,
                    "TEMP_AGUA":             self._read_array_object(hijos["TEMP AGUA"])     if "TEMP AGUA"             in hijos else [],
                    "TEMP_PRODUCTO":         self._read_array_object(hijos["TEMP PRODUCTO"]) if "TEMP PRODUCTO"         in hijos else [],
                    "TIEMPO_CORTE":          self._read_array_object(hijos["TIEMPO CORTE"])  if "TIEMPO CORTE"          in hijos else [],
                    "TIPO_CORTE":            self._read_array_object(hijos["TIPO CORTE"])    if "TIPO CORTE"            in hijos else [],
                    "TIEMPO_CORTE_ENFRIADO": hijos["TIEMPO CORTE ENFRIADO"].get_value()      if "TIEMPO CORTE ENFRIADO" in hijos else None,
                    "TEMP_CORTE_ENFRIADO":   hijos["TEMP CORTE ENFRIADO"].get_value()        if "TEMP CORTE ENFRIADO"   in hijos else None,
                    "TIPO_CORTE_ENFRIADO":   hijos["TIPO CORTE ENFRIADO"].get_value()        if "TIPO CORTE ENFRIADO"   in hijos else None,
                    "PESO_X_TORRE":          hijos["PESO X TORRE"].get_value()               if "PESO X TORRE"          in hijos else None,
                    "ON_OFF_ALARMA":         hijos["ON/OFF ALARMA"].get_value()              if "ON/OFF ALARMA"         in hijos else None,
                    "TIEMPO_PARA_ALARMA":    hijos["TIEMPO PARA ALARMA"].get_value()         if "TIEMPO PARA ALARMA"    in hijos else None,
                }
        except Exception as e:
            logger.error(f"Error obteniendo recetario OPC: {e}")
        return recetario_cache

    # -----------------------------------------------------------------------
    # Suscripcion / inicializacion (publica)
    # -----------------------------------------------------------------------

    async def iniciar_suscripcion(self, period_ms: int = 500):
        """
        Navega nuevamente el �rbol OPC y crea una suscripci�n nueva.

        Debe llamarse al conectar por primera vez y despu�s de cada
        reconexi�n.
        """

        # No navegar si connect() no termin� correctamente.
        if (
            not self.conexion_servidor.connected
            or self.conexion_servidor.client is None
        ):
            raise ConnectionError(
                "No se puede navegar el �rbol OPC: cliente no conectado"
            )

        # No reutilizar nodos de una conexi�n anterior.
        self._equipos = []
        self._recetario_cache = {}
        self._pf_l1_node = None

        # Evitar valores antiguos despu�s de una reconexi�n.
        self.conexion_servidor.handler.clear()

        # No limpiar self._io_state.
        # Contiene valores y fechas, no objetos Node.
        # Se conserva para no perder tramos IO abiertos.

        def _setup():
            root_node = (
                self.conexion_servidor.client.get_root_node()
            )

            nodos = self._navegar_arbol_sync(root_node)

            self._recetario_cache = (
                self._obtener_recetario_desde_opc_sync(
                    self._pf_l1_node
                )
            )

            return nodos

        try:
            nodos = await asyncio.to_thread(_setup)

            suscripcion_ok = (
                await self.conexion_servidor.suscribir_nodos(
                    nodos,
                    period_ms,
                )
            )

            if not suscripcion_ok:
                raise RuntimeError(
                    "No se pudo crear la suscripci�n OPC UA"
                )

            self._subscription_generation += 1
            self._equipos_pendientes_reconciliacion = {
                self._clave_equipo(
                    equipo["linea"],
                    equipo["tipo"],
                    equipo["numero_local"],
                )
                for equipo in self._equipos
            }
            self._tt_retrocesos.clear()
            self._advertencias_reconciliacion.clear()

        except Exception:
            # Nunca dejar equipos o nodos parcialmente navegados.
            self._equipos = []
            self._recetario_cache = {}
            self._pf_l1_node = None

            # El error debe llegar al monitor de main.py.
            raise

        logger.info(
            "Suscripci�n iniciada: %s nodos, %s recetas cargadas.",
            len(nodos),
            len(self._recetario_cache),
        )

        return True

    async def cargar_recetario(self):
        if not self._pf_l1_node:
            logger.warning("No se puede recargar el recetario: arbol no navegado aun.")
            return
        self._recetario_cache = await asyncio.to_thread(
            self._obtener_recetario_desde_opc_sync, self._pf_l1_node
        )
        logger.info(f"Recetario recargado: {len(self._recetario_cache)} recetas.")

    # -----------------------------------------------------------------------
    # Cache OPC
    # -----------------------------------------------------------------------

    def _leer_datos_equipo(self, equipo: dict) -> dict:
        cache    = self.conexion_servidor.handler.get_all()
        node_map = equipo["node_map"]
        return {campo: cache.get(node_id) for campo, node_id in node_map.items()}

    # -----------------------------------------------------------------------
    # Helpers de negocio generales
    # -----------------------------------------------------------------------

    def obtener_nombre_equipo(self, id_equipo: int) -> str:
        return EQUIPOS_MAP.get(id_equipo, f"Equipo {id_equipo}")

    def _estado_a_texto(self, estado_raw):
        try:
            estado_raw = int(estado_raw)
        except Exception:
            return str(estado_raw).strip().upper()
        return self.ESTADOS_EQUIPO_MAP.get(estado_raw, f"DESCONOCIDO_{estado_raw}")

    @staticmethod
    def _clave_equipo(linea: str, tipo: str, numero_local: int) -> str:
        return f"{linea}-{tipo}-{int(numero_local)}"

    @staticmethod
    def _normalizar_tiempo_trans(valor) -> int | None:
        """
        Convierte TIEMPO_TRANS a minutos enteros.

        En este proyecto el PLC entrega un entero cuya unidad ya es minutos:
        10 significa diez minutos. Para compatibilidad tambien se admiten
        timedelta y cadenas HH:MM:SS; en esos dos casos se conservan solamente
        los minutos completos. No se infiere la unidad por el tamano del valor.
        """
        if valor is None or isinstance(valor, bool):
            return None

        if isinstance(valor, timedelta):
            segundos = valor.total_seconds()
            if not math.isfinite(segundos) or segundos < 0:
                return None
            return int(segundos // 60)

        if isinstance(valor, str):
            texto = valor.strip()
            if not texto:
                return None
            if ":" in texto:
                partes = texto.split(":")
                if len(partes) != 3:
                    return None
                try:
                    horas = int(partes[0])
                    minutos = int(partes[1])
                    segundos = float(partes[2])
                except (TypeError, ValueError):
                    return None
                if (
                    horas < 0
                    or minutos < 0
                    or minutos >= 60
                    or segundos < 0
                    or segundos >= 60
                    or not math.isfinite(segundos)
                ):
                    return None
                return int((horas * 3600 + minutos * 60 + segundos) // 60)
            valor = texto

        try:
            numero = float(valor)
        except (TypeError, ValueError, OverflowError):
            return None
        if not math.isfinite(numero) or numero < 0 or not numero.is_integer():
            return None
        return int(numero)

    def _advertir_reconciliacion_una_vez(
        self,
        equipo_key: str,
        codigo: str,
        mensaje: str,
    ) -> None:
        token = (self._subscription_generation, equipo_key, codigo)
        if token in self._advertencias_reconciliacion:
            return
        self._advertencias_reconciliacion.add(token)
        logger.warning(mensaje)

    def _datos_reconciliacion_frescos(
        self,
        equipo: dict,
        datos: dict,
        equipo_key: str,
    ) -> tuple[str, int] | None:
        """
        Valida la fotografia minima necesaria para operar un ciclo.

        iniciar_suscripcion() limpia el cache antes de crear la nueva
        suscripcion. Por eso, que ambos node_id vuelvan a estar presentes en el
        cache demuestra que ESTADO_EQUIPO y TIEMPO_TRANS fueron recibidos en la
        generacion actual y no quedaron de una conexion anterior.
        """
        if not self.conexion_servidor.connected:
            return None

        node_map = equipo.get("node_map", {})
        cache = self.conexion_servidor.handler.get_all()
        faltantes = []
        for campo in ("ESTADO_EQUIPO", "TIEMPO_TRANS"):
            node_id = node_map.get(campo)
            if not node_id or node_id not in cache:
                faltantes.append(campo)

        if faltantes:
            self._advertir_reconciliacion_una_vez(
                equipo_key,
                "CAMPOS_FRESCOS",
                f"[{equipo_key}] Esperando datos OPC frescos de "
                f"{', '.join(faltantes)} para la suscripcion "
                f"{self._subscription_generation}.",
            )
            return None

        estado_actual = self._estado_a_texto(datos.get("ESTADO_EQUIPO"))
        estados_validos = set(self.ESTADOS_EQUIPO_MAP.values())
        if estado_actual not in estados_validos:
            self._advertir_reconciliacion_una_vez(
                equipo_key,
                f"ESTADO_{estado_actual}",
                f"[{equipo_key}] ESTADO_EQUIPO no permite reconciliar un "
                f"ciclo: {estado_actual!r}.",
            )
            return None

        tt_actual = self._normalizar_tiempo_trans(datos.get("TIEMPO_TRANS"))
        if tt_actual is None:
            self._advertir_reconciliacion_una_vez(
                equipo_key,
                "TT_INVALIDO",
                f"[{equipo_key}] TIEMPO_TRANS invalido. Se conserva cualquier "
                "JSONL pendiente y no se escriben muestras.",
            )
            return None

        return estado_actual, tt_actual

    def _leer_confirmacion_tt_directa(
        self,
        equipo: dict,
        equipo_key: str,
    ) -> tuple[str, int] | None:
        """
        Confirma un posible retroceso con lecturas OPC de red, no repitiendo la
        misma fotografia del cache. Solo se usa para la segunda observacion de
        un TT menor, por lo que no agrega lecturas directas al flujo normal.

        Es sincrona a proposito: se invoca desde _procesar_equipo, que ya se
        ejecuta fuera del event loop.
        """
        if not self.conexion_servidor.connected:
            return None

        node_map = equipo.get("node_map", {})
        estado_node_id = node_map.get("ESTADO_EQUIPO")
        tt_node_id = node_map.get("TIEMPO_TRANS")
        if not estado_node_id or not tt_node_id:
            return None

        try:
            estado_raw = self.conexion_servidor.read_node(estado_node_id)
            tt_raw = self.conexion_servidor.read_node(tt_node_id)
        except Exception as exc:
            self._advertir_reconciliacion_una_vez(
                equipo_key,
                "CONFIRMACION_DIRECTA",
                f"[{equipo_key}] No se pudo confirmar el reinicio de "
                f"TIEMPO_TRANS mediante lectura directa: {exc}. Se mantiene "
                "el JSONL sin cambios.",
            )
            return None

        estado_actual = self._estado_a_texto(estado_raw)
        tt_actual = self._normalizar_tiempo_trans(tt_raw)
        if (
            estado_actual not in set(self.ESTADOS_EQUIPO_MAP.values())
            or tt_actual is None
        ):
            return None
        return estado_actual, tt_actual

    def _slugs_equipo(self, linea, tipo, numero_local):
        linea_slug = "l1" if linea == "PF-L1" else "l2"
        tipo_slug = "cocina" if tipo == "COCINA" else "enfriador"
        return linea_slug, tipo_slug, int(numero_local)

    def _directorio_historial(self, linea, tipo, numero_local):
        linea_slug, tipo_slug, numero = self._slugs_equipo(
            linea, tipo, numero_local
        )
        return os.path.join(
            self._jsonl_dir,
            f"{linea_slug}_{tipo_slug}_{numero}",
        )

    def _archivo_historial(self, linea, tipo, numero_local, id_ciclo):
        directorio = self._directorio_historial(linea, tipo, numero_local)
        return os.path.join(directorio, f"ciclo_{int(id_ciclo)}.jsonl")

    def _archivo_legacy(self, linea, tipo, numero_local):
        linea_slug, tipo_slug, numero = self._slugs_equipo(
            linea, tipo, numero_local
        )
        return f"{tipo_slug}_{numero}_{linea_slug}.json"

    def _informar_archivos_pendientes(self):
        pendientes = []
        en_proceso = []
        for raiz, _, archivos in os.walk(self._jsonl_dir):
            for nombre in archivos:
                ruta = os.path.join(raiz, nombre)
                if nombre.endswith(".jsonl"):
                    pendientes.append(ruta)
                elif nombre.endswith(".processing"):
                    en_proceso.append(ruta)
        if pendientes:
            logger.warning(
                "[JSONL] Se detectaron %s archivos pendientes de ciclos.",
                len(pendientes),
            )
        if en_proceso:
            logger.warning(
                "[JSONL] Se detectaron %s archivos .processing; "
                "se reintentaran al procesar el equipo.",
                len(en_proceso),
            )

    def _validar_registro_jsonl(self, registro, archivo, numero_linea=None):
        contexto = f"Archivo={archivo}"
        if numero_linea is not None:
            contexto += f" Linea={numero_linea}"
        if not isinstance(registro, dict):
            raise ValueError(f"{contexto}: el registro no es un objeto JSON")

        # tiempo_trans_opc es obligatorio para registros nuevos, pero se acepta
        # ausente al leer JSONL creados por versiones anteriores. Esos archivos
        # se conservan y su reconciliacion queda bloqueada de forma segura.
        faltantes = [
            campo
            for campo in CAMPOS_HISTORIAL
            if campo != "tiempo_trans_opc" and campo not in registro
        ]
        if faltantes:
            raise ValueError(f"{contexto}: faltan campos {faltantes}")

        try:
            id_historial = int(registro["id_historial"])
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{contexto}: id_historial invalido") from exc
        if id_historial <= 0:
            raise ValueError(f"{contexto}: id_historial debe ser positivo")

        tiempo = registro.get("tiempo")
        try:
            fecha = _parse_tiempo(tiempo)
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"{contexto}: tiempo debe respetar YYYY-MM-DD HH:MM:SS"
            ) from exc
        if fecha.strftime("%Y-%m-%d %H:%M:%S") != tiempo:
            raise ValueError(f"{contexto}: formato de tiempo no canonico")

        try:
            id_ciclo = int(registro["idCiclo"])
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{contexto}: idCiclo invalido") from exc

        tiempo_trans_opc = self._normalizar_tiempo_trans(
            registro.get("tiempo_trans_opc")
        )
        if (
            "tiempo_trans_opc" in registro
            and registro.get("tiempo_trans_opc") is not None
            and tiempo_trans_opc is None
        ):
            raise ValueError(
                f"{contexto}: tiempo_trans_opc debe ser una cantidad valida "
                "de minutos enteros"
            )

        limpio = {
            "id_historial": id_historial,
            "tiempo": tiempo,
            "estado": str(registro.get("estado") or "").strip().upper(),
            "idCiclo": id_ciclo,
            "lote": str(registro.get("lote") or "").strip(),
            "tiempo_trans_opc": tiempo_trans_opc,
            "temp_agua": _numero_finito_o_none(registro.get("temp_agua")),
            "temp_ingreso": _numero_finito_o_none(
                registro.get("temp_ingreso")
            ),
            "temp_prod": _numero_finito_o_none(registro.get("temp_prod")),
            "niv_agua": _numero_finito_o_none(registro.get("niv_agua")),
        }
        return limpio

    def _iterar_historial_jsonl(
        self,
        archivo,
        tolerar_ultima_incompleta=True,
    ):
        """Lee y valida JSONL sin cargar el archivo completo en memoria."""
        if not archivo or not os.path.exists(archivo):
            return

        ultimo_id = 0
        with open(archivo, "r", encoding="utf-8", newline="") as f:
            for numero_linea, linea_original in enumerate(f, start=1):
                linea = linea_original.strip()
                if not linea:
                    continue
                try:
                    registro = json.loads(linea)
                except json.JSONDecodeError as exc:
                    ultima_incompleta = (
                        tolerar_ultima_incompleta
                        and not linea_original.endswith(("\n", "\r"))
                    )
                    if ultima_incompleta:
                        logger.warning(
                            "[JSONL] Ultima linea incompleta ignorada. "
                            "Archivo=%s Linea=%s",
                            archivo,
                            numero_linea,
                        )
                        break
                    raise ValueError(
                        f"Archivo={archivo} Linea={numero_linea}: JSON invalido"
                    ) from exc

                registro = self._validar_registro_jsonl(
                    registro, archivo, numero_linea
                )
                esperado = ultimo_id + 1
                if registro["id_historial"] != esperado:
                    raise ValueError(
                        f"Archivo={archivo} Linea={numero_linea}: "
                        f"id_historial={registro['id_historial']} "
                        f"pero se esperaba {esperado}"
                    )
                ultimo_id = registro["id_historial"]
                yield registro

    def _reparar_ultima_linea_jsonl(self, archivo):
        """
        Antes de reanudar un archivo activo, separa una ultima linea parcial.
        Los bytes incompletos quedan en un .partial recuperable y las lineas
        validas anteriores permanecen intactas.
        """
        if not archivo.endswith(".jsonl") or not os.path.exists(archivo):
            return
        with open(archivo, "rb+") as f:
            f.seek(0, os.SEEK_END)
            tamano = f.tell()
            if tamano == 0:
                return
            f.seek(-1, os.SEEK_END)
            if f.read(1) in (b"\n", b"\r"):
                return

            posicion = tamano
            acumulado = b""
            inicio_ultima = 0
            while posicion > 0:
                bloque = min(8192, posicion)
                posicion -= bloque
                f.seek(posicion)
                acumulado = f.read(bloque) + acumulado
                indice = acumulado.rfind(b"\n")
                if indice >= 0:
                    inicio_ultima = posicion + indice + 1
                    acumulado = acumulado[indice + 1:]
                    break

            try:
                registro = json.loads(acumulado.decode("utf-8"))
                self._validar_registro_jsonl(registro, archivo)
            except (UnicodeDecodeError, json.JSONDecodeError, ValueError):
                marca = datetime.now().strftime("%Y%m%d_%H%M%S")
                parcial = f"{archivo}.{marca}.partial"
                with open(parcial, "wb") as respaldo:
                    respaldo.write(acumulado)
                    respaldo.flush()
                    if self._jsonl_fsync:
                        os.fsync(respaldo.fileno())
                f.seek(inicio_ultima)
                f.truncate()
                f.flush()
                if self._jsonl_fsync:
                    os.fsync(f.fileno())
                logger.warning(
                    "[JSONL] Ultima linea parcial separada. "
                    "Archivo=%s Respaldo=%s",
                    archivo,
                    parcial,
                )
            else:
                # El objeto era valido y solo faltaba el salto de linea.
                f.seek(0, os.SEEK_END)
                f.write(b"\n")
                f.flush()
                if self._jsonl_fsync:
                    os.fsync(f.fileno())

    def _cargar_estado_jsonl(self, archivo):
        if not archivo:
            return
        lock = self._jsonl_locks[archivo]
        with lock:
            if archivo in self._jsonl_loaded:
                return

            self._reparar_ultima_linea_jsonl(archivo)

            preview = deque(maxlen=HISTORIAL_PREVIEW_MAX)
            primero = None
            ultimo = None
            ultimo_id = 0
            for registro in self._iterar_historial_jsonl(archivo):
                preview.append(registro)
                primero = primero or registro["tiempo"]
                ultimo = registro["tiempo"]
                ultimo_id = registro["id_historial"]
                self._jsonl_last_record[archivo] = registro

            self._jsonl_preview[archivo] = preview
            self._jsonl_next_id[archivo] = ultimo_id + 1
            if primero:
                self._jsonl_first_time[archivo] = primero
            if ultimo:
                self._jsonl_last_time[archivo] = ultimo
            self._jsonl_loaded.add(archivo)

    def _obtener_preview_jsonl(self, archivo):
        if not archivo:
            return []
        self._cargar_estado_jsonl(archivo)
        return list(self._jsonl_preview.get(archivo, ()))

    def _ultimo_registro_del_archivo(self, archivo):
        if not archivo:
            return None
        self._cargar_estado_jsonl(archivo)
        ultimo = self._jsonl_last_record.get(archivo)
        return dict(ultimo) if ultimo else None

    def _agregar_registro_jsonl(self, archivo, registro):
        if not archivo.endswith(".jsonl"):
            raise ValueError(f"No se puede agregar sobre {archivo}")

        if self._normalizar_tiempo_trans(
            registro.get("tiempo_trans_opc")
        ) is None:
            raise ValueError(
                "No se puede agregar un registro JSONL sin "
                "tiempo_trans_opc valido"
            )

        lock = self._jsonl_locks[archivo]
        with lock:
            base = archivo[:-len(".jsonl")]
            if os.path.exists(base + ".processing") or os.path.exists(
                base + ".done"
            ):
                raise RuntimeError(
                    f"El ciclo ya se esta cerrando o fue archivado: {archivo}"
                )
            self._cargar_estado_jsonl(archivo)
            id_historial = self._jsonl_next_id.get(archivo, 1)
            completo = dict(registro)
            completo["id_historial"] = id_historial
            completo = self._validar_registro_jsonl(completo, archivo)

            linea = json.dumps(
                completo,
                ensure_ascii=False,
                separators=(",", ":"),
                allow_nan=False,
                default=datetime_to_string,
            )

            os.makedirs(os.path.dirname(archivo), exist_ok=True)
            with open(archivo, "a", encoding="utf-8", newline="\n") as f:
                f.write(linea)
                f.write("\n")
                f.flush()
                if self._jsonl_fsync:
                    os.fsync(f.fileno())

            preview = self._jsonl_preview.setdefault(
                archivo, deque(maxlen=HISTORIAL_PREVIEW_MAX)
            )
            preview.append(completo)
            self._jsonl_next_id[archivo] = id_historial + 1
            self._jsonl_first_time.setdefault(archivo, completo["tiempo"])
            self._jsonl_last_time[archivo] = completo["tiempo"]
            self._jsonl_last_record[archivo] = completo
            self._jsonl_loaded.add(archivo)
            logger.debug(
                "[JSONL] Archivo=%s id_historial=%s agregado",
                archivo,
                id_historial,
            )
            return completo

    def _olvidar_estado_jsonl(self, *archivos):
        with self._jsonl_meta_lock:
            for archivo in archivos:
                if not archivo:
                    continue
                self._jsonl_loaded.discard(archivo)
                self._jsonl_next_id.pop(archivo, None)
                self._jsonl_preview.pop(archivo, None)
                self._jsonl_first_time.pop(archivo, None)
                self._jsonl_last_time.pop(archivo, None)
                self._jsonl_last_record.pop(archivo, None)
                self._jsonl_locks.pop(archivo, None)

    def _renombrar_a_processing(self, archivo):
        if archivo.endswith(".processing"):
            return archivo
        if not archivo.endswith(".jsonl"):
            raise ValueError(f"Extension de historial inesperada: {archivo}")

        destino = archivo[:-len(".jsonl")] + ".processing"
        lock = self._jsonl_locks[archivo]
        with lock:
            if not os.path.exists(archivo):
                if os.path.exists(destino):
                    return destino
                raise FileNotFoundError(archivo)
            if os.path.exists(destino):
                raise FileExistsError(
                    f"Ya existe un procesamiento pendiente: {destino}"
                )
            os.replace(archivo, destino)
        self._olvidar_estado_jsonl(archivo, destino)
        return destino

    def _archivar_jsonl(self, archivo_processing):
        base = (
            archivo_processing[:-len(".processing")]
            if archivo_processing.endswith(".processing")
            else os.path.splitext(archivo_processing)[0]
        )
        destino = base + ".done"
        if os.path.exists(destino):
            marca = datetime.now().strftime("%Y%m%d_%H%M%S")
            destino = f"{base}_{marca}.done"
        os.replace(archivo_processing, destino)
        self._olvidar_estado_jsonl(archivo_processing, destino)
        logger.info("[JSONL] Archivo procesado archivado: %s", destino)
        return destino

    def _migrar_json_legacy(self, linea, tipo, numero_local):
        legacy = self._archivo_legacy(linea, tipo, numero_local)
        if not os.path.exists(legacy) or os.path.getsize(legacy) == 0:
            return
        try:
            with open(legacy, "r", encoding="utf-8") as f:
                historial = json.load(f)
            if not isinstance(historial, list):
                raise ValueError("el JSON anterior no contiene un arreglo")
            if not historial:
                destino_vacio = legacy + ".migrated"
                if os.path.exists(destino_vacio):
                    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
                    destino_vacio = f"{destino_vacio}_{marca}"
                if not os.path.exists(destino_vacio):
                    os.replace(legacy, destino_vacio)
                return

            ciclos_legacy = {int(fila["idCiclo"]) for fila in historial}
            if len(ciclos_legacy) != 1:
                raise ValueError(
                    f"el JSON anterior mezcla ciclos: {sorted(ciclos_legacy)}"
                )
            id_ciclo = ciclos_legacy.pop()
            destino = self._archivo_historial(
                linea, tipo, numero_local, id_ciclo
            )
            if os.path.exists(destino):
                logger.warning(
                    "[JSONL] No se migra %s porque ya existe %s",
                    legacy,
                    destino,
                )
                return

            os.makedirs(os.path.dirname(destino), exist_ok=True)
            temporal = destino + ".migrating"
            esperado = 1
            with open(temporal, "w", encoding="utf-8", newline="\n") as f:
                for numero_linea, fila in enumerate(historial, start=1):
                    limpio = self._validar_registro_jsonl(
                        fila, legacy, numero_linea
                    )
                    if limpio["id_historial"] != esperado:
                        raise ValueError(
                            f"id_historial discontinuo en {legacy}: "
                            f"esperado={esperado}"
                        )
                    esperado += 1
                    f.write(json.dumps(
                        limpio,
                        ensure_ascii=False,
                        separators=(",", ":"),
                        allow_nan=False,
                    ))
                    f.write("\n")
                f.flush()
                if self._jsonl_fsync:
                    os.fsync(f.fileno())
            os.replace(temporal, destino)
            respaldo_legacy = legacy + ".migrated"
            if os.path.exists(respaldo_legacy):
                marca = datetime.now().strftime("%Y%m%d_%H%M%S")
                respaldo_legacy = f"{respaldo_legacy}_{marca}"
            os.replace(legacy, respaldo_legacy)
            logger.warning(
                "[JSONL] Historial anterior migrado: %s -> %s",
                legacy,
                destino,
            )
        except Exception as exc:
            logger.error(
                "[JSONL] No se pudo migrar historial anterior %s: %s",
                legacy,
                exc,
            )
            raise

    def _buscar_historial_pendiente(self, linea, tipo, numero_local):
        self._migrar_json_legacy(linea, tipo, numero_local)
        directorio = self._directorio_historial(linea, tipo, numero_local)
        if not os.path.isdir(directorio):
            return None

        nombres = os.listdir(directorio)
        processing = [
            os.path.join(directorio, n)
            for n in nombres
            if n.endswith(".processing")
        ]
        activos = [
            os.path.join(directorio, n)
            for n in nombres
            if n.endswith(".jsonl")
        ]
        candidatos = processing or activos
        if not candidatos:
            return None
        candidatos.sort(key=os.path.getmtime, reverse=True)
        if len(candidatos) > 1 or (processing and activos):
            logger.error(
                "[JSONL] Hay varios historiales pendientes para %s/%s/%s: %s",
                linea,
                tipo,
                numero_local,
                processing + activos,
            )
            raise RuntimeError("Mas de un historial pendiente para el equipo")
        return candidatos[0]

    def calcular_tiempo_transcurrido_jsonl(self, archivo):
        if not archivo:
            return "00:00:00"
        try:
            self._cargar_estado_jsonl(archivo)
            inicio = self._jsonl_first_time.get(archivo)
            fin = self._jsonl_last_time.get(archivo)
            if not inicio or not fin:
                return "00:00:00"
            return _delta_str(_parse_tiempo(inicio), _parse_tiempo(fin))
        except Exception as exc:
            logger.error(
                "[JSONL] Error calculando tiempo de %s: %s", archivo, exc
            )
            return "00:00:00"

    @staticmethod
    def _historial_para_websocket(historial):
        presentacion = []
        for registro in historial:
            fila = dict(registro)
            # Campo interno de reconciliacion: no cambia el contrato existente
            # del historial enviado al frontend.
            fila.pop("tiempo_trans_opc", None)
            for campo in CAMPOS_TEMPERATURA:
                fila[campo] = _redondear_para_presentacion(fila.get(campo))
            presentacion.append(fila)
        return presentacion

    # -----------------------------------------------------------------------
    # Reconciliacion del historial por TIEMPO_TRANS
    # -----------------------------------------------------------------------

    # LOTE_CICLO se conserva como dato informativo. La identidad y continuidad
    # de los ciclos se resuelve exclusivamente en _reconciliar_historial_por_tt.

    def _reconciliar_historial_por_tt(
        self,
        session: Session,
        archivo_historial: str | None,
        tt_actual: int,
        estado_actual: str,
        equipo_key: str,
        linea: str,
        tipo: str,
        numero_local: int,
    ) -> dict:
        """
        Decide si el JSONL pendiente y la fotografia OPC pertenecen al mismo
        ciclo. LOTE_CICLO no participa de ninguna decision.

        CONFIRMAR_TT e INDETERMINADA bloquean escrituras para el equipo, pero
        conservan el archivo sin modificarlo.

        Limitacion del dato disponible: si durante una desconexion empieza otro
        ciclo y al volver su TT ya es mayor que el ultimo TT anterior, ambos son
        indistinguibles sin un ID_CICLO_PLC o contador monotonicamente creciente.
        """
        if not archivo_historial:
            self._tt_retrocesos.pop(equipo_key, None)
            self._equipos_pendientes_reconciliacion.discard(equipo_key)
            return {
                "decision": self.RECON_SIN_HISTORIAL,
                "id_ciclo": None,
                "archivo": None,
                "historial": [],
            }

        historial = self._obtener_preview_jsonl(archivo_historial)
        ultimo = self._ultimo_registro_del_archivo(archivo_historial)
        if not ultimo:
            self._advertir_reconciliacion_una_vez(
                equipo_key,
                "JSONL_VACIO",
                f"[{linea}/{tipo}/{numero_local}] El JSONL pendiente esta "
                "vacio. Se conserva y el equipo queda bloqueado.",
            )
            return {
                "decision": self.RECON_INDETERMINADA,
                "id_ciclo": None,
                "archivo": archivo_historial,
                "historial": historial,
            }

        id_ciclo_anterior = int(ultimo["idCiclo"])
        tt_anterior = self._normalizar_tiempo_trans(
            ultimo.get("tiempo_trans_opc")
        )
        if tt_anterior is None:
            self._advertir_reconciliacion_una_vez(
                equipo_key,
                "JSONL_SIN_TT",
                f"[{linea}/{tipo}/{numero_local}] El JSONL del ciclo "
                f"{id_ciclo_anterior} fue creado sin tiempo_trans_opc. "
                "No se elimina, no se mezcla con lecturas nuevas y requiere "
                "una decision manual.",
            )
            return {
                "decision": self.RECON_INDETERMINADA,
                "id_ciclo": id_ciclo_anterior,
                "archivo": archivo_historial,
                "historial": historial,
            }

        fecha_ultima_captura = _parse_tiempo(ultimo["tiempo"])

        # Un .processing representa un cierre que ya habia comenzado. Nunca se
        # vuelve a escribir sobre el; se reintenta el cierre idempotente.
        if archivo_historial.endswith(".processing"):
            mismo_ciclo_terminal = (
                estado_actual in self.ESTADOS_FIN
                and tt_actual >= tt_anterior
            )
            estado_cierre = (
                estado_actual if mismo_ciclo_terminal else "CANCELADO"
            )
            fecha_forzada = (
                None if mismo_ciclo_terminal else fecha_ultima_captura
            )
            ok = self.finalizar_ciclo_completo(
                id_ciclo=id_ciclo_anterior,
                estado_maquina=estado_cierre,
                archivo_historial=archivo_historial,
                equipo_key=equipo_key,
                tipo=tipo,
                fecha_fin_forzada=fecha_forzada,
                session=session,
            )
            if not ok:
                raise RuntimeError(
                    f"No se pudo recuperar {archivo_historial}; no se crean "
                    "ni se mezclan ciclos nuevos"
                )
            self._tt_retrocesos.pop(equipo_key, None)
            self._equipos_pendientes_reconciliacion.discard(equipo_key)
            return {
                "decision": self.RECON_NUEVO_CICLO,
                "id_ciclo": None,
                "archivo": None,
                "historial": [],
            }

        if tt_actual >= tt_anterior:
            self._tt_retrocesos.pop(equipo_key, None)
            self._equipos_pendientes_reconciliacion.discard(equipo_key)
            return {
                "decision": self.RECON_MISMO_CICLO,
                "id_ciclo": id_ciclo_anterior,
                "archivo": archivo_historial,
                "historial": historial,
            }

        # TT menor: exigir dos observaciones consecutivas de la misma
        # generacion de suscripcion antes de cerrar el historial anterior.
        firma = (
            self._subscription_generation,
            os.path.abspath(archivo_historial),
            id_ciclo_anterior,
            tt_anterior,
        )
        retroceso = self._tt_retrocesos.get(equipo_key)
        if not retroceso or retroceso.get("firma") != firma:
            retroceso = {
                "firma": firma,
                "confirmaciones": 1,
                "primer_tt_actual": tt_actual,
            }
        else:
            retroceso["confirmaciones"] += 1
        retroceso["ultimo_tt_actual"] = tt_actual
        self._tt_retrocesos[equipo_key] = retroceso

        if retroceso["confirmaciones"] < 2:
            self._advertir_reconciliacion_una_vez(
                equipo_key,
                f"RETROCESO_{id_ciclo_anterior}_{tt_anterior}",
                f"[{linea}/{tipo}/{numero_local}] Posible reinicio de "
                f"TIEMPO_TRANS: JSONL={tt_anterior} min, OPC={tt_actual} min. "
                "Se espera una segunda observacion antes de cortar el ciclo.",
            )
            return {
                "decision": self.RECON_CONFIRMAR_TT,
                "id_ciclo": id_ciclo_anterior,
                "archivo": archivo_historial,
                "historial": historial,
            }

        logger.warning(
            "[%s/%s/%s] Reinicio de TIEMPO_TRANS confirmado: "
            "JSONL=%s min, OPC=%s min. El ciclo %s se cierra como CANCELADO "
            "en su ultima captura conocida (%s).",
            linea,
            tipo,
            numero_local,
            tt_anterior,
            tt_actual,
            id_ciclo_anterior,
            ultimo["tiempo"],
        )
        ok = self.finalizar_ciclo_completo(
            id_ciclo=id_ciclo_anterior,
            estado_maquina="CANCELADO",
            archivo_historial=archivo_historial,
            equipo_key=equipo_key,
            tipo=tipo,
            fecha_fin_forzada=fecha_ultima_captura,
            session=session,
        )
        if not ok:
            raise RuntimeError(
                f"No se pudo cerrar el ciclo {id_ciclo_anterior}; el JSONL "
                "queda .processing y el equipo permanece bloqueado"
            )

        self._tt_retrocesos.pop(equipo_key, None)
        self._equipos_pendientes_reconciliacion.discard(equipo_key)
        return {
            "decision": self.RECON_NUEVO_CICLO,
            "id_ciclo": None,
            "archivo": None,
            "historial": [],
        }

    # -----------------------------------------------------------------------
    # BD \u2014 ciclos
    # -----------------------------------------------------------------------

    def _obtener_ciclo_activo(self, session: Session, id_equipo):
        try:
            return (
                session.query(Ciclo)
                .filter(Ciclo.idEquipo == id_equipo, Ciclo.fecha_fin.is_(None))
                .order_by(Ciclo.id.desc())
                .first()
            )
        except Exception as e:
            logger.error(f"Error buscando ciclo activo para equipo {id_equipo}: {e}")
            return None

    def _cerrar_ciclo_bd(self, session: Session, ciclo: Ciclo, estado_maquina: str):
        """
        Cierre m�nimo de un ciclo directamente en BD, sin historial JSON.
        Se conserva para compatibilidad con cierres que no tienen JSONL.
        El lote no participa de esta decision.
        """
        try:
            if ciclo.fecha_fin is not None:
                return  # Ya cerrado, nada que hacer
            fecha_fin = datetime.now().replace(microsecond=0)
            ciclo.fecha_fin          = fecha_fin
            ciclo.estadoMaquina      = estado_maquina
            ciclo.cantidadPausas     = 0
            ciclo.tiempoTranscurrido = _delta_str(ciclo.fecha_inicio, fecha_fin)
            session.commit()
            logger.warning(
                f"Ciclo BD {ciclo.id} cerrado como '{estado_maquina}' "
                "sin historial JSONL."
            )
        except Exception as e:
            session.rollback()
            logger.error(f"Error en _cerrar_ciclo_bd (ciclo={ciclo.id}): {e}")

    def _resolver_id_ciclo(
        self,
        session: Session,
        decision_reconciliacion: str,
        id_ciclo_reconciliado: int | None,
        id_equipo: int,
        lote_ciclo: str,
        receta_id: int | None,
        datos_equipo: dict,
        estado_actual: str,
    ) -> int:
        """
        Resuelve un id real de MySQL sin utilizar el lote como identidad.

        - MISMO_CICLO: verifica y reutiliza exclusivamente el id del JSONL.
        - NUEVO_CICLO: exige que el anterior ya este cerrado y crea otro.
        - SIN_HISTORIAL: reutiliza el unico ciclo abierto del equipo solamente
          para recuperar la ventana entre el commit de creacion y la primera
          escritura JSONL; si no existe, crea uno.

        Nunca genera ids locales, aleatorios ni hashes.
        """
        if decision_reconciliacion == self.RECON_MISMO_CICLO:
            if id_ciclo_reconciliado is None:
                raise RuntimeError(
                    "La reconciliacion indico MISMO_CICLO sin idCiclo"
                )
            ciclo = session.query(Ciclo).filter_by(
                id=int(id_ciclo_reconciliado)
            ).first()
            if not ciclo:
                raise RuntimeError(
                    f"El idCiclo {id_ciclo_reconciliado} del JSONL no existe "
                    "en MySQL"
                )
            if int(ciclo.idEquipo) != int(id_equipo):
                raise RuntimeError(
                    f"El ciclo {ciclo.id} pertenece al equipo "
                    f"{ciclo.idEquipo}, no al {id_equipo}"
                )
            if ciclo.fecha_fin is not None:
                raise RuntimeError(
                    f"El ciclo {ciclo.id} del JSONL ya esta cerrado en MySQL"
                )
            return int(ciclo.id)

        if decision_reconciliacion not in {
            self.RECON_SIN_HISTORIAL,
            self.RECON_NUEVO_CICLO,
        }:
            raise RuntimeError(
                "No se puede resolver un ciclo mientras la reconciliacion "
                f"esta en {decision_reconciliacion}"
            )

        ciclo_activo = self._obtener_ciclo_activo(session, id_equipo)
        if decision_reconciliacion == self.RECON_SIN_HISTORIAL and ciclo_activo:
            logger.warning(
                "Equipo %s sin JSONL, pero con ciclo MySQL abierto %s. "
                "Se reutiliza por idEquipo para recuperar una posible "
                "interrupcion entre el commit y la primera escritura.",
                id_equipo,
                ciclo_activo.id,
            )
            return int(ciclo_activo.id)

        if decision_reconciliacion == self.RECON_NUEVO_CICLO and ciclo_activo:
            raise RuntimeError(
                f"No se crea un ciclo nuevo: el equipo {id_equipo} aun tiene "
                f"abierto el ciclo MySQL {ciclo_activo.id}"
            )

        if receta_id is None:
            raise RuntimeError(
                "No se crea el ciclo porque la receta no pudo confirmarse "
                "en MySQL"
            )

        id_ciclo = self.guardarEnBaseCiclo(
            {
                "estadoMaquina": estado_actual,
                "cantidadTorres": int(datos_equipo.get("CANTIDAD_TORRES") or 0),
                "lote": lote_ciclo or "",
                "fecha_inicio": datetime.now().replace(microsecond=0),
                "peso": int(datos_equipo.get("PESO_PRODUCTO") or 0),
                "idEquipo": id_equipo,
                "idReceta": receta_id,
            },
            session=session,
        )
        if id_ciclo is None:
            raise RuntimeError(
                "MySQL no confirmo la creacion del ciclo; no se escribira "
                "ningun JSONL con un id inexistente"
            )
        return int(id_ciclo)

    def obtener_o_crear_receta(self, receta_opc, numero_receta=None, session=None):
        # Compatibilidad: si no se recibe una sesion se abre y se cierra una
        # propia, de modo que el metodo sigue siendo invocable desde afuera.
        if session is None:
            with self._sesion_bd() as propia:
                return self.obtener_o_crear_receta(
                    receta_opc, numero_receta, propia
                )

        try:
            nombre_receta = receta_opc.get("NOMBRE") if receta_opc else None
            if not nombre_receta:
                nombre_receta = f"RECETA_{int(numero_receta or 0):02}"
            nro_paso = receta_opc.get("PASOS", 0) if receta_opc else 0
            tipo_fin = receta_opc.get("TIPO CORTE", False) if receta_opc else False

            receta = session.query(Receta).filter(Receta.nombre == nombre_receta).first()
            if receta:
                receta.nroPaso = nro_paso
                receta.tipoFin = tipo_fin

                session.commit()
                return receta.id

            ultimo_id = session.query(Receta).order_by(Receta.id.desc()).first()
            nuevo_id  = 1 if not ultimo_id else ultimo_id.id + 1
            session.add(Receta(id=nuevo_id, nombre=nombre_receta,
                               nroPaso=nro_paso, tipoFin=tipo_fin))
            session.commit()
            logger.info(f"Nueva receta creada - ID: {nuevo_id}, Nombre: {nombre_receta}")
            return nuevo_id
        except Exception as e:
            session.rollback()
            logger.error(f"Error al obtener/crear receta: {e}")
            return None

    def guardarEnBaseCiclo(self, datos, session=None):
        if session is None:
            with self._sesion_bd() as propia:
                return self.guardarEnBaseCiclo(datos, session=propia)

        try:
            nuevo_ciclo = Ciclo(
                estadoMaquina  = datos["estadoMaquina"],
                cantidadTorres = datos["cantidadTorres"],
                lote           = datos["lote"],
                fecha_inicio   = datos["fecha_inicio"],
                peso           = datos["peso"],
                idEquipo       = datos["idEquipo"],
                idReceta       = datos["idReceta"],
            )
            session.add(nuevo_ciclo)
            session.commit()
            logger.info(
                f"\033[1;93m[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] "
                f"{self.obtener_nombre_equipo(datos['idEquipo'])} "
                f"NUEVO CICLO: {nuevo_ciclo.id}\033[0m"
            )
            return nuevo_ciclo.id
        except Exception as e:
            session.rollback()
            logger.error(f"Error al guardar nuevo ciclo: {e}")
            return None

    def obtener_ultimo_ciclo_finalizado(self, id_equipo, session=None):
        if session is None:
            with self._sesion_bd() as propia:
                return self.obtener_ultimo_ciclo_finalizado(
                    id_equipo, session=propia
                )

        try:
            ultimo = (
                session.query(Ciclo)
                .filter(
                    Ciclo.idEquipo == id_equipo,
                    Ciclo.estadoMaquina.in_(["FINALIZADO", "CANCELADO", "INACTIVO"]),
                    Ciclo.fecha_fin.isnot(None),
                )
                .order_by(Ciclo.fecha_fin.desc())
                .first()
            )
            return ultimo.id if ultimo else None
        except Exception:
            return None

    def obtener_id_ciclo_existente(self, lote, idEquipo, session=None):
        """Compatibilidad: devuelve un ciclo abierto real, nunca inventa ids."""
        if session is None:
            with self._sesion_bd() as propia:
                return self.obtener_id_ciclo_existente(
                    lote, idEquipo, session=propia
                )

        try:
            ciclo = (
                session.query(Ciclo)
                .filter_by(idEquipo=idEquipo, fecha_fin=None)
                .order_by(Ciclo.id.desc())
                .first()
            )
            return ciclo.id if ciclo else None
        except Exception as e:
            logger.error(f"Error consultando ciclo abierto real: {e}")
            return None

    # -----------------------------------------------------------------------
    # SensoresIO \u2014 trazabilidad en tiempo real
    # -----------------------------------------------------------------------

    def _io_map_para_tipo(self, tipo: str) -> dict[str, int]:
        return IO_SENSOR_MAP_COCINA if tipo == "COCINA" else IO_SENSOR_MAP_ENFRIADOR

    def _persistir_tramo_io(self, session: Session, id_ciclo: int,
                             id_sensor: int, valor: bool,
                             fecha_inicio: datetime, fecha_fin: datetime,
                             confirmar: bool = True) -> bool:
        """
        Inserta un tramo IO en BD.
        Idempotente: la unicidad se verifica por (idCiclo, idSensor, fechaInicio).
        """
        try:
            existe = session.query(SensoresIO).filter_by(
                idCiclo=id_ciclo,
                idSensor=id_sensor,
                fechaInicio=fecha_inicio,
            ).first()
            if existe:
                return True
            session.add(SensoresIO(
                idSensor    = id_sensor,
                valor       = valor,
                fechaInicio = fecha_inicio,
                fechaFin    = fecha_fin,
                idCiclo     = id_ciclo,
            ))
            if confirmar:
                session.commit()
            else:
                session.flush()
            return True
        except Exception as e:
            session.rollback()
            logger.error(
                f"Error persistiendo tramo IO "
                f"(ciclo={id_ciclo}, sensor={id_sensor}): {e}"
            )
            if not confirmar:
                raise
            return False

    def _actualizar_io_state(self, session: Session, equipo_key: str,
                              tipo: str, datos: dict,
                              id_ciclo: int, ahora: datetime):
        """
        Llamado cada segundo mientras el ciclo esta activo.

        Logica por cada senal booleana del equipo:
        - Primera vez visto  -> abrir tramo en memoria.
        - Mismo ciclo, valor igual    -> no hacer nada.
        - Mismo ciclo, valor cambio   -> persistir tramo cerrado, abrir nuevo.
        - Ciclo diferente (equipo reutilizado) -> cerrar tramos del ciclo
          anterior y abrir nuevos para el nuevo ciclo.
        """
        io_map = self._io_map_para_tipo(tipo)

        if equipo_key not in self._io_state:
            self._io_state[equipo_key] = {}

        estado_equipo = self._io_state[equipo_key]

        for campo, id_sensor in io_map.items():
            valor_raw = datos.get(campo)
            if valor_raw is None:
                # Cache vacio (reconexion en curso): no actualizar estado
                continue
            valor = bool(valor_raw)

            if campo not in estado_equipo:
                # Primera muestra de esta senal
                estado_equipo[campo] = {
                    "valor":    valor,
                    "inicio":   ahora,
                    "id_ciclo": id_ciclo,
                }
            else:
                tramo = estado_equipo[campo]

                if tramo["id_ciclo"] != id_ciclo:
                    # Ciclo nuevo: cerrar tramo del ciclo anterior inmediatamente
                    persistido = self._persistir_tramo_io(
                        session, tramo["id_ciclo"], id_sensor,
                        tramo["valor"], tramo["inicio"], ahora,
                    )
                    if not persistido:
                        continue
                    estado_equipo[campo] = {
                        "valor":    valor,
                        "inicio":   ahora,
                        "id_ciclo": id_ciclo,
                    }
                elif tramo["valor"] != valor:
                    # Valor cambio: cerrar tramo actual y abrir nuevo
                    persistido = self._persistir_tramo_io(
                        session, id_ciclo, id_sensor,
                        tramo["valor"], tramo["inicio"], ahora,
                    )
                    if not persistido:
                        continue
                    estado_equipo[campo] = {
                        "valor":    valor,
                        "inicio":   ahora,
                        "id_ciclo": id_ciclo,
                    }
                # Valor igual y mismo ciclo: no hacer nada

    def _cerrar_io_tramos_equipo(self, session: Session, equipo_key: str,
                                  tipo: str,
                                  id_ciclo: int, fecha_fin: datetime,
                                  limpiar_estado: bool = True):
        """
        Al finalizar el ciclo: persiste todos los tramos IO abiertos en
        memoria para este equipo y limpia el estado.
        Los tramos de ciclos anteriores (ya persistidos) se ignoran.
        """
        if equipo_key not in self._io_state:
            return

        io_map        = self._io_map_para_tipo(tipo)
        estado_equipo = self._io_state[equipo_key]

        for campo, id_sensor in io_map.items():
            if campo not in estado_equipo:
                continue
            tramo = estado_equipo[campo]
            if tramo["id_ciclo"] != id_ciclo:
                continue  # Tramo de un ciclo anterior, ya fue persistido
            self._persistir_tramo_io(
                session, id_ciclo, id_sensor,
                tramo["valor"], tramo["inicio"], fecha_fin,
                confirmar=False,
            )

        if limpiar_estado:
            self._io_state.pop(equipo_key, None)

    # -----------------------------------------------------------------------
    # SensoresAA \u2014 persistencia historica al cierre del ciclo
    # -----------------------------------------------------------------------

    @staticmethod
    def _orden_historial(fila):
        return (
            fila.get("tiempo", ""),
            int(fila.get("id_historial") or 0),
        )

    def _agrupar_temperaturas_por_segundo(self, historial):
        """
        Reduce las capturas que comparten segundo a una fila representativa.
        Las tres temperaturas se promedian; NIVEL_AGUA conserva el ultimo
        valor valido. La fila mantiene los cuatro sensores juntos.
        """
        grupos = defaultdict(list)
        for fila in sorted(historial, key=self._orden_historial):
            grupos[fila["tiempo"]].append(fila)

        resultado = []
        for tiempo in sorted(grupos, key=_parse_tiempo):
            filas = sorted(
                grupos[tiempo],
                key=lambda r: int(r.get("id_historial") or 0),
            )
            representativa = dict(filas[-1])

            for campo in CAMPOS_TEMPERATURA:
                valores = [
                    _numero_finito_o_none(fila.get(campo))
                    for fila in filas
                ]
                valores = [
                    float(valor)
                    for valor in valores
                    if valor is not None and not isinstance(valor, bool)
                ]
                representativa[campo] = mean(valores) if valores else None

            niveles = [
                _numero_finito_o_none(fila.get("niv_agua"))
                for fila in reversed(filas)
            ]
            representativa["niv_agua"] = next(
                (valor for valor in niveles if valor is not None),
                None,
            )
            resultado.append(representativa)
        return resultado

    def _filtrar_historial_para_bd(self, historial):
        """
        Filtra al finalizar el ciclo.

        - Temperaturas: promedio por segundo + ventana movil de 10 segundos,
          umbral de 0.5, tres confirmaciones y muestra maxima cada 300 s.
        - NIVEL_AGUA: conserva una fila al primer valor y en cada cambio.
        - La seleccion final es la union de ambos criterios y cada fila siempre
          conserva temp_agua, temp_ingreso, temp_prod y niv_agua.
        """
        if not historial:
            return []

        crudo = sorted(historial, key=self._orden_historial)
        por_segundo = self._agrupar_temperaturas_por_segundo(crudo)
        resumen_por_tiempo = {fila["tiempo"]: fila for fila in por_segundo}
        seleccionados: dict[int, dict] = {}

        # NIVEL_AGUA se evalua sobre cada captura para no perder dos cambios
        # ocurridos dentro del mismo segundo.
        ultimo_nivel = None
        nivel_inicializado = False
        for fila in crudo:
            nivel = _numero_finito_o_none(fila.get("niv_agua"))
            if nivel is None:
                continue
            if not nivel_inicializado or nivel != ultimo_nivel:
                candidata = dict(fila)
                promedio_segundo = resumen_por_tiempo.get(fila["tiempo"], {})
                for campo in CAMPOS_TEMPERATURA:
                    candidata[campo] = promedio_segundo.get(
                        campo, candidata.get(campo)
                    )
                seleccionados[candidata["id_historial"]] = candidata
                ultimo_nivel = nivel
                nivel_inicializado = True

        ventanas = {
            campo: deque(maxlen=VENTANA_PROMEDIO_SEGUNDOS)
            for campo in CAMPOS_TEMPERATURA
        }
        referencias = {campo: None for campo in CAMPOS_TEMPERATURA}
        persistencias = {campo: 0 for campo in CAMPOS_TEMPERATURA}
        ultimas_fechas = {campo: None for campo in CAMPOS_TEMPERATURA}

        for fila in por_segundo:
            fecha = _parse_tiempo(fila["tiempo"])
            promedios_moviles = {}
            sensores_disparados = []

            for campo in CAMPOS_TEMPERATURA:
                valor = _numero_finito_o_none(fila.get(campo))
                if valor is None or isinstance(valor, bool):
                    persistencias[campo] = 0
                    continue

                ventanas[campo].append(float(valor))
                if len(ventanas[campo]) < VENTANA_PROMEDIO_SEGUNDOS:
                    continue

                promedio_actual = mean(ventanas[campo])
                promedios_moviles[campo] = promedio_actual
                referencia = referencias[campo]

                guardar = False
                if referencia is None:
                    guardar = True
                elif (
                    ultimas_fechas[campo] is not None
                    and (fecha - ultimas_fechas[campo]).total_seconds()
                    >= TIEMPO_MAXIMO_SEGUNDOS
                ):
                    guardar = True
                elif abs(promedio_actual - referencia) >= UMBRAL_CAMBIO:
                    persistencias[campo] += 1
                    guardar = (
                        persistencias[campo] >= PERSISTENCIA_REQUERIDA
                    )
                else:
                    persistencias[campo] = 0

                if guardar:
                    sensores_disparados.append(campo)
                    referencias[campo] = promedio_actual
                    ultimas_fechas[campo] = fecha
                    persistencias[campo] = 0

            if sensores_disparados:
                candidata = dict(fila)
                # Se guardan los promedios moviles disponibles de las tres
                # temperaturas, aunque el disparo sea de una sola.
                for campo, promedio in promedios_moviles.items():
                    candidata[campo] = promedio
                existente = seleccionados.get(candidata["id_historial"])
                if existente:
                    existente.update({
                        campo: candidata[campo]
                        for campo in CAMPOS_TEMPERATURA
                    })
                else:
                    seleccionados[candidata["id_historial"]] = candidata

        # Un ciclo corto o sin NIVEL_AGUA no debe perder todos sus datos.
        if not seleccionados and por_segundo:
            seleccionados[por_segundo[-1]["id_historial"]] = dict(
                por_segundo[-1]
            )

        resultado = sorted(
            seleccionados.values(),
            key=self._orden_historial,
        )
        logger.info(
            "[PROCESAMIENTO] Registros_crudos=%s Registros_filtrados=%s",
            len(historial),
            len(resultado),
        )
        return resultado

    def _persistir_sensores_aa(self, session: Session, id_ciclo: int,
                                historial: list) -> bool:
        """
        Inserta en SensoresAA las muestras filtradas obtenidas del JSONL.
        Idempotente: si ya existen filas para este id_ciclo, no inserta nada.
        Usa flush (sin commit) para participar en la transaccion del cierre.

        Returns True si OK, False si hubo error.
        """
        try:
            ya_existen = session.query(SensoresAA).filter_by(
                idCiclo=id_ciclo
            ).first()
            if ya_existen:
                logger.info(
                    f"SensoresAA para ciclo {id_ciclo} ya existen. Se omite."
                )
                return True

            if not historial:
                return True

            ordenado = sorted(historial, key=self._orden_historial)
            objetos  = []

            for fila in ordenado:
                fecha = _parse_tiempo(fila["tiempo"])

                for campo_json, id_sensor in AA_SENSOR_MAP.items():
                    valor = fila.get(campo_json)
                    if valor is None:
                        continue
                    try:
                        valor_float = float(valor)
                    except (TypeError, ValueError):
                        continue

                    if not math.isfinite(valor_float):
                        continue
                    objetos.append(SensoresAA(
                        idSensor      = id_sensor,
                        valor         = _redondear_para_presentacion(valor_float),
                        idCiclo       = id_ciclo,
                        fechaRegistro = fecha,
                    ))

            if objetos:
                session.bulk_save_objects(objetos)
            session.flush()
            logger.info(
                f"SensoresAA: {len(objetos)} filas preparadas "
                f"para ciclo {id_ciclo}."
            )
            return True

        except Exception as e:
            logger.error(
                f"Error preparando SensoresAA para ciclo {id_ciclo}: {e}"
            )
            return False

    # -----------------------------------------------------------------------
    # EstadoCiclo \u2014 persistencia de tramos al cierre del ciclo
    # -----------------------------------------------------------------------

    def _persistir_estados_ciclo(self, session: Session, id_ciclo: int,
                                  tramos: list) -> bool:
        """
        Inserta tramos de estado en EstadoCiclo.
        Idempotente: si ya existen filas para este id_ciclo, no inserta nada.
        Usa flush (sin commit) para participar en la transaccion del cierre.

        Returns True si OK, False si hubo error.
        """
        try:
            ya_existen = session.query(EstadoCiclo).filter_by(
                idCiclo=id_ciclo
            ).first()
            if ya_existen:
                logger.info(
                    f"EstadoCiclo para ciclo {id_ciclo} ya existen. Se omite."
                )
                return True

            if not tramos:
                return True

            objetos = [
                EstadoCiclo(
                    nombre             = t["nombre"],
                    fechaInicio        = t["fechaInicio"],
                    fechaFin           = t["fechaFin"],
                    tiempoTranscurrido = t["tiempoTranscurrido"],
                    idCiclo            = id_ciclo,
                )
                for t in tramos
            ]
            session.bulk_save_objects(objetos)
            session.flush()
            logger.info(
                f"EstadoCiclo: {len(objetos)} tramos preparados "
                f"para ciclo {id_ciclo}."
            )
            return True

        except Exception as e:
            logger.error(
                f"Error preparando EstadoCiclo para ciclo {id_ciclo}: {e}"
            )
            return False

    # -----------------------------------------------------------------------
    # Cierre completo del ciclo (reemplaza al anterior finalizar_ciclo)
    # -----------------------------------------------------------------------

    def finalizar_ciclo_completo(
        self,
        id_ciclo:          int,
        estado_maquina:    str,
        archivo_historial: str,
        equipo_key:        str,
        tipo:              str,
        fecha_fin_forzada: datetime | None = None,
        session:           Session | None = None,
    ) -> bool:
        """
        Cierre completo y atomico de un ciclo.

        Pasos en orden:
          1. Renombrar atomica y exclusivamente a .processing.
          2. Leer y validar el JSONL completo.
          3. Filtrar temperaturas y cambios de NIVEL_AGUA.
          4. Preparar SensoresAA, EstadoCiclo, IO y Ciclo sin commits internos.
          5. Ejecutar un unico commit.
          6. Archivar como .done solamente despues del commit.

        Si fecha_fin_forzada se informa (reinicio de TIEMPO_TRANS), el ciclo se
        cierra en la ultima captura JSONL conocida y no en la reconexion.

        Si algo falla se ejecuta rollback y el .processing se conserva.

        Returns:
            True  -> ciclo cerrado correctamente
            False -> ya estaba cerrado o hubo un error
        """
        # Todo el cierre ocurre dentro de una unica sesion y de un unico
        # commit. Si el llamador no aporta una, se abre y se cierra aqui.
        if session is None:
            with self._sesion_bd() as propia:
                return self.finalizar_ciclo_completo(
                    id_ciclo=id_ciclo,
                    estado_maquina=estado_maquina,
                    archivo_historial=archivo_historial,
                    equipo_key=equipo_key,
                    tipo=tipo,
                    fecha_fin_forzada=fecha_fin_forzada,
                    session=propia,
                )

        archivo_processing = None
        try:
            if not archivo_historial or not os.path.exists(archivo_historial):
                logger.warning(
                    "[JSONL] No existe archivo para cerrar ciclo %s: %s",
                    id_ciclo,
                    archivo_historial,
                )
                return False

            archivo_processing = self._renombrar_a_processing(
                archivo_historial
            )
            historial = list(self._iterar_historial_jsonl(archivo_processing))
            if not historial:
                raise ValueError("El historial JSONL no contiene registros validos")

            ciclos_en_archivo = {int(fila["idCiclo"]) for fila in historial}
            if ciclos_en_archivo != {int(id_ciclo)}:
                raise ValueError(
                    f"El JSONL mezcla ciclos: esperado={id_ciclo}, "
                    f"encontrados={sorted(ciclos_en_archivo)}"
                )

            ciclo = session.query(Ciclo).filter_by(id=id_ciclo).first()
            if not ciclo:
                logger.warning(f"Ciclo {id_ciclo} no encontrado en BD.")
                return False

            ya_estaba_cerrado = ciclo.fecha_fin is not None
            fecha_fin = (
                ciclo.fecha_fin
                if ya_estaba_cerrado
                else (
                    fecha_fin_forzada.replace(microsecond=0)
                    if fecha_fin_forzada is not None
                    else datetime.now().replace(microsecond=0)
                )
            )

            fecha_inicio = ciclo.fecha_inicio or _parse_tiempo(
                historial[0]["tiempo"]
            )
            if fecha_fin < fecha_inicio:
                raise ValueError(
                    f"fecha_fin={fecha_fin} anterior a fecha_inicio="
                    f"{fecha_inicio} para ciclo {id_ciclo}"
                )

            # --- Tramos de estado -----------------------------------------
            tramos = _construir_tramos_estado(historial)

            # Llevar los tramos hasta la fecha_fin real del ciclo. Si el
            # historial termino mucho antes del cierre, ese intervalo se
            # registra como periodo sin datos en lugar de estirar el ultimo
            # estado observado.
            _cerrar_tramos_hasta(tramos, fecha_fin)

            # cantidadPausas = cantidad de tramos con nombre "PAUSADO"
            cantidad_pausas = sum(1 for t in tramos if t["nombre"] == "PAUSADO")

            tramos_sin_datos = sum(
                1 for t in tramos if t["nombre"] == ESTADO_SIN_DATOS
            )
            if tramos_sin_datos:
                logger.warning(
                    "[CICLO %s] Se registraron %s periodos '%s'. Ese tiempo "
                    "no se atribuye a ningun estado productivo.",
                    id_ciclo,
                    tramos_sin_datos,
                    ESTADO_SIN_DATOS,
                )

            # Tiempo total
            tiempo_transcurrido = _delta_str(fecha_inicio, fecha_fin)

            # --- Persistencias (flush sin commit) -------------------------
            historial_filtrado = self._filtrar_historial_para_bd(historial)
            ok_aa = self._persistir_sensores_aa(
                session, id_ciclo, historial_filtrado
            )
            if not ok_aa:
                raise RuntimeError("Fallo _persistir_sensores_aa")

            ok_ec = self._persistir_estados_ciclo(session, id_ciclo, tramos)
            if not ok_ec:
                raise RuntimeError("Fallo _persistir_estados_ciclo")

            # Los cambios IO anteriores se confirmaron en tiempo real. Los
            # tramos abiertos participan ahora del mismo commit del cierre.
            if not ya_estaba_cerrado:
                self._cerrar_io_tramos_equipo(
                    session,
                    equipo_key,
                    tipo,
                    id_ciclo,
                    fecha_fin,
                    limpiar_estado=False,
                )

            # --- Actualizar Ciclo -----------------------------------------
            if not ya_estaba_cerrado:
                ciclo.fecha_fin          = fecha_fin
                ciclo.estadoMaquina      = estado_maquina
                ciclo.cantidadPausas     = cantidad_pausas
                ciclo.tiempoTranscurrido = tiempo_transcurrido

            # --- COMMIT unico (SensoresAA + EstadoCiclo + IO + Ciclo) ------
            session.commit()
            self._io_state.pop(equipo_key, None)

            logger.info(
                f"\033[1;92m[{fecha_fin}] CICLO {id_ciclo} FINALIZADO | "
                f"estado={estado_maquina} | pausas={cantidad_pausas} | "
                f"tiempo={tiempo_transcurrido} | "
                f"registros_crudos={len(historial)} | "
                f"registros_filtrados={len(historial_filtrado)} | "
                f"tramos_estado={len(tramos)}\033[0m"
            )

            # --- Archivar JSONL SOLO tras commit exitoso ------------------
            self._archivar_jsonl(archivo_processing)
            return True

        except Exception as e:
            try:
                session.rollback()
            except Exception:
                pass
            logger.error(
                f"Error en finalizar_ciclo_completo (ciclo={id_ciclo}): {e}. "
                f"Historial JSONL conservado para reintento en "
                f"{archivo_processing or archivo_historial}."
            )
            return False

    # -----------------------------------------------------------------------
    # datosGenerales \u2014 loop principal de lectura y publicacion WebSocket
    # -----------------------------------------------------------------------

    async def datosGenerales(self):
        """
        Punto de entrada asincronico del loop de main.py.

        Toda la lectura del cache, las consultas MySQL y la escritura de los
        JSONL se ejecutan en un hilo aparte. El event loop de FastAPI queda
        libre para atender WebSockets y peticiones HTTP mientras dura la
        pasada, que con 14 equipos puede tardar bastante mas de un segundo si
        MySQL responde lento.
        """
        if not self.conexion_servidor.connected:
            return {
                "datos-cocinas":     [],
                "datos-enfriadores": [],
            }

        return await asyncio.to_thread(self._datos_generales_sync)

    def _datos_generales_sync(self):
        """
        Recorre los equipos fuera del event loop. Un fallo de un equipo no
        interrumpe a los demas.
        """
        resultado = {
            "datos-cocinas":     [],
            "datos-enfriadores": [],
        }

        # Defensa adicional: el Event de main.py normalmente ya detiene este
        # loop, pero nunca se procesa el cache si el cliente esta desconectado.
        if not self.conexion_servidor.connected:
            return resultado

        # Copia local de la lista: iniciar_suscripcion() puede reemplazarla
        # mientras esta pasada esta en curso.
        equipos = list(self._equipos)

        for equipo in equipos:
            try:
                salida = self._procesar_equipo(equipo)
            except Exception as e:
                logger.error(
                    f"Error procesando equipo "
                    f"{equipo.get('linea')}/{equipo.get('tipo')}"
                    f"/{equipo.get('numero_local')}: {e}"
                )
                continue

            if salida is None:
                continue

            clave, payload = salida
            resultado[clave].append(payload)

        return resultado

    def _procesar_equipo(self, equipo: dict):
        """
        Procesa un equipo completo dentro de una unica sesion SQLAlchemy que
        se cierra siempre al terminar.

        Devuelve None cuando el equipo no debe publicarse en esta pasada, o
        una tupla (clave_resultado, payload) en caso contrario.
        """
        linea        = equipo["linea"]
        tipo         = equipo["tipo"]
        numero_local = equipo["numero_local"]
        id_equipo    = equipo["id_equipo"]

        with self._sesion_bd() as session:
            datos = self._leer_datos_equipo(equipo)
            key_estado = self._clave_equipo(
                linea, tipo, numero_local
            )

            fotografia = self._datos_reconciliacion_frescos(
                equipo,
                datos,
                key_estado,
            )
            if fotografia is None:
                # Todavia no llegaron ESTADO_EQUIPO y TIEMPO_TRANS de
                # la suscripcion actual. No usar el resto del cache.
                return None
            estado_actual, tt_actual = fotografia

            if key_estado in self._tt_retrocesos:
                confirmacion_directa = self._leer_confirmacion_tt_directa(
                    equipo,
                    key_estado,
                )
                if confirmacion_directa is None:
                    return None
                estado_actual, tt_actual = confirmacion_directa

            numero_receta = int(datos.get("NUMERO_RECETA") or 0)
            receta_opc    = self._recetario_cache.get(numero_receta, {})
            nombre_receta = receta_opc.get("NOMBRE", f"RECETA_{numero_receta:02}")

            lote_ciclo = str(
                datos.get("LOTE_CICLO") or ""
            ).strip()
            archivo_historial = self._buscar_historial_pendiente(
                linea, tipo, numero_local
            )
            if estado_actual in (
                self.ESTADOS_CONTINUOS | self.ESTADOS_FIN
            ):
                reconciliacion = self._reconciliar_historial_por_tt(
                    session=session,
                    archivo_historial=archivo_historial,
                    tt_actual=tt_actual,
                    estado_actual=estado_actual,
                    equipo_key=key_estado,
                    linea=linea,
                    tipo=tipo,
                    numero_local=numero_local,
                )
            else:
                # LIMPIEZA conserva el comportamiento visual anterior,
                # pero no crea, cierra ni mezcla ciclos pendientes.
                reconciliacion = {
                    "decision": self.RECON_INDETERMINADA,
                    "id_ciclo": None,
                    "archivo": archivo_historial,
                    "historial": self._obtener_preview_jsonl(
                        archivo_historial
                    ),
                }

            decision_reconciliacion = reconciliacion["decision"]
            id_ciclo_reconciliado = reconciliacion["id_ciclo"]
            archivo_historial = reconciliacion["archivo"]
            historial_actual = reconciliacion["historial"]
            operaciones_bloqueadas = decision_reconciliacion in {
                self.RECON_CONFIRMAR_TT,
                self.RECON_INDETERMINADA,
            }

            # ---------------------------------------------------------
            # Ciclo activo: guardar muestra y actualizar IO state
            # ---------------------------------------------------------
            if (
                estado_actual in self.ESTADOS_CONTINUOS
                and not operaciones_bloqueadas
            ):
                receta_id = self.obtener_o_crear_receta(
                    receta_opc, numero_receta, session=session
                )
                id_ciclo  = self._resolver_id_ciclo(
                    session=session,
                    decision_reconciliacion=decision_reconciliacion,
                    id_ciclo_reconciliado=id_ciclo_reconciliado,
                    id_equipo=id_equipo,
                    lote_ciclo=lote_ciclo,
                    receta_id=receta_id,
                    datos_equipo=datos,
                    estado_actual=estado_actual,
                )

                archivo_ciclo = self._archivo_historial(
                    linea,
                    tipo,
                    numero_local,
                    id_ciclo,
                )
                if (
                    archivo_historial
                    and os.path.abspath(archivo_historial)
                    != os.path.abspath(archivo_ciclo)
                ):
                    raise RuntimeError(
                        "El ciclo resuelto no coincide con el archivo "
                        f"pendiente: ciclo={id_ciclo} "
                        f"archivo={archivo_historial}"
                    )
                archivo_historial = archivo_ciclo

                # Timestamp sin microsegundos (consistencia con BD)
                ahora = datetime.now().replace(microsecond=0)

                nuevo_paso = {
                    "tiempo":       ahora.strftime("%Y-%m-%d %H:%M:%S"),
                    "estado":       estado_actual,
                    "idCiclo":      id_ciclo,
                    "lote":         lote_ciclo,
                    "tiempo_trans_opc": tt_actual,
                    "temp_agua":    datos.get("TEMP_AGUA"),
                    "temp_ingreso": datos.get("TEMP_INGRESO"),
                    "temp_prod":    datos.get("TEMP_PRODUCTO"),
                    "niv_agua":     datos.get("NIVEL_AGUA"),
                }
                self._agregar_registro_jsonl(
                    archivo_historial,
                    nuevo_paso,
                )
                historial_actual = self._obtener_preview_jsonl(
                    archivo_historial
                )

                # Trazabilidad IO \u2014 deteccion de cambios en tiempo real
                self._actualizar_io_state(
                    session, key_estado, tipo, datos, id_ciclo, ahora
                )

            # ---------------------------------------------------------
            # Transicion a estado de fin (flanco unico)
            # ---------------------------------------------------------
            if (
                estado_actual in self.ESTADOS_FIN
                and not operaciones_bloqueadas
                and decision_reconciliacion
                != self.RECON_NUEVO_CICLO
            ):
                id_ciclo = None
                ultimo_registro = self._ultimo_registro_del_archivo(
                    archivo_historial
                )
                if ultimo_registro:
                    id_ciclo = ultimo_registro.get("idCiclo")
                if not id_ciclo:
                    ciclo_activo = self._obtener_ciclo_activo(
                        session, id_equipo
                    )
                    if ciclo_activo:
                        id_ciclo = ciclo_activo.id

                if id_ciclo is not None and archivo_historial:
                    ok = self.finalizar_ciclo_completo(
                        id_ciclo          = id_ciclo,
                        estado_maquina    = estado_actual,
                        archivo_historial = archivo_historial,
                        equipo_key        = key_estado,
                        tipo              = tipo,
                        session           = session,
                    )
                    if ok:
                        historial_actual = []
                        archivo_historial = None
                elif id_ciclo is not None:
                    ciclo_activo = session.query(Ciclo).filter_by(
                        id=id_ciclo
                    ).first()
                    if ciclo_activo:
                        self._cerrar_ciclo_bd(
                            session, ciclo_activo, estado_actual
                        )

            # ---------------------------------------------------------
            # Armar payload WebSocket
            # ---------------------------------------------------------
            tiempo_transcurrido = (
                self.calcular_tiempo_transcurrido_jsonl(
                    archivo_historial
                )
                if archivo_historial else "00:00:00"
            )
            ultimo_ciclo = self.obtener_ultimo_ciclo_finalizado(
                id_equipo, session=session
            )
            historial_websocket = self._historial_para_websocket(
                historial_actual
            )

            equipo_general = {
                "tipo":               tipo,
                "id":                 id_equipo,
                "linea":              linea,
                "estado":             estado_actual,
                "temp_agua":          _redondear_para_presentacion(
                    datos.get("TEMP_AGUA")
                ),
                "temp_prod":          _redondear_para_presentacion(
                    datos.get("TEMP_PRODUCTO")
                ),
                "temp_ingreso":       _redondear_para_presentacion(
                    datos.get("TEMP_INGRESO")
                ),
                "niv_agua":           datos.get("NIVEL_AGUA"),
                "receta":             nombre_receta,
                "receta_paso_actual": datos.get("PASO_ACTUAL"),
                "tiempoTranscurrido": tiempo_transcurrido,
                "ultimo_ciclo":       ultimo_ciclo,
            }

            if tipo == "COCINA":
                equipo_detalle = {
                    "num_cocina":    numero_local,
                    "num_receta":    numero_receta,
                    "nom_receta":    nombre_receta,
                    "cant_torres":   datos.get("CANTIDAD_TORRES"),
                    "tipo_fin":      datos.get("CICLO_TIPO_FIN"),
                    "peso_producto": datos.get("PESO_PRODUCTO"),
                    "lote_ciclo":    lote_ciclo,
                    "sector_io": [{
                        "filtro_succion_agua":  datos.get("FILTRO_SUCCION_AGUA"),
                        "entrada_agua":         datos.get("CARGA_AGUA"),
                        "bomba_recirculacion":  datos.get("BOMBA_CENTRIFUGA"),
                        "vapor_serpentina":     datos.get("VAPOR_SERPENTINA"),
                        "vapor_serpentina_acc": datos.get("VAPOR_SERPENTINA_ACC"),
                        "vapor_vivo":           datos.get("VAPOR_VIVO"),
                        "vapor_vivo_acc":       datos.get("VAPOR_VIVO_ACC"),
                    }],
                    "historial": historial_websocket,
                }
                return (
                    "datos-cocinas",
                    [equipo_general, equipo_detalle],
                )

            if tipo == "ENFRIADOR":
                equipo_detalle = {
                    "num_enfriador": numero_local,
                    "num_receta":    numero_receta,
                    "nom_receta":    nombre_receta,
                    "cant_torres":   datos.get("CANTIDAD_TORRES"),
                    "tipo_fin":      datos.get("CICLO_TIPO_FIN"),
                    "peso_producto": datos.get("PESO_PRODUCTO"),
                    "lote_ciclo":    lote_ciclo,
                    "sector_io": [{
                        "filtro_succion_agua":  datos.get("FILTRO_SUCCION_AGUA"),
                        "entrada_agua":         datos.get("CARGA_AGUA"),
                        "bomba_recirculacion":  datos.get("BOMBA_CENTRIFUGA"),
                        "valvula_amoniaco":     datos.get("AMONIACO"),
                        "valvula_amoniaco_acc": datos.get("AMONIACO_ACC"),
                        "vapor_limpieza":       datos.get("VAPOR_LIMPIEZA"),
                        "vapor_limpieza_acc":   datos.get("VAPOR_LIMPIEZA_ACC"),
                    }],
                    "historial": historial_websocket,
                }
                return (
                    "datos-enfriadores",
                    [equipo_general, equipo_detalle],
                )

        return None

    # -----------------------------------------------------------------------
    # Recetario \u2014 sincronizacion a BD
    # -----------------------------------------------------------------------

    async def actualizarRecetas(self):
        await self.cargar_recetario()
        self.guardarRecetaEnBD(self._recetario_cache)

    def guardarRecetaEnBD(self, datosPLC):
        # La sesion se abre y se cierra en el contexto. Si la conexion falla,
        # el error se registra tal cual, sin quedar enmascarado por un
        # rollback sobre una variable inexistente.
        try:
            with self._sesion_bd() as session:
                for numero_receta, datosReceta in sorted(
                    datosPLC.items(), key=lambda x: x[0]
                ):
                    receta_id        = int(numero_receta) + 1
                    receta_existente = session.query(Receta).filter(
                        Receta.id == receta_id
                    ).first()
                    nombre   = datosReceta.get("NOMBRE", f"RECETA_{int(numero_receta):02}")
                    nro_paso = datosReceta.get("PASOS", 0)
                    tipo_fin = datosReceta.get("TIPO_CORTE_ENFRIADO", False)
                    if receta_existente:
                        receta_existente.nombre  = nombre
                        receta_existente.nroPaso = nro_paso
                        receta_existente.tipoFin = tipo_fin
                    else:
                        session.add(Receta(id=receta_id, nombre=nombre,
                                           nroPaso=nro_paso, tipoFin=tipo_fin))
                session.commit()
            logger.info("Recetario sincronizado correctamente desde OPC")
        except Exception as e:
            logger.error(f"Error al guardar/actualizar recetas en BD: {e}")
