from sqlalchemy.orm import Session
from datetime import datetime
from collections import defaultdict, deque
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
import ua

from config.db import get_db

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

# Mapeo campo JSON --> id sensor analógico (SensoresAA)
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
    "temp_agua", "temp_ingreso", "temp_prod", "niv_agua",
)

UMBRAL_CAMBIO = 0.5
VENTANA_PROMEDIO_SEGUNDOS = 10
PERSISTENCIA_REQUERIDA = 3
TIEMPO_MAXIMO_SEGUNDOS = 300
HISTORIAL_PREVIEW_MAX = 100


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


def _construir_tramos_estado(historial: list) -> list[dict]:
    """
    Agrupa filas consecutivas del mismo estado en tramos.

    Maneja correctamente:
    - Historial vacío        -> []
    - Una sola muestra       -> un tramo con fechaInicio == fechaFin
    - Timestamps repetidos   -> el tramo se extiende sin error
    - Timestamps fuera orden -> se ordenan antes de procesar

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
        if nombre == nombre_actual:
            fin_actual = t
        else:
            tramos.append({
                "nombre":             nombre_actual,
                "fechaInicio":        inicio_actual,
                "fechaFin":           fin_actual,
                "tiempoTranscurrido": _delta_str(inicio_actual, fin_actual),
            })
            nombre_actual = nombre
            inicio_actual = t
            fin_actual    = t

    # Último tramo (siempre existe)
    tramos.append({
        "nombre":             nombre_actual,
        "fechaInicio":        inicio_actual,
        "fechaFin":           fin_actual,
        "tiempoTranscurrido": _delta_str(inicio_actual, fin_actual),
    })
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
    ESTADOS_FIN       = {"FINALIZADO", "CANCELADO"}

    def __init__(self, conexion_servidor):
        self.conexion_servidor  = conexion_servidor
        self.session = next(get_db())

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
        # Cuando una señal cambia de valor el tramo anterior se persiste
        # en BD inmediatamente y se abre uno nuevo en memoria.
        # Al cierre del ciclo se persisten los tramos aún abiertos.
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

        os.makedirs(self._jsonl_dir, exist_ok=True)
        self._informar_archivos_pendientes()

    def __del__(self):
        session = getattr(self, "session", None)
        if session is not None:
            try:
                session.close()
            except Exception:
                pass

    # -----------------------------------------------------------------------
    # Navegación del árbol OPC
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
        Reset completo del estado interno + navegacion + suscripcion.
        Llamado al arrancar y despues de cada reconexion OPC.
        """
        self._equipos         = []
        self._recetario_cache = {}
        self._pf_l1_node      = None
        # _io_state no contiene objetos Node, solo valores y fechas. Se conserva
        # en una reconexion para no perder el tramo abierto antes del corte OPC.

        def _setup():
            root_node = self.conexion_servidor.client.get_root_node()
            nodos = self._navegar_arbol_sync(root_node)
            self._recetario_cache = self._obtener_recetario_desde_opc_sync(self._pf_l1_node)
            return nodos

        nodos = await asyncio.to_thread(_setup)
        await self.conexion_servidor.suscribir_nodos(nodos, period_ms)
        logger.info(
            f"Suscripcion iniciada: {len(nodos)} nodos, "
            f"{len(self._recetario_cache)} recetas cargadas."
        )

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

        faltantes = [campo for campo in CAMPOS_HISTORIAL if campo not in registro]
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

        limpio = {
            "id_historial": id_historial,
            "tiempo": tiempo,
            "estado": str(registro.get("estado") or "").strip().upper(),
            "idCiclo": id_ciclo,
            "lote": str(registro.get("lote") or "").strip(),
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
            for campo in CAMPOS_TEMPERATURA:
                fila[campo] = _redondear_para_presentacion(fila.get(campo))
            presentacion.append(fila)
        return presentacion

    # -----------------------------------------------------------------------
    # Saneamiento del historial por cambio de lote
    # -----------------------------------------------------------------------

    def _lote_del_historial(self, historial: list) -> str:
        """
        Devuelve el campo 'lote' de la última fila del historial,
        o cadena vacía si el historial está vacío o no tiene ese campo.
        """
        if not historial:
            return ""
        return str(historial[-1].get("lote") or "").strip()

    def _sanear_historial_por_lote(
        self,
        historial:         list,
        lote_actual:       str,
        estado_actual:     str,
        id_equipo:         int,
        archivo_historial: str,
        equipo_key:        str,
        linea:             str,
        tipo:              str,
        numero_local:      int,
    ) -> list:
        """
        Verifica si el historial JSONL corresponde al lote que el OPC informa
        ahora. Si hay discrepancia, cierra el ciclo colgado con
        finalizar_ciclo_completo() y devuelve [] para iniciar uno nuevo.

        Tabla de decisión
        ─────────────────────────────────────────────────────────────────
        historial vacío             → devolver []   (nada que sanear)
        lote_json vacío             → devolver historial sin tocar
                                      (no se puede comparar)
        lote_actual vacío           → devolver historial sin tocar
                                      (OPC en estado transitorio)
        lote_actual == lote_json    → devolver historial (continuación ok)
        lote_actual != lote_json    → CICLO COLGADO:
                                        1. finalizar_ciclo_completo()
                                        2. devolver []
        ─────────────────────────────────────────────────────────────────
        """
        if not historial:
            return historial

        # Un .processing indica que un cierre anterior no termino. Antes de
        # permitir nuevas capturas se reintenta el cierre idempotente.
        if archivo_historial.endswith(".processing"):
            ultimo = historial[-1]
            id_ciclo_pendiente = ultimo.get("idCiclo")
            if not id_ciclo_pendiente:
                raise RuntimeError(
                    f"Archivo .processing sin idCiclo: {archivo_historial}"
                )
            estado_json = ultimo.get("estado", "CANCELADO")
            estado_cierre = (
                estado_json
                if estado_json in self.ESTADOS_FIN
                else "CANCELADO"
            )
            if not self.finalizar_ciclo_completo(
                id_ciclo=id_ciclo_pendiente,
                estado_maquina=estado_cierre,
                archivo_historial=archivo_historial,
                equipo_key=equipo_key,
                tipo=tipo,
            ):
                raise RuntimeError(
                    f"No se pudo recuperar {archivo_historial}"
                )
            return []

        lote_json = self._lote_del_historial(historial)

        if not lote_json:
            logger.debug(
                f"[{linea}/{tipo}/{numero_local}] Historial sin campo 'lote'. "
                f"No se puede determinar cambio de lote."
            )
            return historial

        if not lote_actual:
            logger.debug(
                f"[{linea}/{tipo}/{numero_local}] LOTE_CICLO OPC vacío. "
                f"Se mantiene historial de lote '{lote_json}' sin cerrar."
            )
            return historial

        if lote_actual == lote_json:
            return historial

        # ── Lote diferente → CICLO COLGADO ──────────────────────────────
        logger.warning(
            f"\033[1;33m[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] "
            f"[{linea}/{tipo}/{numero_local}] CICLO COLGADO DETECTADO. "
            f"Lote JSON='{lote_json}' vs OPC='{lote_actual}'. "
            f"Cerrando ciclo viejo...\033[0m"
        )

        id_ciclo_viejo  = historial[-1].get("idCiclo")
        estado_del_json = historial[-1].get("estado", "CANCELADO")
        estado_cierre   = estado_del_json if estado_del_json in self.ESTADOS_FIN else "CANCELADO"

        if id_ciclo_viejo:
            ok = self.finalizar_ciclo_completo(
                id_ciclo          = id_ciclo_viejo,
                estado_maquina    = estado_cierre,
                archivo_historial = archivo_historial,
                equipo_key        = equipo_key,
                tipo              = tipo,
            )
            if ok:
                logger.info(
                    f"[{linea}/{tipo}/{numero_local}] Ciclo colgado "
                    f"{id_ciclo_viejo} cerrado con estado '{estado_cierre}'."
                )
            else:
                raise RuntimeError(
                    f"No se pudo cerrar el ciclo colgado {id_ciclo_viejo}; "
                    "el JSONL se conserva y el equipo queda bloqueado"
                )
        else:
            raise RuntimeError(
                f"[{linea}/{tipo}/{numero_local}] Historial pendiente sin "
                "idCiclo; no se descarta automaticamente"
            )

        return []

    # -----------------------------------------------------------------------
    # BD — ciclos
    # -----------------------------------------------------------------------

    def _obtener_ciclo_activo(self, id_equipo):
        try:
            return (
                self.session.query(Ciclo)
                .filter(Ciclo.idEquipo == id_equipo, Ciclo.fecha_fin.is_(None))
                .order_by(Ciclo.id.desc())
                .first()
            )
        except Exception as e:
            logger.error(f"Error buscando ciclo activo para equipo {id_equipo}: {e}")
            return None

    def _cerrar_ciclo_bd(self, ciclo: Ciclo, estado_maquina: str):
        """
        Cierre mínimo de un ciclo directamente en BD, sin historial JSON.
        Se usa cuando se detecta una inconsistencia de lote en BD durante
        _resolver_id_ciclo (el ciclo abierto pertenece a un lote diferente).
        """
        try:
            if ciclo.fecha_fin is not None:
                return  # Ya cerrado, nada que hacer
            fecha_fin = datetime.now().replace(microsecond=0)
            ciclo.fecha_fin          = fecha_fin
            ciclo.estadoMaquina      = estado_maquina
            ciclo.cantidadPausas     = 0
            ciclo.tiempoTranscurrido = _delta_str(ciclo.fecha_inicio, fecha_fin)
            self.session.commit()
            logger.warning(
                f"Ciclo BD {ciclo.id} (lote='{ciclo.lote}') cerrado como "
                f"'{estado_maquina}' por inconsistencia de lote."
            )
        except Exception as e:
            self.session.rollback()
            logger.error(f"Error en _cerrar_ciclo_bd (ciclo={ciclo.id}): {e}")

    def _resolver_id_ciclo(self, historial_actual, id_equipo, lote_ciclo,
                            receta_id, datos_equipo, estado_actual):
        """
        Determina el id del ciclo activo.

        Prioridad:
          1. idCiclo del historial JSON — solo si el lote coincide.
             (_sanear_historial_por_lote ya garantizó la consistencia,
             pero esta comprobación actúa como segunda línea de defensa.)
          2. Ciclo abierto en BD para este equipo — solo si el lote coincide.
             Si el ciclo en BD tiene lote diferente, se cierra como CANCELADO
             y se sigue al paso 3.
          3. Crear ciclo nuevo en BD.
        """
        try:
            # 1. Del historial — solo si el lote coincide
            if historial_actual:
                lote_json = self._lote_del_historial(historial_actual)
                if lote_json and lote_ciclo and lote_json == lote_ciclo:
                    ultimo_id = historial_actual[-1].get("idCiclo")
                    if ultimo_id:
                        return ultimo_id
                elif lote_json and lote_ciclo and lote_json != lote_ciclo:
                    # Discrepancia no capturada por el saneamiento previo
                    logger.warning(
                        f"_resolver_id_ciclo: lote JSON '{lote_json}' != "
                        f"lote OPC '{lote_ciclo}'. No se reutiliza idCiclo del historial."
                    )
                else:
                    # Alguno de los dos lotes está vacío: usar el id igualmente
                    # (caso transitorio al arrancar el sistema)
                    ultimo_id = historial_actual[-1].get("idCiclo")
                    if ultimo_id:
                        return ultimo_id

            # 2. Ciclo abierto en BD — verificar lote
            ciclo_activo = self._obtener_ciclo_activo(id_equipo)
            if ciclo_activo:
                lote_bd = str(ciclo_activo.lote or "").strip()
                if not lote_ciclo or not lote_bd or lote_bd == lote_ciclo:
                    # Coincide o alguno está vacío → reutilizar
                    return ciclo_activo.id
                else:
                    # Lote diferente → inconsistencia en BD
                    logger.warning(
                        f"Ciclo activo BD {ciclo_activo.id} tiene lote '{lote_bd}' "
                        f"pero OPC informa lote '{lote_ciclo}'. "
                        f"Se cierra el ciclo BD y se crea uno nuevo."
                    )
                    self._cerrar_ciclo_bd(ciclo_activo, "CANCELADO")

            # 3. Crear ciclo nuevo
            id_ciclo = self.guardarEnBaseCiclo({
                "estadoMaquina":  estado_actual,
                "cantidadTorres": int(datos_equipo.get("CANTIDAD_TORRES") or 0),
                "lote":           lote_ciclo or "",
                "fecha_inicio":   datetime.now(),
                "peso":           int(datos_equipo.get("PESO_PRODUCTO") or 0),
                "idEquipo":       id_equipo,
                "idReceta":       receta_id,
            })
            if id_ciclo is not None:
                return id_ciclo
            return self.obtener_id_ciclo_existente(lote_ciclo, id_equipo)

        except Exception as e:
            logger.error(f"Error resolviendo id de ciclo: {e}")
            return self.obtener_id_ciclo_existente(lote_ciclo, id_equipo)

    def obtener_o_crear_receta(self, receta_opc, numero_receta=None):
        try:
            nombre_receta = receta_opc.get("NOMBRE") if receta_opc else None
            if not nombre_receta:
                nombre_receta = f"RECETA_{int(numero_receta or 0):02}"
            nro_paso = receta_opc.get("PASOS", 0) if receta_opc else 0
            tipo_fin = receta_opc.get("TIPO CORTE ENFRIADO", False) if receta_opc else False

            receta = self.session.query(Receta).filter(Receta.nombre == nombre_receta).first()
            if receta:
                receta.nroPaso = nro_paso
                receta.tipoFin = tipo_fin
                
                self.session.commit()
                return receta.id

            ultimo_id = self.session.query(Receta).order_by(Receta.id.desc()).first()
            nuevo_id  = 1 if not ultimo_id else ultimo_id.id + 1
            self.session.add(Receta(id=nuevo_id, nombre=nombre_receta,
                                    nroPaso=nro_paso, tipoFin=tipo_fin))
            self.session.commit()
            logger.info(f"Nueva receta creada - ID: {nuevo_id}, Nombre: {nombre_receta}")
            return nuevo_id
        except Exception as e:
            self.session.rollback()
            logger.error(f"Error al obtener/crear receta: {e}")
            return None

    def guardarEnBaseCiclo(self, datos):
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
            self.session.add(nuevo_ciclo)
            self.session.commit()
            logger.info(
                f"\033[1;93m[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] "
                f"{self.obtener_nombre_equipo(datos['idEquipo'])} "
                f"NUEVO CICLO: {nuevo_ciclo.id}\033[0m"
            )
            return nuevo_ciclo.id
        except Exception as e:
            self.session.rollback()
            logger.error(f"Error al guardar nuevo ciclo: {e}")
            return None

    def obtener_ultimo_ciclo_finalizado(self, id_equipo):
        try:
            ultimo = (
                self.session.query(Ciclo)
                .filter(
                    Ciclo.idEquipo == id_equipo,
                    Ciclo.estadoMaquina.in_(["FINALIZADO", "CANCELADO"]),
                    Ciclo.fecha_fin.isnot(None),
                )
                .order_by(Ciclo.fecha_fin.desc())
                .first()
            )
            return ultimo.id if ultimo else None
        except Exception:
            return None

    def obtener_id_ciclo_existente(self, lote, idEquipo):
        try:
            ciclo = (
                self.session.query(Ciclo)
                .filter_by(lote=lote, idEquipo=idEquipo, fecha_fin=None)
                .order_by(Ciclo.id.desc())
                .first()
            )
            if ciclo:
                return ciclo.id
            import hashlib
            return int(hashlib.md5(f"{lote}-{idEquipo}".encode()).hexdigest()[:8], 16)
        except Exception as e:
            logger.error(f"Error generando ID para ciclo: {e}")
            import random
            return random.randint(1, 1_000_000)

    # -----------------------------------------------------------------------
    # SensoresIO — trazabilidad en tiempo real
    # -----------------------------------------------------------------------

    def _io_map_para_tipo(self, tipo: str) -> dict[str, int]:
        return IO_SENSOR_MAP_COCINA if tipo == "COCINA" else IO_SENSOR_MAP_ENFRIADOR

    def _persistir_tramo_io(self, id_ciclo: int, id_sensor: int, valor: bool,
                             fecha_inicio: datetime, fecha_fin: datetime,
                             confirmar: bool = True) -> bool:
        """
        Inserta un tramo IO en BD.
        Idempotente: la unicidad se verifica por (idCiclo, idSensor, fechaInicio).
        """
        try:
            existe = self.session.query(SensoresIO).filter_by(
                idCiclo=id_ciclo,
                idSensor=id_sensor,
                fechaInicio=fecha_inicio,
            ).first()
            if existe:
                return True
            self.session.add(SensoresIO(
                idSensor    = id_sensor,
                valor       = valor,
                fechaInicio = fecha_inicio,
                fechaFin    = fecha_fin,
                idCiclo     = id_ciclo,
            ))
            if confirmar:
                self.session.commit()
            else:
                self.session.flush()
            return True
        except Exception as e:
            self.session.rollback()
            logger.error(
                f"Error persistiendo tramo IO "
                f"(ciclo={id_ciclo}, sensor={id_sensor}): {e}"
            )
            if not confirmar:
                raise
            return False

    def _actualizar_io_state(self, equipo_key: str, tipo: str, datos: dict,
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
                        tramo["id_ciclo"], id_sensor,
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
                        id_ciclo, id_sensor,
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

    def _cerrar_io_tramos_equipo(self, equipo_key: str, tipo: str,
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
                id_ciclo, id_sensor,
                tramo["valor"], tramo["inicio"], fecha_fin,
                confirmar=False,
            )

        if limpiar_estado:
            self._io_state.pop(equipo_key, None)

    # -----------------------------------------------------------------------
    # SensoresAA — persistencia historica al cierre del ciclo
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

    def _persistir_sensores_aa(self, id_ciclo: int, historial: list) -> bool:
        """
        Inserta en SensoresAA las muestras filtradas obtenidas del JSONL.
        Idempotente: si ya existen filas para este id_ciclo, no inserta nada.
        Usa flush (sin commit) para participar en la transaccion del cierre.

        Returns True si OK, False si hubo error.
        """
        try:
            ya_existen = self.session.query(SensoresAA).filter_by(
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
                        valor         = valor_float,
                        idCiclo       = id_ciclo,
                        fechaRegistro = fecha,
                    ))

            if objetos:
                self.session.bulk_save_objects(objetos)
            self.session.flush()
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
    # EstadoCiclo — persistencia de tramos al cierre del ciclo
    # -----------------------------------------------------------------------

    def _persistir_estados_ciclo(self, id_ciclo: int, tramos: list) -> bool:
        """
        Inserta tramos de estado en EstadoCiclo.
        Idempotente: si ya existen filas para este id_ciclo, no inserta nada.
        Usa flush (sin commit) para participar en la transaccion del cierre.

        Returns True si OK, False si hubo error.
        """
        try:
            ya_existen = self.session.query(EstadoCiclo).filter_by(
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
            self.session.bulk_save_objects(objetos)
            self.session.flush()
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

        Si algo falla se ejecuta rollback y el .processing se conserva.

        Returns:
            True  -> ciclo cerrado correctamente
            False -> ya estaba cerrado o hubo un error
        """
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

            ciclo = self.session.query(Ciclo).filter_by(id=id_ciclo).first()
            if not ciclo:
                logger.warning(f"Ciclo {id_ciclo} no encontrado en BD.")
                return False

            ya_estaba_cerrado = ciclo.fecha_fin is not None
            fecha_fin = (
                ciclo.fecha_fin
                if ya_estaba_cerrado
                else datetime.now().replace(microsecond=0)
            )

            # --- Tramos de estado -----------------------------------------
            tramos = _construir_tramos_estado(historial)

            # Cerrar el ultimo tramo con la fecha_fin real del ciclo
            # (el historial puede terminar algunos segundos antes)
            if tramos:
                tramos[-1]["fechaFin"]           = fecha_fin
                tramos[-1]["tiempoTranscurrido"] = _delta_str(
                    tramos[-1]["fechaInicio"], fecha_fin
                )

            # cantidadPausas = cantidad de tramos con nombre "PAUSADO"
            cantidad_pausas = sum(1 for t in tramos if t["nombre"] == "PAUSADO")

            # Tiempo total
            fecha_inicio = ciclo.fecha_inicio or _parse_tiempo(
                historial[0]["tiempo"]
            )
            tiempo_transcurrido = _delta_str(fecha_inicio, fecha_fin)

            # --- Persistencias (flush sin commit) -------------------------
            historial_filtrado = self._filtrar_historial_para_bd(historial)
            ok_aa = self._persistir_sensores_aa(
                id_ciclo, historial_filtrado
            )
            if not ok_aa:
                raise RuntimeError("Fallo _persistir_sensores_aa")

            ok_ec = self._persistir_estados_ciclo(id_ciclo, tramos)
            if not ok_ec:
                raise RuntimeError("Fallo _persistir_estados_ciclo")

            # Los cambios IO anteriores se confirmaron en tiempo real. Los
            # tramos abiertos participan ahora del mismo commit del cierre.
            if not ya_estaba_cerrado:
                self._cerrar_io_tramos_equipo(
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
            self.session.commit()
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
            self.session.rollback()
            logger.error(
                f"Error en finalizar_ciclo_completo (ciclo={id_ciclo}): {e}. "
                f"Historial JSONL conservado para reintento en "
                f"{archivo_processing or archivo_historial}."
            )
            return False

    # -----------------------------------------------------------------------
    # datosGenerales — loop principal de lectura y publicacion WebSocket
    # -----------------------------------------------------------------------

    async def datosGenerales(self):
        resultado = {
            "datos-cocinas":     [],
            "datos-enfriadores": [],
        }

        try:
            for equipo in self._equipos:
                try:
                    linea        = equipo["linea"]
                    tipo         = equipo["tipo"]
                    numero_local = equipo["numero_local"]
                    id_equipo    = equipo["id_equipo"]

                    datos = self._leer_datos_equipo(equipo)

                    estado_actual   = self._estado_a_texto(datos.get("ESTADO_EQUIPO", 0))
                    key_estado      = f"{linea}-{tipo}-{numero_local}"

                    numero_receta = int(datos.get("NUMERO_RECETA") or 0)
                    receta_opc    = self._recetario_cache.get(numero_receta, {})
                    nombre_receta = receta_opc.get("NOMBRE", f"RECETA_{numero_receta:02}")

                    lote_ciclo = str(
                        datos.get("LOTE_CICLO") or ""
                    ).strip()
                    archivo_historial = self._buscar_historial_pendiente(
                        linea, tipo, numero_local
                    )
                    historial_actual = self._obtener_preview_jsonl(
                        archivo_historial
                    )

                    if (
                        estado_actual in self.ESTADOS_CONTINUOS
                        and archivo_historial
                    ):
                        historial_actual = self._sanear_historial_por_lote(
                            historial         = historial_actual,
                            lote_actual       = lote_ciclo,
                            estado_actual     = estado_actual,
                            id_equipo         = id_equipo,
                            archivo_historial = archivo_historial,
                            equipo_key        = key_estado,
                            linea             = linea,
                            tipo              = tipo,
                            numero_local      = numero_local,
                        )
                        if not historial_actual:
                            archivo_historial = None

                    # ---------------------------------------------------------
                    # Ciclo activo: guardar muestra y actualizar IO state
                    # ---------------------------------------------------------
                    if estado_actual in self.ESTADOS_CONTINUOS:
                        receta_id = self.obtener_o_crear_receta(receta_opc, numero_receta)
                        id_ciclo  = self._resolver_id_ciclo(
                            historial_actual = historial_actual,
                            id_equipo        = id_equipo,
                            lote_ciclo       = lote_ciclo,
                            receta_id        = receta_id,
                            datos_equipo     = datos,
                            estado_actual    = estado_actual,
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

                        # Trazabilidad IO — deteccion de cambios en tiempo real
                        self._actualizar_io_state(key_estado, tipo, datos, id_ciclo, ahora)

                    # ---------------------------------------------------------
                    # Transicion a estado de fin (flanco unico)
                    # ---------------------------------------------------------
                    if estado_actual in self.ESTADOS_FIN:
                        id_ciclo = None
                        ultimo_registro = self._ultimo_registro_del_archivo(
                            archivo_historial
                        )
                        if ultimo_registro:
                            id_ciclo = ultimo_registro.get("idCiclo")
                        if not id_ciclo:
                            ciclo_activo = self._obtener_ciclo_activo(id_equipo)
                            if ciclo_activo:
                                id_ciclo = ciclo_activo.id

                        if id_ciclo is not None and archivo_historial:
                            ok = self.finalizar_ciclo_completo(
                                id_ciclo          = id_ciclo,
                                estado_maquina    = estado_actual,
                                archivo_historial = archivo_historial,
                                equipo_key        = key_estado,
                                tipo              = tipo,
                            )
                            if ok:
                                historial_actual = []
                                archivo_historial = None
                        elif id_ciclo is not None:
                            ciclo_activo = self.session.query(Ciclo).filter_by(
                                id=id_ciclo
                            ).first()
                            if ciclo_activo:
                                self._cerrar_ciclo_bd(
                                    ciclo_activo, estado_actual
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
                    ultimo_ciclo = self.obtener_ultimo_ciclo_finalizado(id_equipo)
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
                        resultado["datos-cocinas"].append([equipo_general, equipo_detalle])

                    elif tipo == "ENFRIADOR":
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
                        resultado["datos-enfriadores"].append([equipo_general, equipo_detalle])

                except Exception as e:
                    logger.error(
                        f"Error procesando equipo "
                        f"{equipo.get('linea')}/{equipo.get('tipo')}"
                        f"/{equipo.get('numero_local')}: {e}"
                    )

        except Exception as e:
            logger.error(f"Error general en datosGenerales: {e}")

        return resultado

    # -----------------------------------------------------------------------
    # Recetario — sincronizacion a BD
    # -----------------------------------------------------------------------

    async def actualizarRecetas(self):
        await self.cargar_recetario()
        self.guardarRecetaEnBD(self._recetario_cache)

    def guardarRecetaEnBD(self, datosPLC):
        try:
            db: Session = next(get_db())
            for numero_receta, datosReceta in sorted(datosPLC.items(), key=lambda x: x[0]):
                receta_id        = int(numero_receta) + 1
                receta_existente = db.query(Receta).filter(Receta.id == receta_id).first()
                nombre   = datosReceta.get("NOMBRE", f"RECETA_{int(numero_receta):02}")
                nro_paso = datosReceta.get("PASOS", 0)
                tipo_fin = datosReceta.get("TIPO_CORTE_ENFRIADO", False)
                if receta_existente:
                    receta_existente.nombre  = nombre
                    receta_existente.nroPaso = nro_paso
                    receta_existente.tipoFin = tipo_fin
                else:
                    db.add(Receta(id=receta_id, nombre=nombre,
                                  nroPaso=nro_paso, tipoFin=tipo_fin))
            db.commit()
            logger.info("Recetario sincronizado correctamente desde OPC")
        except Exception as e:
            db.rollback()
            logger.error(f"Error al guardar/actualizar recetas en BD: {e}")
        finally:
            db.close()