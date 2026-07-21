from sqlalchemy.orm import Session
from datetime import datetime
from models.ciclo import Ciclo
from models.receta import Receta
from models.sensoresAA import SensoresAA
from models.sensoresIO import SensoresIO
from models.estadoCiclo import EstadoCiclo

import os
import json
import logging
import asyncio
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


# -----------------------------------------------------------------------
# Helpers puros (sin estado de clase)
# -----------------------------------------------------------------------

def datetime_to_string(obj):
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Object of type {type(obj)} is not JSON serializable")


def limpiar_archivo_json(archivo: str):
    try:
        if os.path.exists(archivo):
            with open(archivo, "w") as f:
                json.dump([], f)
            logger.info(f"Historial JSON limpiado: {archivo}")
    except Exception as e:
        logger.error(f"Error al limpiar {archivo}: {e}")


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

    ordenado = sorted(historial, key=lambda r: r.get("tiempo", ""))

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
        self.estados_anteriores: dict = {}
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

    def __del__(self):
        if hasattr(self, "session"):
            self.session.close()

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
        # Descartar tramos IO en memoria: los nodos OPC anteriores son invalidos
        self._io_state        = {}

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

    def _archivo_historial(self, linea, tipo, numero_local):
        linea_slug = "l1" if linea == "PF-L1" else "l2"
        tipo_slug  = "cocina" if tipo == "COCINA" else "enfriador"
        return f"{tipo_slug}_{numero_local}_{linea_slug}.json"

    def _cargar_historial_json(self, archivo):
        try:
            if os.path.exists(archivo) and os.path.getsize(archivo) > 0:
                try:
                    with open(archivo, "r") as f:
                        return json.load(f)
                except json.JSONDecodeError:
                    logger.error(f"JSON invalido en {archivo}, se limpia.")
                    limpiar_archivo_json(archivo)
                    return []
            return []
        except Exception as e:
            logger.error(f"Error al cargar {archivo}: {e}")
            return []

    def _guardar_historial_json(self, archivo, historial):
        try:
            with open(archivo, "w") as f:
                json.dump(historial, f, indent=2, default=datetime_to_string)
        except Exception as e:
            logger.error(f"Error guardando historial en {archivo}: {e}")

    def calcular_tiempo_transcurrido_json(self, historial_actual):
        try:
            if not historial_actual or len(historial_actual) < 2:
                return "00:00:00"
            t0 = _parse_tiempo(historial_actual[0]["tiempo"])
            t1 = _parse_tiempo(historial_actual[-1]["tiempo"])
            return _delta_str(t0, t1)
        except Exception:
            return "00:00:00"

    # -----------------------------------------------------------------------
    # Saneamiento del historial por cambio de lote
    # -----------------------------------------------------------------------

    def _lote_del_historial(self, historial: list) -> str:
        """
        Devuelve el campo 'lote' de la última fila del historial JSON,
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
        Verifica si el historial JSON corresponde al lote que el OPC informa
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
                historial         = historial,
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
                # finalizar_ciclo_completo ya logueó el error;
                # de todas formas limpiamos el JSON para no seguir acumulando.
                logger.warning(
                    f"[{linea}/{tipo}/{numero_local}] No se pudo cerrar "
                    f"el ciclo colgado {id_ciclo_viejo}. Se descarta el JSON igual."
                )
                limpiar_archivo_json(archivo_historial)
        else:
            logger.warning(
                f"[{linea}/{tipo}/{numero_local}] Historial colgado sin idCiclo. "
                f"Se descarta el JSON sin persistir."
            )
            limpiar_archivo_json(archivo_historial)

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
                             fecha_inicio: datetime, fecha_fin: datetime):
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
                return
            self.session.add(SensoresIO(
                idSensor    = id_sensor,
                valor       = valor,
                fechaInicio = fecha_inicio,
                fechaFin    = fecha_fin,
                idCiclo     = id_ciclo,
            ))
            self.session.commit()
        except Exception as e:
            self.session.rollback()
            logger.error(
                f"Error persistiendo tramo IO "
                f"(ciclo={id_ciclo}, sensor={id_sensor}): {e}"
            )

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
                    self._persistir_tramo_io(
                        tramo["id_ciclo"], id_sensor,
                        tramo["valor"], tramo["inicio"], ahora,
                    )
                    estado_equipo[campo] = {
                        "valor":    valor,
                        "inicio":   ahora,
                        "id_ciclo": id_ciclo,
                    }
                elif tramo["valor"] != valor:
                    # Valor cambio: cerrar tramo actual y abrir nuevo
                    self._persistir_tramo_io(
                        id_ciclo, id_sensor,
                        tramo["valor"], tramo["inicio"], ahora,
                    )
                    estado_equipo[campo] = {
                        "valor":    valor,
                        "inicio":   ahora,
                        "id_ciclo": id_ciclo,
                    }
                # Valor igual y mismo ciclo: no hacer nada

    def _cerrar_io_tramos_equipo(self, equipo_key: str, tipo: str,
                                  id_ciclo: int, fecha_fin: datetime):
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
            )

        self._io_state.pop(equipo_key, None)

    # -----------------------------------------------------------------------
    # SensoresAA — persistencia historica al cierre del ciclo
    # -----------------------------------------------------------------------

    def _persistir_sensores_aa(self, id_ciclo: int, historial: list) -> bool:
        """
        Inserta en SensoresAA todas las muestras del historial JSON.
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

            ordenado = sorted(historial, key=lambda r: r.get("tiempo", ""))
            objetos  = []

            for fila in ordenado:
                try:
                    fecha = _parse_tiempo(fila["tiempo"])
                except Exception:
                    continue  # Timestamp invalido: ignorar fila

                for campo_json, id_sensor in AA_SENSOR_MAP.items():
                    valor = fila.get(campo_json)
                    if valor is None:
                        continue
                    try:
                        valor_float = float(valor)
                    except (TypeError, ValueError):
                        continue
                    objetos.append(SensoresAA(
                        idSensor      = id_sensor,
                        valor         = valor_float,
                        idCiclo       = id_ciclo,
                        fechaRegistro = fecha,
                    ))

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
        historial:         list,
        archivo_historial: str,
        equipo_key:        str,
        tipo:              str,
    ) -> bool:
        """
        Cierre completo y atomico de un ciclo.

        Pasos en orden:
          1. Verificar que el ciclo existe en BD y no esta ya cerrado.
          2. Construir tramos de estado desde el historial JSON.
          3. Corregir el ultimo tramo para que cierre con la fecha_fin real.
          4. Calcular cantidadPausas (tramos con nombre == "PAUSADO").
          5. _persistir_sensores_aa  (flush, sin commit aun)
          6. _persistir_estados_ciclo (flush, sin commit aun)
          7. _cerrar_io_tramos_equipo (commit propio por senal)
          8. Actualizar Ciclo.
          9. COMMIT unico de los pasos 5, 6 y 8.
         10. Limpiar historial JSON SOLO despues del commit exitoso.

        Si el commit falla: ROLLBACK automatico del ORM y el JSON se conserva
        intacto para reintentar en la siguiente deteccion del estado de fin.

        Returns:
            True  -> ciclo cerrado correctamente
            False -> ya estaba cerrado o hubo un error
        """
        try:
            ciclo = self.session.query(Ciclo).filter_by(id=id_ciclo).first()
            if not ciclo:
                logger.warning(f"Ciclo {id_ciclo} no encontrado en BD.")
                return False

            if ciclo.fecha_fin is not None:
                # Ya estaba cerrado (doble disparo): solo limpiar JSON
                logger.info(
                    f"Ciclo {id_ciclo} ya estaba finalizado. Limpiando JSON."
                )
                limpiar_archivo_json(archivo_historial)
                return True

            fecha_fin = datetime.now().replace(microsecond=0)

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
            tiempo_transcurrido = _delta_str(ciclo.fecha_inicio, fecha_fin)

            # --- Persistencias (flush sin commit) -------------------------

            ok_aa = self._persistir_sensores_aa(id_ciclo, historial)
            if not ok_aa:
                raise RuntimeError("Fallo _persistir_sensores_aa")

            ok_ec = self._persistir_estados_ciclo(id_ciclo, tramos)
            if not ok_ec:
                raise RuntimeError("Fallo _persistir_estados_ciclo")

            # SensoresIO: cada tramo IO ya tiene su propio commit en tiempo real.
            # Aqui solo se cierran los tramos que quedaron abiertos al finalizar.
            self._cerrar_io_tramos_equipo(equipo_key, tipo, id_ciclo, fecha_fin)

            # --- Actualizar Ciclo -----------------------------------------
            ciclo.fecha_fin          = fecha_fin
            ciclo.estadoMaquina      = estado_maquina
            ciclo.cantidadPausas     = cantidad_pausas
            ciclo.tiempoTranscurrido = tiempo_transcurrido

            # --- COMMIT unico (SensoresAA + EstadoCiclo + Ciclo) ----------
            self.session.commit()

            logger.info(
                f"\033[1;92m[{fecha_fin}] CICLO {id_ciclo} FINALIZADO | "
                f"estado={estado_maquina} | pausas={cantidad_pausas} | "
                f"tiempo={tiempo_transcurrido} | "
                f"muestras_aa={len(historial)*len(AA_SENSOR_MAP)} | "
                f"tramos_estado={len(tramos)}\033[0m"
            )

            # --- Limpiar JSON SOLO tras commit exitoso --------------------
            limpiar_archivo_json(archivo_historial)
            return True

        except Exception as e:
            self.session.rollback()
            logger.error(
                f"Error en finalizar_ciclo_completo (ciclo={id_ciclo}): {e}. "
                f"Historial JSON conservado para reintento."
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
                    estado_anterior = self.estados_anteriores.get(key_estado, "")
                    self.estados_anteriores[key_estado] = estado_actual

                    numero_receta = int(datos.get("NUMERO_RECETA") or 0)
                    receta_opc    = self._recetario_cache.get(numero_receta, {})
                    nombre_receta = receta_opc.get("NOMBRE", f"RECETA_{numero_receta:02}")

                    archivo_historial = self._archivo_historial(linea, tipo, numero_local)
                    historial_actual  = self._cargar_historial_json(archivo_historial)
                    lote_ciclo        = str(datos.get("LOTE_CICLO") or "").strip()

                    if estado_actual in self.ESTADOS_CONTINUOS:
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

                        # Timestamp sin microsegundos (consistencia con BD)
                        ahora = datetime.now().replace(microsecond=0)

                        nuevo_paso = {
                            "id_historial": len(historial_actual) + 1,
                            "tiempo":       ahora.strftime("%Y-%m-%d %H:%M:%S"),
                            "estado":       estado_actual,
                            "idCiclo":      id_ciclo,
                            "lote":         lote_ciclo,
                            "temp_agua":    datos.get("TEMP_AGUA"),
                            "temp_ingreso": datos.get("TEMP_INGRESO"),
                            "temp_prod":    datos.get("TEMP_PRODUCTO"),
                            "niv_agua":     datos.get("NIVEL_AGUA"),
                        }
                        historial_actual.append(nuevo_paso)
                        self._guardar_historial_json(archivo_historial, historial_actual)

                        # Trazabilidad IO — deteccion de cambios en tiempo real
                        self._actualizar_io_state(key_estado, tipo, datos, id_ciclo, ahora)

                    # ---------------------------------------------------------
                    # Transicion a estado de fin (flanco unico)
                    # ---------------------------------------------------------
                    if estado_actual in self.ESTADOS_FIN and estado_anterior not in self.ESTADOS_FIN:
                        id_ciclo = None
                        if historial_actual:
                            id_ciclo = historial_actual[-1].get("idCiclo")
                        if not id_ciclo:
                            ciclo_activo = self._obtener_ciclo_activo(id_equipo)
                            if ciclo_activo:
                                id_ciclo = ciclo_activo.id

                        if id_ciclo is not None:
                            ok = self.finalizar_ciclo_completo(
                                id_ciclo          = id_ciclo,
                                estado_maquina    = estado_actual,
                                historial         = historial_actual,
                                archivo_historial = archivo_historial,
                                equipo_key        = key_estado,
                                tipo              = tipo,
                            )
                            if ok:
                                historial_actual = []

                    # ---------------------------------------------------------
                    # Armar payload WebSocket
                    # ---------------------------------------------------------
                    tiempo_transcurrido = (
                        self.calcular_tiempo_transcurrido_json(historial_actual)
                        if historial_actual else "00:00:00"
                    )
                    ultimo_ciclo = self.obtener_ultimo_ciclo_finalizado(id_equipo)

                    equipo_general = {
                        "tipo":               tipo,
                        "id":                 id_equipo,
                        "linea":              linea,
                        "estado":             estado_actual,
                        "temp_agua":          datos.get("TEMP_AGUA"),
                        "temp_prod":          datos.get("TEMP_PRODUCTO"),
                        "temp_ingreso":       datos.get("TEMP_INGRESO"),
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
                            "historial": historial_actual,
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
                            "historial": historial_actual,
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