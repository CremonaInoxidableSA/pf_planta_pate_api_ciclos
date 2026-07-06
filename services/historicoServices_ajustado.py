from collections import defaultdict
from datetime import date, datetime, time, timedelta
from bisect import bisect_right
 
from models.ciclo import Ciclo
from models.sensoresAA import SensoresAA
from models.sensores import Sensores
from models.estadoCiclo import EstadoCiclo
from models.receta import Receta
from models.equipo import Equipo
from models.historicoAlarma import HistoricoAlarma
from models.alarmas import Alarmas
from models.alarmasL2 import AlarmasL2
from models.sensoresIO import SensoresIO
 
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from io import BytesIO
 
from pathlib import Path
import copy as _copy
import unicodedata
import logging
 
logger = logging.getLogger("uvicorn")
 
_TEMPLATE_PATH = (
    Path(__file__).resolve().parent.parent
    / "data"
    / "informe_alarmas_fecha_registro.xlsx"
)
 
_TEMPLATE_INFORME_CICLO_PATH = (
    Path(__file__).resolve().parent.parent
    / "data"
    / "INFORME_CICLO.xlsx"
)
 
_HOJA_DETALLE = "DETALLES CICLO"
_HOJA_SENSORES_A_QUITAR = "SENSORES"
 
# Fila donde empieza el header de la tabla en la plantilla y donde
# empiezan los datos (confirmado inspeccionando INFORME_CICLO.xlsx).
_HEADER_ROW = 10
_DATA_START_ROW = 11
_TABLE_COLS = 10  # A..J
 
# ------------------------------------------------------------
# Mapeo de señales -> nombres candidatos en la tabla `Sensores`.
#
# IMPORTANTE: no se pudo verificar contra insert_sensores.sql ni
# contra los servicios OPC (no fueron adjuntados), por lo que se
# resuelven los IDs POR NOMBRE en cada request (igual que ya hacía
# el código original para los sensores analógicos), probando estas
# variantes en orden hasta encontrar coincidencia. Ajustar esta
# lista si los nombres reales en BDD difieren.
# ------------------------------------------------------------

# COCINAS
ID_SENSOR_TEMP_AGUA = 1
ID_SENSOR_TEMP_INGRESO = 2
ID_SENSOR_TEMP_PRODUCTO = 3
ID_SENSOR_NIVEL_AGUA = 5
ID_SENSOR_BOMBA_CENTRIFUGA = 6
ID_SENSOR_VAPOR_SERPENTINA_ACC = 13   # accionamiento, único persistido
ID_SENSOR_VAPOR_SERPENTINA = 7 
ID_SENSOR_VAPOR_VIVO_ACC = 14         # accionamiento, único persistido
ID_SENSOR_VAPOR_VIVO = 9
       
ID_SENSOR_FILTRO_SUCCION_AGUA = 15

# ENFRIADORES 
ID_SENSOR_AMONIACO = 8
ID_SENSOR_VAPOR_LIMPIEZA = 10

_IDS_ANALOGICOS = {
    "temp_agua": ID_SENSOR_TEMP_AGUA,
    "temp_ingreso": ID_SENSOR_TEMP_INGRESO,
    "temp_producto": ID_SENSOR_TEMP_PRODUCTO,
    "nivel_agua": ID_SENSOR_NIVEL_AGUA,
}
 
# Cada tipo de equipo mantiene su propio conjunto de señales.
# Se usan tuplas para evitar modificaciones accidentales entre requests.
_SENALES_BOOLEANAS_COCINA = (
    ("bomba_centrifuga", "BOMBA CENTRIFUGA", ID_SENSOR_BOMBA_CENTRIFUGA),
    ("filtro_succion_agua", "FILTRO SUCCION AGUA", ID_SENSOR_FILTRO_SUCCION_AGUA),
    ("vapor_serpentina", "VAPOR SERPENTINA", ID_SENSOR_VAPOR_SERPENTINA),
    ("vapor_vivo", "VAPOR VIVO", ID_SENSOR_VAPOR_VIVO),
)

_SENALES_BOOLEANAS_ENFRIADOR = (
    ("bomba_centrifuga", "BOMBA CENTRIFUGA", ID_SENSOR_BOMBA_CENTRIFUGA),
    ("filtro_succion_agua", "FILTRO SUCCION AGUA", ID_SENSOR_FILTRO_SUCCION_AGUA),
    ("vapor_limpieza", "VAPOR LIMPIEZA", ID_SENSOR_VAPOR_LIMPIEZA),
    ("amoniaco", "AMONIACO", ID_SENSOR_AMONIACO),
)
 
_HOJA_DETALLE = "DETALLES CICLO"
_HOJA_SENSORES_A_QUITAR = "SENSORES"
_HEADER_ROW = 10
_DATA_START_ROW = 11
_TABLE_COLS = 10  # A..J
 
_TEMPLATE_INFORME_CICLO_PATH = (
    Path(__file__).resolve().parent.parent
    / "data"
    / "INFORME_CICLO.xlsx"
)
 
 
# ============================================================
# FUNCIONES AUXILIARES (reemplazan las versiones previas)
# ============================================================
 
def _construir_serie_analogica(registros: list) -> tuple:
    """
    A partir de una lista de SensoresAA (de un único sensor), devuelve
    (fechas_ordenadas, valores_ordenados) listas para bisect.
 
    Timestamps repetidos: se conserva el de mayor `id` (registro más
    reciente físicamente insertado), asumiendo que id autoincremental
    refleja el orden real de escritura aunque el reloj se repita.
    """
    if not registros:
        return [], []
    registros_ordenados = sorted(registros, key=lambda r: (r.fechaRegistro, r.id))
    fechas, valores = [], []
    for r in registros_ordenados:
        if fechas and fechas[-1] == r.fechaRegistro:
            valores[-1] = r.valor  # mismo timestamp, mayor id gana (orden estable)
        else:
            fechas.append(r.fechaRegistro)
            valores.append(r.valor)
    return fechas, valores
 
 
def _valor_vigente_ffill(fechas: list, valores: list, ts):
    """
    Último valor cuya fechaRegistro <= ts (forward-fill). Si ts es
    anterior a cualquier registro, devuelve None (no se inventan
    valores numéricos, ver requisito #11).
    """
    if not fechas:
        return None
    idx = bisect_right(fechas, ts) - 1
    if idx < 0:
        return None
    return valores[idx]
 
 
def _construir_intervalos(registros: list, campo_inicio: str, campo_fin: str, campo_valor: str,
                           etiqueta: str = "") -> list:
    """
    Convierte una lista de objetos (EstadoCiclo o SensoresIO) en una
    lista de tuplas (inicio, fin, valor) ordenada por inicio.
    fin=None se trata como intervalo abierto (vigente sin límite).
 
    Detecta y loguea (sin abortar) intervalos inválidos o solapados:
    fechaFin < fechaInicio, e intervalos que se superponen entre sí.
    """
    intervalos = []
    for r in registros:
        inicio = getattr(r, campo_inicio)
        fin = getattr(r, campo_fin)
        valor = getattr(r, campo_valor)
        if inicio is None:
            continue
        if fin is not None and fin < inicio:
            logger.warning(
                f"[generar_informe_ciclo] Intervalo inválido en {etiqueta}: "
                f"fechaFin ({fin}) < fechaInicio ({inicio}). Se conserva tal cual."
            )
        intervalos.append((inicio, fin, valor))
 
    intervalos.sort(key=lambda x: x[0])
 
    for i in range(1, len(intervalos)):
        prev_inicio, prev_fin, _ = intervalos[i - 1]
        cur_inicio, _, _ = intervalos[i]
        if prev_fin is not None and cur_inicio < prev_fin:
            logger.warning(
                f"[generar_informe_ciclo] Intervalos solapados en {etiqueta}: "
                f"[{prev_inicio} - {prev_fin}] se superpone con inicio {cur_inicio}."
            )
 
    return intervalos
 
 
def _valor_vigente_intervalo(intervalos: list, inicios: list, ts):
    """
    Política de intervalos: semántica [inicio, fin) -> el intervalo que
    EMPIEZA en ts tiene prioridad sobre el que termina en ts (requisito #5).
 
    Excepción: si ts coincide exactamente con el fin del intervalo más
    reciente que comenzó antes o en ts (típicamente el último intervalo,
    o el cierre oficial del ciclo) y ningún otro intervalo lo sucede en
    ese instante, se devuelve su valor de forma inclusiva, para que la
    fila final del reporte no quede sin información (requisito #17).
    """
    if not intervalos:
        return None
    idx = bisect_right(inicios, ts) - 1
    if idx < 0:
        return None
    inicio, fin, valor = intervalos[idx]
    if inicio <= ts and (fin is None or ts < fin):
        return valor
    if fin is not None and ts == fin:
        return valor
    return None
 
 
def _booleano_a_texto(valor) -> str:
    if valor is True:
        return "ACTIVO"
    if valor is False:
        return "INACTIVO"
    return "INDETERMINADO"
 
 
def _aplicar_estilo(celda_destino, estilo: dict):
    celda_destino.font = _copy.copy(estilo["font"])
    celda_destino.fill = _copy.copy(estilo["fill"])
    celda_destino.border = _copy.copy(estilo["border"])
    celda_destino.alignment = _copy.copy(estilo["alignment"])
    celda_destino.number_format = estilo["number_format"]
 
 
# ============================================================
# FUNCIÓN PRINCIPAL
# ============================================================
 
def generar_informe_ciclo(db, id_ciclo: int, equipo: str):
    """
    Genera la hoja 'DETALLES CICLO' del informe de ciclo, combinando
    cronológicamente registros analógicos (SensoresAA), cambios de
    estado (EstadoCiclo), cambios booleanos (SensoresIO) y el inicio /
    fin oficial del ciclo (Ciclo.fecha_inicio / Ciclo.fecha_fin).
 
    Reglas de búsqueda de ciclo (sin cambios):
      - id_ciclo != 0 -> busca ese ciclo puntual.
      - id_ciclo == 0 -> busca el último ciclo registrado.
      - Si no existe -> ValueError("Ciclo no encontrado").
 
    Devuelve un BytesIO listo para StreamingResponse.
    """
    # ---------- 1. Ciclo ----------
    if id_ciclo != 0:
        ciclo: Ciclo = db.query(Ciclo).filter(Ciclo.id == id_ciclo).first()
    else:
        ciclo: Ciclo = db.query(Ciclo).order_by(Ciclo.id.desc()).first()
 
    if not ciclo:
        logger.info(f"[generar_informe_ciclo] No se encontró el ciclo id={id_ciclo} en la BDD")
        raise ValueError("Ciclo no encontrado")
 
    logger.info(
        f"[generar_informe_ciclo] ciclo.id={ciclo.id} "
        f"fecha_inicio={ciclo.fecha_inicio} fecha_fin={ciclo.fecha_fin} "
        f"estadoMaquina={ciclo.estadoMaquina} (equipo param='{equipo}')"
    )
 
    # ---------- 2. Relaciones ----------
    receta = db.query(Receta).filter(Receta.id == ciclo.idReceta).first() if ciclo.idReceta else None
    maquina = db.query(Equipo).filter(Equipo.id == ciclo.idEquipo).first() if ciclo.idEquipo else None
 
    nombre_receta = receta.nombre if receta else "Sin receta"
    tipo_fin = receta.tipoFin if receta else "N/A"
    nombre_equipo = maquina.nombre if maquina else "Sin equipo"
 
    # Selección local por request. No se reasignan constantes globales,
    # evitando UnboundLocalError y cruces de configuración entre equipos.
    nombre_para_tipo = f"{nombre_equipo} {equipo or ''}".lower()
    es_enfriador = "enfriador" in nombre_para_tipo
    senales_booleanas = (
        _SENALES_BOOLEANAS_ENFRIADOR
        if es_enfriador
        else _SENALES_BOOLEANAS_COCINA
    )

    logger.info(
        f"[generar_informe_ciclo] Tipo de equipo detectado: "
        f"{'ENFRIADOR' if es_enfriador else 'COCINA'}. Señales del reporte: "
        + ", ".join(etiqueta for _, etiqueta, _ in senales_booleanas)
    )
 
    # ---------- 3. Consultas agrupadas (sin N+1) ----------
    registros_analogicos = (
        db.query(SensoresAA)
        .filter(SensoresAA.idCiclo == ciclo.id, SensoresAA.idSensor.in_(_IDS_ANALOGICOS.values()))
        .all()
    )
    registros_por_sensor_aa = defaultdict(list)
    for r in registros_analogicos:
        registros_por_sensor_aa[r.idSensor].append(r)
 
    fechas_agua, valores_agua = _construir_serie_analogica(registros_por_sensor_aa.get(ID_SENSOR_TEMP_AGUA, []))
    fechas_ingreso, valores_ingreso = _construir_serie_analogica(registros_por_sensor_aa.get(ID_SENSOR_TEMP_INGRESO, []))
    fechas_producto, valores_producto = _construir_serie_analogica(registros_por_sensor_aa.get(ID_SENSOR_TEMP_PRODUCTO, []))
    fechas_nivel, valores_nivel = _construir_serie_analogica(registros_por_sensor_aa.get(ID_SENSOR_NIVEL_AGUA, []))
 
    for clave, fechas in (("temp_agua", fechas_agua), ("temp_ingreso", fechas_ingreso),
                          ("temp_producto", fechas_producto), ("nivel_agua", fechas_nivel)):
        if not fechas:
            logger.warning(f"[generar_informe_ciclo] Ciclo {ciclo.id}: sin registros de '{clave}'.")
 
    logger.info(
        f"[generar_informe_ciclo] Registros analógicos -> "
        f"agua={len(fechas_agua)} ({fechas_agua[0]} a {fechas_agua[-1]} si hay) "
        f"ingreso={len(fechas_ingreso)} producto={len(fechas_producto)} nivel={len(fechas_nivel)}"
        if fechas_agua else
        f"[generar_informe_ciclo] Registros analógicos -> agua=0 ingreso={len(fechas_ingreso)} "
        f"producto={len(fechas_producto)} nivel={len(fechas_nivel)}"
    )
 
    # ---------- 4. Estados del ciclo ----------
    estados_ciclo = (
        db.query(EstadoCiclo)
        .filter(EstadoCiclo.idCiclo == ciclo.id)
        .order_by(EstadoCiclo.fechaInicio)
        .all()
    )
    intervalos_estado = _construir_intervalos(estados_ciclo, "fechaInicio", "fechaFin", "nombre", "EstadoCiclo")
    inicios_estado = [i[0] for i in intervalos_estado]
 
    if intervalos_estado:
        logger.info(
            f"[generar_informe_ciclo] EstadoCiclo: {len(intervalos_estado)} intervalos, "
            f"{intervalos_estado[0][0]} a {intervalos_estado[-1][1]}"
        )
    else:
        logger.warning(f"[generar_informe_ciclo] Ciclo {ciclo.id}: sin registros de EstadoCiclo.")
 
    # ---------- 5. Señales booleanas (SensoresIO) ----------
    ids_booleanos = [id_sensor for _, _, id_sensor in senales_booleanas]


    registros_booleanos = (
        db.query(SensoresIO)
        .filter(SensoresIO.idCiclo == ciclo.id, SensoresIO.idSensor.in_(ids_booleanos))
        .order_by(SensoresIO.fechaInicio)
        .all()
    )
    booleanos_por_sensor = defaultdict(list)
    for r in registros_booleanos:
        booleanos_por_sensor[r.idSensor].append(r)
 
    intervalos_booleanos = {}
    for clave, etiqueta, id_sensor in senales_booleanas:
        regs = booleanos_por_sensor.get(id_sensor, [])
        intervalos = _construir_intervalos(regs, "fechaInicio", "fechaFin", "valor", f"SensoresIO[{etiqueta}]")
        intervalos_booleanos[clave] = (intervalos, [i[0] for i in intervalos])
        if not intervalos:
            logger.warning(f"[generar_informe_ciclo] Ciclo {ciclo.id}: sin intervalos para señal booleana '{etiqueta}'.")
 
    logger.info(
        "[generar_informe_ciclo] Intervalos booleanos -> "
        + ", ".join(f"{etq}={len(intervalos_booleanos[clave][0])}" for clave, etq, _ in senales_booleanas)
    )
 
    # ---------- 6. Línea temporal maestra ----------
    # Incluye TODAS las fuentes temporales relevantes: analógicos,
    # inicio Y fin de cada estado, inicio Y fin de cada intervalo
    # booleano, y el inicio/fin oficial del ciclo.
    fechas_analogicas = set(fechas_agua) | set(fechas_ingreso) | set(fechas_producto) | set(fechas_nivel)
 
    fechas_estado = set()
    for inicio, fin, _ in intervalos_estado:
        fechas_estado.add(inicio)
        if fin is not None:
            fechas_estado.add(fin)
 
    fechas_booleanas = set()
    for clave, _, _ in senales_booleanas:
        for inicio, fin, _ in intervalos_booleanos[clave][0]:
            fechas_booleanas.add(inicio)
            if fin is not None:
                fechas_booleanas.add(fin)
 
    fechas_ciclo = set()
    if ciclo.fecha_inicio is not None:
        fechas_ciclo.add(ciclo.fecha_inicio)
    if ciclo.fecha_fin is not None:
        fechas_ciclo.add(ciclo.fecha_fin)
 
    todas_las_fechas = fechas_analogicas | fechas_estado | fechas_booleanas | fechas_ciclo
    cantidad_antes_dedup = (
        len(fechas_agua) + len(fechas_ingreso) + len(fechas_producto) + len(fechas_nivel)
        + sum(2 if fin is not None else 1 for _, fin, _ in intervalos_estado)
        + sum(2 if fin is not None else 1 for clave, _, _ in senales_booleanas for _, fin, _ in intervalos_booleanos[clave][0])
        + len(fechas_ciclo)
    )
    duplicados_eliminados = max(0, cantidad_antes_dedup - len(todas_las_fechas))
 
    timeline = sorted(todas_las_fechas)
 
    # ---------- 7. Acotar al rango oficial del ciclo ----------
    descartados = 0
    if ciclo.fecha_inicio is not None and ciclo.fecha_fin is not None:
        timeline_filtrado = []
        for ts in timeline:
            if ciclo.fecha_inicio <= ts <= ciclo.fecha_fin:
                timeline_filtrado.append(ts)
            else:
                descartados += 1
                logger.warning(
                    f"[generar_informe_ciclo] Timestamp {ts} fuera del rango oficial del "
                    f"ciclo [{ciclo.fecha_inicio} - {ciclo.fecha_fin}]. Se descarta."
                )
        timeline = timeline_filtrado
 
    logger.info(
        f"[generar_informe_ciclo] Timeline final: {len(timeline)} timestamps "
        f"({timeline[0] if timeline else 'N/A'} a {timeline[-1] if timeline else 'N/A'}), "
        f"{duplicados_eliminados} duplicados eliminados, {descartados} descartados por estar fuera del ciclo."
    )
 
    if not timeline:
        # No se inventan datos: si no hay NINGUNA fuente temporal disponible
        # (ni analógicos, ni estados, ni booleanos, ni fechas de ciclo), no
        # hay nada que reportar.
        logger.error(f"[generar_informe_ciclo] Ciclo {ciclo.id}: no se encontró ninguna fuente temporal de datos.")
        raise ValueError("El ciclo no posee ningún registro temporal (analógico, de estado o de E/S) para generar el reporte")
 
    # ---------- 8. Construcción de filas con forward-fill ----------
    filas = []
    for ts in timeline:
        estado_vigente = _valor_vigente_intervalo(intervalos_estado, inicios_estado, ts)
 
        # Regla del requisito #6: en el instante exacto de cierre oficial del
        # ciclo, si el ciclo está FINALIZADO/CANCELADO, ese estado tiene
        # prioridad sobre el último tramo de EstadoCiclo (que puede no
        # reflejar el cierre real, p. ej. si quedó en "OPERACIONAL").
        if (
            ciclo.fecha_fin is not None and ts == ciclo.fecha_fin
            and ciclo.estadoMaquina in ("FINALIZADO", "CANCELADO")
        ):
            estado_vigente = ciclo.estadoMaquina
 
        if estado_vigente is None:
            estado_vigente = "INDETERMINADO"
 
        # Mantener el orden de las cuatro columnas booleanas según el
        # tipo de equipo. Así cocina conserva su layout y enfriador usa
        # Vapor limpieza / Amoníaco sin depender de claves de cocina.
        valores_bool_ordenados = []
        for clave, _, _ in senales_booleanas:
            intervalos, inicios = intervalos_booleanos[clave]
            valor = _valor_vigente_intervalo(intervalos, inicios, ts)
            valores_bool_ordenados.append(_booleano_a_texto(valor))
 
        fila = [
            ts,
            estado_vigente,
            *valores_bool_ordenados,
            _valor_vigente_ffill(fechas_ingreso, valores_ingreso, ts),
            _valor_vigente_ffill(fechas_agua, valores_agua, ts),
            _valor_vigente_ffill(fechas_producto, valores_producto, ts),
            _valor_vigente_ffill(fechas_nivel, valores_nivel, ts),
        ]
        filas.append(fila)
 
    logger.info(
        f"[generar_informe_ciclo] Última fila generada: fecha={filas[-1][0]} "
        f"estado={filas[-1][1]} temp_agua={filas[-1][7]}"
    )
 
    # ---------- 9. Construcción del Excel a partir de la plantilla ----------
    if not _TEMPLATE_INFORME_CICLO_PATH.exists():
        logger.error(f"[generar_informe_ciclo] Plantilla no encontrada en {_TEMPLATE_INFORME_CICLO_PATH}")
        raise FileNotFoundError(f"No se encontró la plantilla en {_TEMPLATE_INFORME_CICLO_PATH}")
 
    wb = load_workbook(_TEMPLATE_INFORME_CICLO_PATH)
 
    if _HOJA_SENSORES_A_QUITAR in wb.sheetnames:
        del wb[_HOJA_SENSORES_A_QUITAR]
 
    ws = wb[_HOJA_DETALLE]
 
    fecha_inicio_str = ciclo.fecha_inicio.strftime("%Y-%m-%d %H:%M:%S") if ciclo.fecha_inicio else "N/A"
    fecha_fin_str = ciclo.fecha_fin.strftime("%Y-%m-%d %H:%M:%S") if ciclo.fecha_fin else "N/A"
    estado_final = ciclo.estadoMaquina or "N/A"
 
    ws["A1"] = (
        f"ID CICLO: {ciclo.id} | LOTE: {ciclo.lote or 'N/A'} | "
        f"INICIO: {fecha_inicio_str} | FIN: {fecha_fin_str} | ESTADO: {estado_final}"
    )
    ws["B3"] = nombre_equipo
    ws["B4"] = nombre_receta
    ws["B5"] = tipo_fin
    ws["B6"] = ciclo.cantidadTorres
    ws["B7"] = ciclo.tiempoTranscurrido
    ws["B8"] = ciclo.peso

    # Las columnas C..F corresponden siempre a las cuatro señales
    # booleanas seleccionadas para el tipo de equipo.
    for columna, (_, etiqueta, _) in enumerate(senales_booleanas, start=3):
        ws.cell(row=_HEADER_ROW, column=columna, value=etiqueta)
 
    estilos_columna = []
    for c in range(1, _TABLE_COLS + 1):
        celda_origen = ws.cell(row=_DATA_START_ROW, column=c)
        estilos_columna.append({
            "font": _copy.copy(celda_origen.font),
            "fill": _copy.copy(celda_origen.fill),
            "border": _copy.copy(celda_origen.border),
            "alignment": _copy.copy(celda_origen.alignment),
            "number_format": celda_origen.number_format,
        })
 
    max_row_actual = ws.max_row
    if max_row_actual >= _DATA_START_ROW:
        ws.delete_rows(_DATA_START_ROW, max_row_actual - _DATA_START_ROW + 1)
 
    DATE_FMT = "yyyy-mm-dd hh:mm:ss"
    TEMP_FMT = "0.00"
 
    for i, fila in enumerate(filas):
        row_num = _DATA_START_ROW + i
        for col_idx, valor in enumerate(fila, start=1):
            celda = ws.cell(row=row_num, column=col_idx, value=valor)
            _aplicar_estilo(celda, estilos_columna[col_idx - 1])
            if col_idx == 1:
                celda.number_format = DATE_FMT
            elif col_idx in (7, 8, 9, 10) and isinstance(valor, (int, float)):
                celda.number_format = TEMP_FMT
    last_row = _DATA_START_ROW + len(filas) - 1
 
    if ws.tables:
        nombre_tabla = next(iter(ws.tables))
        tabla_actual = ws.tables[nombre_tabla]
        estilo_tabla = tabla_actual.tableStyleInfo
        del ws.tables[nombre_tabla]
        nueva_tabla = Table(displayName=nombre_tabla, ref=f"A{_HEADER_ROW}:J{last_row}")
        nueva_tabla.tableStyleInfo = estilo_tabla
        ws.add_table(nueva_tabla)
 
    logger.info(f"[generar_informe_ciclo] Tabla Excel actualizada a rango A{_HEADER_ROW}:J{last_row}")
 
    excel_stream = BytesIO()
    wb.save(excel_stream)
    wb.close()
    excel_stream.seek(0)
 
    return excel_stream
 
def obtener_historico_alarmas(fecha_inicio, fecha_fin, session):
    q = session.query(HistoricoAlarma)

    if fecha_inicio and fecha_fin:
        q = q.filter(
            HistoricoAlarma.fechaInicio >= fecha_inicio,
            HistoricoAlarma.fechaInicio <= fecha_fin
        )
    else:
        fecha_actual = datetime.now()
        q = q.filter(HistoricoAlarma.fechaInicio <= fecha_actual)

    historico_alarmas = q.all()
    if not historico_alarmas:
        return []

    ids_l1 = {h.idAlarma for h in historico_alarmas if h.tipoLinea == "L1"}
    ids_l2 = {h.idAlarma for h in historico_alarmas if h.tipoLinea != "L1"}

    rows_by_id_l1 = {}
    rows_by_id_l2 = {}

    if ids_l1:
        alarma_l1 = session.query(Alarmas).filter(Alarmas.id.in_(ids_l1)).all()
        rows_by_id_l1 = {
            row.id: {"nombre": row.nombre, "descripcion": row.descripcion, "tipoAlarma": row.tipoAlarma}
            for row in alarma_l1
        }

    if ids_l2:
        alarma_l2 = session.query(AlarmasL2).filter(AlarmasL2.id.in_(ids_l2)).all()
        rows_by_id_l2 = {
            row.id: {"nombre": row.nombre, "descripcion": row.descripcion, "tipoAlarma": row.tipoAlarma}
            for row in alarma_l2
        }

    datos = []
    for h in historico_alarmas:
        a = rows_by_id_l1.get(h.idAlarma) if h.tipoLinea == "L1" else rows_by_id_l2.get(h.idAlarma)
        if h.fechaInicio is not None and h.fechaFin is not None:
            delta = h.fechaFin - h.fechaInicio
            total_segundos = int(delta.total_seconds())
            horas = total_segundos // 3600
            minutos = (total_segundos % 3600) // 60
            segundos = total_segundos % 60
            tiempo_transcurrido = f"{horas:02}:{minutos:02}:{segundos:02}"
        else:
            tiempo_transcurrido = None
        datos.append({
            "id": h.id,
            "nombre_alarmas": a.get("nombre", "") if a else "",
            "descripcion": a.get("descripcion", "") if a else "",
            "tipo_alarma": a.get("tipoAlarma", "") if a else "",
            "seccion": h.tipoLinea,
            "fecha_inicio": h.fechaInicio,
            "fecha_fin": h.fechaFin,
            "tiempo_transcurrido": tiempo_transcurrido
        })

    return datos

def generar_reporte_alarmas_descarga(session, fecha_inicio_dt, fecha_fin_dt):
    datos = obtener_historico_alarmas(fecha_inicio_dt, fecha_fin_dt, session)

    wb = load_workbook(_TEMPLATE_PATH)
    ws = wb["informe alarmas"]

    # Fechas de encabezado (mantiene el formato dd/mm/yyyy de la plantilla)
    ws["B2"] = fecha_inicio_dt
    ws["B3"] = fecha_fin_dt

    # ── Capturar estilos de la fila placeholder (ahora es fila 6) ──
    _NCOLS = 7
    col_styles = []
    for c in range(1, _NCOLS + 1):
        src = ws.cell(row=6, column=c)          # ← fila 6
        col_styles.append({
            "font":          _copy.copy(src.font),
            "fill":          _copy.copy(src.fill),
            "border":        _copy.copy(src.border),
            "alignment":     _copy.copy(src.alignment),
            "number_format": src.number_format,
        })

    ws.delete_rows(6)                            # ← eliminar fila 6

    DATA_START = 6                               # ← datos desde fila 6
    DATE_FMT   = "DD/MM/YYYY HH:MM:SS"

    if not datos:
        for c in range(1, _NCOLS + 1):
            cell = ws.cell(row=DATA_START, column=c, value=None)
            _apply_style(cell, col_styles[c - 1])
        last_row = DATA_START
    else:
        for i, row_data in enumerate(datos):
            row_num = DATA_START + i
            fi = row_data.get("fecha_inicio")
            ff = row_data.get("fecha_fin")
            row_values = [
                row_data.get("id"),
                row_data.get("nombre_alarmas"),
                row_data.get("tipo_alarma"),
                row_data.get("seccion"),
                fi,
                ff,
                _calc_elapsed(fi, ff),
            ]
            for col_idx, value in enumerate(row_values, start=1):
                cell = ws.cell(row=row_num, column=col_idx, value=value)
                _apply_style(cell, col_styles[col_idx - 1])
                if col_idx in (5, 6) and isinstance(value, datetime):
                    cell.number_format = DATE_FMT

        last_row = DATA_START + len(datos) - 1

    _update_table(ws, new_last_row=last_row)

    output = BytesIO()
    wb.save(output)
    wb.close()
    output.seek(0)
    return output

def _apply_style(cell, style: dict):
    cell.font          = _copy.copy(style["font"])
    cell.fill          = _copy.copy(style["fill"])
    cell.border        = _copy.copy(style["border"])
    cell.alignment     = _copy.copy(style["alignment"])
    cell.number_format = style["number_format"]

def _calc_elapsed(fi, ff) -> str:
    if fi is None or ff is None:
        return ""
    try:
        total_s = int((ff - fi).total_seconds())
        if total_s < 0:
            return ""
        h, rem = divmod(total_s, 3600)
        m, s   = divmod(rem, 60)
        return f"{h:02d}:{m:02d}:{s:02d}"
    except Exception:
        return ""

def _update_table(ws, new_last_row: int):
    if not ws.tables:
        return
    tbl_name = next(iter(ws.tables))
    tbl      = ws.tables[tbl_name]
    style    = tbl.tableStyleInfo          # puede ser None, lo preservamos igual
    del ws.tables[tbl_name]
    new_tbl = Table(displayName=tbl_name, ref=f"A5:G{new_last_row}")  # ← A5
    new_tbl.tableStyleInfo = style
    ws.add_table(new_tbl)

def generar_reporte_productividad(id_equipo, fecha_inicio, fecha_fin, db):
    fecha_inicio = datetime.combine(fecha_inicio, datetime.min.time())
    fecha_fin = datetime.combine(fecha_fin, datetime.max.time())
    lista_ciclos: list[Ciclo] = [] 
    #linea1 = [7,8,9,10]
    #linea2 = [11,12,13,14]
    linea1 = [1,2,3]
    linea2 = [4,5,6]

    nombre_maquina = ""
    def buscarReceta(id_receta):
        receta = (
            db.query(Receta)
            .filter(id_receta == Receta.id)
            .first()
        )
        return receta.nombre
    
    def bucarNombreEquipo(id_maquina):
        maquina = (db.query(Equipo).filter(id_maquina == Equipo.id).first())
        return maquina.nombre

    if id_equipo == 15:
        lista_ciclos= (
            db.query(Ciclo)
            .filter(Ciclo.idEquipo.in_(linea1))
            .filter(Ciclo.fecha_fin.between(fecha_inicio, fecha_fin))
            .all()
        )
        nombre_maquina = "Linea 1"
    if id_equipo == 16:
        lista_ciclos= (
            db.query(Ciclo)
            .filter(Ciclo.idEquipo.in_(linea2))
            .filter(Ciclo.fecha_fin.between(fecha_inicio, fecha_fin))
            .all()
        )
        nombre_maquina = "Linea 2"
    if id_equipo <= 14 and id_equipo != 0:
        lista_ciclos = (
            db.query(Ciclo)
            .filter(Ciclo.fecha_fin.between(fecha_inicio, fecha_fin))
            .filter(Ciclo.idEquipo == id_equipo)
            .all()
        )
        nombre_maquina = bucarNombreEquipo(id_equipo)
    if id_equipo == 0:
        lista_ciclos = (
            db.query(Ciclo)
            .filter(Ciclo.fecha_fin.between(fecha_inicio, fecha_fin))
            .all()
        )
        nombre_maquina = "Completo"

    #CONSTRUCCION DEL ARCHIVO XLMS
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Informe de productividad"
    logoPath = "cremona.png"
    img = Image(logoPath)
    img.width = 280
    img.height = 70
    sheet.add_image(img, 'D1')

    sheet.append(["Lista ciclos finalizados correctamente"])
    producto_cell = sheet.cell(row=sheet.max_row, column=1)
    producto_cell.font = Font(bold= True, size=20)

    sheet.append(["Fecha Inicio:", fecha_inicio.strftime("%Y-%m-%d")])
    fechaInicio_cell = sheet.cell(row=sheet.max_row, column=1)
    fechaInicio_cell.font = Font(bold=True, size=12)
    sheet.append(["Fecha Fin:", fecha_fin.strftime("%Y-%m-%d")])
    fechaFin_cell = sheet.cell(row=sheet.max_row, column=1)
    fechaFin_cell.font = Font(bold=True, size=12)

    sheet.append([f"Buscar: {nombre_maquina}"])
    producto_cell = sheet.cell(row=sheet.max_row, column=5)
    producto_cell.font = Font(bold= True, size=16)

    headers = ["ID_CICLO", "ESTADO_CICLO", "CANTIDAD_TORRE", "LOTE", "TIEMPO_TRANSCURRIDO", "FECHA_INICIO", "FECHA_FIN", "EQUIPO", "PESO","RECETA"]
    sheet.append(headers)
    header_fill = PatternFill(start_color="145f82", end_color="145f82", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)  

    for col in range(1, len(headers) + 1):
        cell = sheet.cell(row=sheet.max_row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    
    start_row = sheet.max_row + 1

    resultado_fila = []

    for elem in lista_ciclos:
        item_equipo = bucarNombreEquipo(elem.idEquipo)
        item_receta = buscarReceta(elem.idReceta)
        resultado_fila.append([
            elem.id,
            elem.estadoMaquina,
            elem.cantidadTorres,
            elem.lote,
            elem.tiempoTranscurrido,
            elem.fecha_inicio,
            elem.fecha_fin,
            item_equipo,
            elem.peso,
            item_receta
        ])

    for item in resultado_fila:
        sheet.append(item)
    
    end_row = sheet.max_row
    start_col = 1
    end_col = len(headers)
    table_range = f"{sheet.cell(row=start_row -1, column=start_col).coordinate}:{sheet.cell(row=end_row, column=end_col).coordinate}"
    table_nombre = "GraficosHistorico"
    tabla = Table(displayName=table_nombre, ref=table_range)
    style = TableStyleInfo(showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tabla.tableStyleInfo = style

    # Agregar la tabla a la hoja
    sheet.add_table(tabla)
    sheet.append([])

    for col in sheet.columns:
        max_length = 0
        column_letter = col[0].column_letter
        for cell in col:
            try:
                max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        sheet.column_dimensions[column_letter].width = max_length + 2

    excel_stream = BytesIO()
    workbook.save(excel_stream)
    workbook.close() 
    excel_stream.seek(0)  

    return excel_stream

def productividad_equipo(db, fecha_inicio, fecha_fin, id_equipo):
    fecha_inicio = datetime.combine(fecha_inicio, datetime.min.time())
    fecha_fin = datetime.combine(fecha_fin, datetime.max.time())

    respuesta = {}
    lista_receta = defaultdict(list)
    lista_ciclos: list[Ciclo] = [] 
    #linea1 = [7,8,9,10]
    #linea2 = [11,12,13,14]
    linea1 = [1,2,3]
    linea2 = [4,5,6]

    if id_equipo == 15:
        lista_ciclos= (
            db.query(Ciclo)
            .filter(Ciclo.idEquipo.in_(linea1))
            .filter(Ciclo.fecha_fin.between(fecha_inicio, fecha_fin))
            .all()
        )
    if id_equipo == 16:
        lista_ciclos= (
            db.query(Ciclo)
            .filter(Ciclo.idEquipo.in_(linea2))
            .filter(Ciclo.fecha_fin.between(fecha_inicio, fecha_fin))
            .all()
        )

    if id_equipo <= 14 and id_equipo != 0:
        lista_ciclos = (
            db.query(Ciclo)
            .filter(Ciclo.fecha_fin.between(fecha_inicio, fecha_fin))
            .filter(Ciclo.idEquipo == id_equipo)
            .all()
        )
    if id_equipo == 0:
        lista_ciclos = (
            db.query(Ciclo)
            .filter(Ciclo.fecha_fin.between(fecha_inicio, fecha_fin))
            .all()
        )

    total_peso = 0
    cantidad_ciclos_correctos = 0
    cantidad_ciclos_incorectos = 0

    for elem in lista_ciclos:
        total_peso+= elem.peso
        if elem.estadoMaquina == "FINALIZADO":
            cantidad_ciclos_correctos+= 1
        if elem.estadoMaquina == "CANCELADO":
            cantidad_ciclos_incorectos+= 1
        
        lista_receta[elem.idReceta].append(elem)

    agrupado_dict = dict(lista_receta)
    list_productos = []

    for id_receta, items in lista_receta.items():

        receta = {
            "nombre_receta": None,
            "capacidad_receta": 0, 
            "cantidad_ciclos": 0, 
        }
        db_receta = (
            db.query(Receta)
            .filter(Receta.id == id_receta)
            .first())
        for item in items:
            receta["nombre_receta"] = db_receta.nombre
            receta["capacidad_receta"] += int(item.peso)
            receta["cantidad_ciclos"] += 1

        list_productos.append(receta)

    respuesta["ciclos_realizados"] = cantidad_ciclos_incorectos + cantidad_ciclos_correctos
    respuesta["produccion_total"] = round(total_peso / 1000, 2)
    respuesta["ciclos_correctos"] = cantidad_ciclos_correctos
    respuesta["ciclos_incorrectos"] = cantidad_ciclos_incorectos
    respuesta["productos_realizados"] = list_productos
 
    return respuesta

def obtener_valor_sensores(db, id_ciclo:int, id_sensor:int, tipo):
    try:
        if tipo == "MAX":
            resultado = (
                db.query(SensoresAA)
                .filter(SensoresAA.idCiclo == id_ciclo, SensoresAA.idSensor == id_sensor)
                .order_by(SensoresAA.valor.desc())
                .limit(1)
                .first()
            )
            return resultado.valor if resultado else 0

        if tipo == "MIN":
            resultado = (
                db.query(SensoresAA)
                .filter(SensoresAA.idCiclo == id_ciclo, SensoresAA.idSensor == id_sensor)
                .order_by(SensoresAA.valor.asc())
                .limit(1)
                .first()
            )
            return resultado.valor if resultado else 0

        return 0
        
    except Exception as e:
        logger.error(f"Error obteniendo valor de sensor {id_sensor} para ciclo {id_ciclo}: {e}")
        return 0

def obtener_datos_graficos(db, id_ciclo:int):
    lista_sensores_data = {}

    ciclo: list[Ciclo] = []
    if id_ciclo != 0: 
        ciclo = (db.query(Ciclo)
                .filter(Ciclo.id == id_ciclo)
                .first()
                )
    if id_ciclo == 0:
        ciclo = db.query(Ciclo).order_by(Ciclo.id.desc()).first()  

    if not ciclo:
        logger.info("No se encontró el ciclo en la BDD")
        raise ValueError("Ciclo no encontrado")


    lista_sensores = (
        db.query(Sensores)
    ).all()

    estado_ciclo = (
        db.query(EstadoCiclo)
        .filter(EstadoCiclo.idCiclo == ciclo.id)
    )
    receta = (
        db.query(Receta)
        .filter(ciclo.idReceta == Receta.id)
        .first()
    )

    general = {}
    general["id_ciclo"] = ciclo.id
    general["ciclo_lote"] = ciclo.lote
    general["tiempo_transcurrido"] = ciclo.tiempoTranscurrido
    general["fecha_inicio"] = ciclo.fecha_inicio
    general["fecha_fin"] = ciclo.fecha_fin
    general["receta"] = receta.nombre

    for sensor in lista_sensores:
        if sensor.nombre == "Temperatura agua":
            temp_agua = (
                db.query(SensoresAA)
                .filter(SensoresAA.idSensor == sensor.id)
                .filter(SensoresAA.idCiclo == ciclo.id)
                .all()
            )
            print(f"Cantidad Filas de registros en BDD: {len(temp_agua)}")
            general["temp_agua_max"] = obtener_valor_sensores(db,id_ciclo, sensor.id, "MAX")
            general["temp_agua_min"] = obtener_valor_sensores(db,id_ciclo, sensor.id, "MIN")
            lista_sensores_data[sensor.nombre] = temp_agua
        
        if sensor.nombre == "Temperatura producto":
            temp_producto = (
                db.query(SensoresAA)
                .filter(SensoresAA.idSensor == sensor.id)
                .filter(SensoresAA.idCiclo == ciclo.id)
                .all()
            )
            print(f"Cantidad Filas de registros en BDD: {len(temp_producto)}")
            general["temp_producto_max"] = obtener_valor_sensores(db,id_ciclo, sensor.id, "MAX")
            general["temp_producto_min"] = obtener_valor_sensores(db,id_ciclo, sensor.id, "MIN")
            lista_sensores_data[sensor.nombre] = temp_producto
        if sensor.nombre == "Nivel agua":
            nivel_agua = (
                db.query(SensoresAA)
                .filter(SensoresAA.idSensor == sensor.id)
                .filter(SensoresAA.idCiclo == ciclo.id)
                .all()
            )
            print(f"Cantidad Filas de registros en BDD: {len(nivel_agua)}")
            general["nivel_agua_max"] = obtener_valor_sensores(db,id_ciclo, sensor.id, "MAX")
            general["nivel_agua_min"] = obtener_valor_sensores(db,id_ciclo, sensor.id, "MIN")

            lista_sensores_data[sensor.nombre] = nivel_agua

    lista_sensores_data["general"] = general

    return lista_sensores_data



