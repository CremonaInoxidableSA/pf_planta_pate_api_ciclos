import math
import random
import socket
import time
from datetime import datetime

from opcua import ua, Server

ESTADO_PREOPERATIVO = 1
ESTADO_OPERATIVO = 2
ESTADO_PAUSADO = 3
ESTADO_INACTIVO = 4
ESTADO_CANCELADO = 5
ESTADO_FINALIZADO = 6
ESTADO_LIMPIEZA = 7

PASO_PREOPERATIVO = 1
PASO_OPERATIVO = 2
PASO_FINALIZADO = 3

CICLO_COCCION_MAX_STEP = 180
CICLO_ENFRIAMIENTO_RESET_STEP = 240

COCCION_START_OFFSETS = [0, 60, 160, 80, 30, 120]
ENFRIAMIENTO_START_OFFSETS = [10, 20, 10, 30, 40, 45, 0, 5]


ALARMAS_L1_TOTAL = 100
ALARMAS_L1_UPDATE_COUNT = 10
ALARMAS_L1_UPDATE_INTERVAL = 30  # segundos
ALARMAS_L1_PULSE_ON_SECONDS = 10  # cuánto tiempo queda True

ALARMAS_L2_TOTAL = 100
ALARMAS_L2_UPDATE_COUNT = 10
ALARMAS_L2_UPDATE_INTERVAL = 60  # segundos
ALARMAS_L2_PULSE_ON_SECONDS = 30

def get_local_ip():
    """
    Intenta obtener la IP local de la PC para publicar el endpoint.
    """
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip



def add_var(parent, idx, name, value, variant_type, writable=True):
    node = parent.add_variable(idx, name, ua.Variant(value, variant_type))
    if writable:
        node.set_writable()
    return node



def set_value(node, value, variant_type):
    node.set_value(ua.Variant(value, variant_type))



def add_array_of_scalars(parent, idx, array_name, values, variant_type):
    """
    Crea un objeto que actúa como array visual:
    TEMP_AGUA
      ├── [0]
      ├── [1]
      ├── [2]
      └── [3]
    """
    arr_obj = parent.add_object(idx, array_name)
    nodes = []
    for i, value in enumerate(values):
        n = add_var(arr_obj, idx, f"[{i}]", value, variant_type)
        nodes.append(n)
    return arr_obj, nodes



def build_alarmas_l1(parent, idx, total_alarmas=ALARMAS_L1_TOTAL):

    nodes = []
    initial_values = [False] * total_alarmas

    for i in range(1, total_alarmas + 1):
        node = add_var(parent, idx, f"[{i}]", False, ua.VariantType.Boolean)
        nodes.append(node)

    state = {
        "values": initial_values,
        "last_pulse_start": time.time(),  # arranca el contador
        "pulse_active": False,
        "pulse_end": 0.0,
    }

    return nodes, state



def update_alarmas_l1(
    nodes,
    state,
    update_count=ALARMAS_L1_UPDATE_COUNT,
    interval=ALARMAS_L1_UPDATE_INTERVAL,
    pulse_on_seconds=ALARMAS_L1_PULSE_ON_SECONDS,
):

    now = time.time()
    limite = min(update_count, len(nodes), len(state["values"]))

    # 1) Si el pulso está activo, ver si ya hay que apagar
    if state["pulse_active"]:
        if now >= state["pulse_end"]:
            for i in range(limite):
                state["values"][i] = False
                set_value(nodes[i], False, ua.VariantType.Boolean)

            state["pulse_active"] = False
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"[{timestamp}] ALARMAS L1 | Pulso OFF ({limite} elementos)")
        return

    # 2) Si NO hay pulso activo, ver si toca iniciar uno nuevo
    if now - state["last_pulse_start"] >= interval:
        for i in range(limite):
            state["values"][i] = True
            set_value(nodes[i], True, ua.VariantType.Boolean)

        state["pulse_active"] = True
        state["pulse_end"] = now + pulse_on_seconds
        state["last_pulse_start"] = now

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] ALARMAS L1 | Pulso ON por {pulse_on_seconds}s ({limite} elementos)")


def build_alarmas_l2(parent, idx, total_alarmas=ALARMAS_L2_TOTAL):

    nodes = []
    initial_values = [False] * total_alarmas

    for i in range(1, total_alarmas + 1):
        node = add_var(parent, idx, f"[{i}]", False, ua.VariantType.Boolean)
        nodes.append(node)

    state = {
        "values": initial_values,
        "last_pulse_start": time.time(),  # arranca el contador
        "pulse_active": False,
        "pulse_end": 0.0,
    }

    return nodes, state

def update_alarmas_l2(
    nodes,
    state,
    update_count=ALARMAS_L2_UPDATE_COUNT,
    interval=ALARMAS_L2_UPDATE_INTERVAL,
    pulse_on_seconds=ALARMAS_L2_PULSE_ON_SECONDS,
):

    now = time.time()
    limite = min(update_count, len(nodes), len(state["values"]))

    # 1) Si el pulso está activo, ver si ya hay que apagar
    if state["pulse_active"]:
        if now >= state["pulse_end"]:
            for i in range(limite):
                state["values"][i] = False
                set_value(nodes[i], False, ua.VariantType.Boolean)

            state["pulse_active"] = False
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"[{timestamp}] ALARMAS L2 | Pulso OFF ({limite} elementos)")
        return

    # 2) Si NO hay pulso activo, ver si toca iniciar uno nuevo
    if now - state["last_pulse_start"] >= interval:
        for i in range(limite):
            state["values"][i] = True
            set_value(nodes[i], True, ua.VariantType.Boolean)

        state["pulse_active"] = True
        state["pulse_end"] = now + pulse_on_seconds
        state["last_pulse_start"] = now

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] ALARMAS L2 | Pulso ON por {pulse_on_seconds}s ({limite} elementos)")

def bool_to_int16(value: bool) -> int:
    return 100 if value else 0



def build_lote(prefix: str, numero_equipo: int, ciclo_id: int):
    return f"{prefix}-{numero_equipo:02d}-C{ciclo_id:04d}"



def log_cycle_event(tipo_equipo: str, numero_equipo: int, lote: str, mensaje: str):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {tipo_equipo} {numero_equipo} | {lote} | {mensaje}")


def get_estado_nombre(estado: int) -> str:
    estados = {
        ESTADO_PREOPERATIVO: "PRE OPERACIONAL",
        ESTADO_OPERATIVO: "OPERACIONAL",
        ESTADO_PAUSADO: "PAUSADO",
        ESTADO_INACTIVO: "INACTIVO",
        ESTADO_CANCELADO: "CANCELADO",
        ESTADO_FINALIZADO: "FINALIZADO",
        ESTADO_LIMPIEZA: "LIMPIEZA",
    }
    return estados.get(estado, f"DESCONOCIDO({estado})")


def log_state_change(tipo_equipo: str, numero_equipo: int, lote: str, estado: int, paso: int, step: int):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    estado_nombre = get_estado_nombre(estado)
    print(
        f"[{timestamp}] {tipo_equipo} {numero_equipo} | {lote} | CAMBIO DE ESTADO -> {estado_nombre} | PASO={paso} | STEP={step}"
    )



def create_cocina_sim_state(numero_equipo: int, start_step: int):
    return {
        "numero_equipo": numero_equipo,
        "step": start_step % (CICLO_COCCION_MAX_STEP + 1),
        "cycle_id": 1,
        "lote": build_lote("COC", numero_equipo, 1),
        "start_logged_for_cycle": set(),
        "end_logged_for_cycle": set(),
        "last_estado": None,
    }



def create_enfriador_sim_state(numero_equipo: int, start_step: int):
    return {
        "numero_equipo": numero_equipo,
        "step": start_step % CICLO_ENFRIAMIENTO_RESET_STEP,
        "cycle_id": 1,
        "lote": build_lote("ENF", numero_equipo, 1),
        "start_logged_for_cycle": set(),
        "end_logged_for_cycle": set(),
        "last_estado": None,
    }


# =========================
# CONSTRUCCIÓN DE ESTRUCTURAS
# =========================
def build_recipe_item(parent, idx, item_name, recipe_index):
    item = parent.add_object(idx, item_name)
    nodes = {}

    nodes["PASOS"] = add_var(item, idx, "PASOS", 4, ua.VariantType.Int16)

    _, nodes["TEMP_AGUA"] = add_array_of_scalars(
        item, idx, "TEMP_AGUA", [70, 75, 80, 85], ua.VariantType.Int16
    )

    _, nodes["TEMP_PRODUCTO"] = add_array_of_scalars(
        item, idx, "TEMP PRODUCTO", [65, 70, 75, 80], ua.VariantType.Int16
    )

    _, nodes["TIEMPO_CORTE"] = add_array_of_scalars(
        item, idx, "TIEMPO CORTE", [10, 12, 14, 16], ua.VariantType.Int16
    )

    _, nodes["TIPO_CORTE"] = add_array_of_scalars(
        item, idx, "TIPO CORTE", [True, False, True, False], ua.VariantType.Boolean
    )

    nodes["NOMBRE"] = add_var(
        item, idx, "NOMBRE", f"RECETA_{recipe_index:02}", ua.VariantType.String
    )
    nodes["TIEMPO_CORTE_ENFRIADO"] = add_var(
        item, idx, "TIEMPO CORTE ENFRIADO", 20, ua.VariantType.Int16
    )
    nodes["TEMP_CORTE_ENFRIADO"] = add_var(
        item, idx, "TEMP CORTE ENFRIADO", 8, ua.VariantType.Int16
    )
    nodes["TIPO_CORTE_ENFRIADO"] = add_var(
        item, idx, "TIPO CORTE ENFRIADO", True, ua.VariantType.Boolean
    )
    nodes["PESO_X_TORRE"] = add_var(
        item, idx, "PESO X TORRE", 25, ua.VariantType.Int16
    )
    nodes["ON_OFF_ALARMA"] = add_var(
        item, idx, "ON/OFF ALARMA", False, ua.VariantType.Boolean
    )
    nodes["TIEMPO_PARA_ALARMA"] = add_var(
        item, idx, "TIEMPO PARA ALARMA", 120, ua.VariantType.Int16
    )

    return nodes



def build_cocina_unit(parent, idx, item_name, numero_equipo):
    item = parent.add_object(idx, item_name)
    nodes = {}

    nodes["FILTRO_SUCCION_AGUA"] = add_var(
        item, idx, "FILTRO_SUCCION_AGUA", False, ua.VariantType.Boolean
    )
    nodes["CARGA_AGUA"] = add_var(
        item, idx, "CARGA_AGUA", False, ua.VariantType.Boolean
    )
    nodes["VAPOR_SERPENTINA_ACC"] = add_var(
        item, idx, "VAPOR_SERPENTINA_ACC", False, ua.VariantType.Boolean
    )
    nodes["VAPOR_VIVO_ACC"] = add_var(
        item, idx, "VAPOR_VIVO_ACC", False, ua.VariantType.Boolean
    )
    nodes["BOMBA_CENTRIFUGA"] = add_var(
        item, idx, "BOMBA_CENTRIFUGA", False, ua.VariantType.Boolean
    )
    nodes["CICLO_TIPO_FIN"] = add_var(
        item, idx, "CICLO_TIPO_FIN", False, ua.VariantType.Boolean
    )

    nodes["TEMP_AGUA"] = add_var(item, idx, "TEMP_AGUA", 20.0, ua.VariantType.Float)
    nodes["TEMP_INGRESO"] = add_var(item, idx, "TEMP_INGRESO", 18.0, ua.VariantType.Float)
    nodes["TEMP_PRODUCTO"] = add_var(item, idx, "TEMP_PRODUCTO", 15.0, ua.VariantType.Float)
    nodes["NIVEL_AGUA"] = add_var(item, idx, "NIVEL_AGUA", 50.0, ua.VariantType.Float)

    nodes["VAPOR_SERPENTINA"] = add_var(
        item, idx, "VAPOR_SERPENTINA", 0, ua.VariantType.Int16
    )
    nodes["VAPOR_VIVO"] = add_var(
        item, idx, "VAPOR_VIVO", 0, ua.VariantType.Int16
    )
    nodes["PASO_ACTUAL"] = add_var(
        item, idx, "PASO_ACTUAL", 0, ua.VariantType.Int16
    )
    nodes["NUMERO_RECETA"] = add_var(
        item, idx, "NUMERO_RECETA", 0, ua.VariantType.Int16
    )
    nodes["CANTIDAD_TORRES"] = add_var(
        item, idx, "CANTIDAD_TORRES", 0, ua.VariantType.Int16
    )
    nodes["ESTADO_EQUIPO"] = add_var(
        item, idx, "ESTADO_EQUIPO", 0, ua.VariantType.Int16
    )
    nodes["PESO_PRODUCTO"] = add_var(
        item, idx, "PESO_PRODUCTO", 0, ua.VariantType.Int16
    )
    nodes["NUMERO_EQUIPO"] = add_var(
        item, idx, "NUMERO_EQUIPO", numero_equipo, ua.VariantType.Int16
    )

    nodes["TIEMPO_TRANS"] = add_var(
        item, idx, "TIEMPO_TRANS", 0, ua.VariantType.Int32
    )

    nodes["LOTE_CICLO"] = add_var(
        item, idx, "LOTE_CICLO", "", ua.VariantType.String
    )

    return nodes



def build_enfriador_unit(parent, idx, item_name, numero_equipo):
    item = parent.add_object(idx, item_name)
    nodes = {}

    nodes["FILTRO_SUCCION_AGUA"] = add_var(
        item, idx, "FILTRO_SUCCION_AGUA", False, ua.VariantType.Boolean
    )
    nodes["CARGA_AGUA"] = add_var(
        item, idx, "CARGA_AGUA", False, ua.VariantType.Boolean
    )
    nodes["AMONIACO_ACC"] = add_var(
        item, idx, "AMONIACO_ACC", False, ua.VariantType.Boolean
    )
    nodes["VAPOR_LIMPIEZA_ACC"] = add_var(
        item, idx, "VAPOR_LIMPIEZA_ACC", False, ua.VariantType.Boolean
    )
    nodes["BOMBA_CENTRIFUGA"] = add_var(
        item, idx, "BOMBA_CENTRIFUGA", False, ua.VariantType.Boolean
    )
    nodes["CICLO_TIPO_FIN"] = add_var(
        item, idx, "CICLO_TIPO_FIN", False, ua.VariantType.Boolean
    )

    nodes["TEMP_AGUA"] = add_var(item, idx, "TEMP_AGUA", 10.0, ua.VariantType.Float)
    nodes["TEMP_INGRESO"] = add_var(item, idx, "TEMP_INGRESO", 8.0, ua.VariantType.Float)
    nodes["TEMP_PRODUCTO"] = add_var(item, idx, "TEMP_PRODUCTO", 6.0, ua.VariantType.Float)
    nodes["NIVEL_AGUA"] = add_var(item, idx, "NIVEL_AGUA", 40.0, ua.VariantType.Float)

    nodes["AMONIACO"] = add_var(item, idx, "AMONIACO", 0, ua.VariantType.Int16)
    nodes["VAPOR_LIMPIEZA"] = add_var(item, idx, "VAPOR_LIMPIEZA", 0, ua.VariantType.Int16)
    nodes["PASO_ACTUAL"] = add_var(item, idx, "PASO_ACTUAL", 0, ua.VariantType.Int16)
    nodes["NUMERO_RECETA"] = add_var(item, idx, "NUMERO_RECETA", 0, ua.VariantType.Int16)
    nodes["CANTIDAD_TORRES"] = add_var(item, idx, "CANTIDAD_TORRES", 0, ua.VariantType.Int16)
    nodes["ESTADO_EQUIPO"] = add_var(item, idx, "ESTADO_EQUIPO", 0, ua.VariantType.Int16)
    nodes["PESO_PRODUCTO"] = add_var(item, idx, "PESO_PRODUCTO", 0, ua.VariantType.Int16)
    nodes["NUMERO_EQUIPO"] = add_var(item, idx, "NUMERO_EQUIPO", numero_equipo, ua.VariantType.Int16)

    nodes["TIEMPO_TRANS"] = add_var(item, idx, "TIEMPO_TRANS", 0, ua.VariantType.Int32)

    nodes["LOTE_CICLO"] = add_var(item, idx, "LOTE_CICLO", "", ua.VariantType.String)

    return nodes


# =========================
# SIMULACIÓN DE VALORES MIGRADA DESDE opc_pf.py
# =========================
def resolve_coccion_phase(step: int):
    if step < 30:
        return ESTADO_PREOPERATIVO, PASO_PREOPERATIVO
    if step < 170:
        return ESTADO_OPERATIVO, PASO_OPERATIVO
    return ESTADO_FINALIZADO, PASO_FINALIZADO



def resolve_enfriamiento_phase(step: int):
    if step < 120:
        return ESTADO_PREOPERATIVO, PASO_PREOPERATIVO
    if step < 230:
        return ESTADO_OPERATIVO, PASO_OPERATIVO
    return ESTADO_FINALIZADO, PASO_FINALIZADO


def update_cocina(nodes, sim_state):
    step = sim_state["step"]
    numero_equipo = sim_state["numero_equipo"]
    estado_equipo, paso_actual = resolve_coccion_phase(step)
    cycle_id = sim_state["cycle_id"]
    lote = sim_state["lote"]

    if step == 0 and cycle_id not in sim_state["start_logged_for_cycle"]:
        log_cycle_event("COCINA", numero_equipo, lote, "INICIO DE CICLO")
        sim_state["start_logged_for_cycle"].add(cycle_id)

    if step == CICLO_COCCION_MAX_STEP and cycle_id not in sim_state["end_logged_for_cycle"]:
        log_cycle_event("COCINA", numero_equipo, lote, "FIN DE CICLO")
        sim_state["end_logged_for_cycle"].add(cycle_id)

    if sim_state["last_estado"] != estado_equipo:
        log_state_change("COCINA", numero_equipo, lote, estado_equipo, paso_actual, step)
        sim_state["last_estado"] = estado_equipo

    peso_producto = random.randint(850, 1050)

    if estado_equipo == ESTADO_PREOPERATIVO:
        carga_agua = True
        vapor_vivo_on = True
        vapor_serpentina_on = True
        bomba_centrifuga = True
        filtro_succion = False
        temp_agua = 24 + ((85 - 24) / 30) * step
        temp_producto = 0.0
        nivel_agua = random.uniform(1560, 1650)
    elif estado_equipo == ESTADO_OPERATIVO:
        carga_agua = False
        vapor_vivo_on = False
        vapor_serpentina_on = True
        bomba_centrifuga = True
        filtro_succion = True
        base_temp = 80 + ((90 - 80) / 140) * (step - 30)
        temp_agua = base_temp + random.uniform(-1, 1)

        t = step - 30
        l_value = 70
        k_value = 0.09
        x0 = 45
        temp_producto = 5 + l_value / (1 + math.exp(-k_value * (t - x0)))
        temp_producto = min(temp_producto, 75)

        if step < 40:
            nivel_agua = 1650 + ((2000 - 1650) / 10) * (step - 30)
        else:
            nivel_agua = random.uniform(1900, 2000)
    else:
        carga_agua = False
        vapor_vivo_on = False
        vapor_serpentina_on = False
        bomba_centrifuga = False
        filtro_succion = False
        temp_agua = 90 + random.uniform(-0.5, 0.5)

        t = step - 30
        l_value = 70
        k_value = 0.09
        x0 = 45
        temp_producto = 5 + l_value / (1 + math.exp(-k_value * (t - x0)))
        temp_producto = min(temp_producto, 75)

        descenso_final = max(0, step - 170)
        nivel_agua = max(1800, 1900 - (descenso_final * random.uniform(1, 3)))

    set_value(nodes["FILTRO_SUCCION_AGUA"], filtro_succion, ua.VariantType.Boolean)
    set_value(nodes["CARGA_AGUA"], carga_agua, ua.VariantType.Boolean)
    set_value(nodes["VAPOR_SERPENTINA_ACC"], vapor_serpentina_on, ua.VariantType.Boolean)
    set_value(nodes["VAPOR_VIVO_ACC"], vapor_vivo_on, ua.VariantType.Boolean)
    set_value(nodes["BOMBA_CENTRIFUGA"], bomba_centrifuga, ua.VariantType.Boolean)
    set_value(nodes["CICLO_TIPO_FIN"], True, ua.VariantType.Boolean)

    set_value(nodes["TEMP_AGUA"], round(temp_agua, 2), ua.VariantType.Float)
    set_value(nodes["TEMP_PRODUCTO"], round(temp_producto, 2), ua.VariantType.Float)
    set_value(nodes["NIVEL_AGUA"], round(nivel_agua, 2), ua.VariantType.Float)

    set_value(nodes["VAPOR_SERPENTINA"], bool_to_int16(vapor_serpentina_on), ua.VariantType.Int16)
    set_value(nodes["VAPOR_VIVO"], bool_to_int16(vapor_vivo_on), ua.VariantType.Int16)
    set_value(nodes["PASO_ACTUAL"], paso_actual, ua.VariantType.Int16)
    set_value(nodes["NUMERO_RECETA"], 1, ua.VariantType.Int16)
    set_value(nodes["CANTIDAD_TORRES"], 1, ua.VariantType.Int16)
    set_value(nodes["ESTADO_EQUIPO"], estado_equipo, ua.VariantType.Int16)
    set_value(nodes["PESO_PRODUCTO"], peso_producto, ua.VariantType.Int16)
    set_value(nodes["NUMERO_EQUIPO"], numero_equipo, ua.VariantType.Int16)
    set_value(nodes["TIEMPO_TRANS"], step, ua.VariantType.Int32)
    set_value(nodes["LOTE_CICLO"], lote, ua.VariantType.String)

    sim_state["step"] += 1
    if sim_state["step"] > CICLO_COCCION_MAX_STEP:
        sim_state["step"] = 0
        sim_state["cycle_id"] += 1
        sim_state["lote"] = build_lote("COC", numero_equipo, sim_state["cycle_id"])
        sim_state["start_logged_for_cycle"] = {cid for cid in sim_state["start_logged_for_cycle"] if cid >= sim_state["cycle_id"] - 1}
        sim_state["end_logged_for_cycle"] = {cid for cid in sim_state["end_logged_for_cycle"] if cid >= sim_state["cycle_id"] - 1}


def update_enfriador(nodes, sim_state):
    step = sim_state["step"]
    numero_equipo = sim_state["numero_equipo"]
    estado_equipo, paso_actual = resolve_enfriamiento_phase(step)
    cycle_id = sim_state["cycle_id"]
    lote = sim_state["lote"]

    if step == 0 and cycle_id not in sim_state["start_logged_for_cycle"]:
        log_cycle_event("ENFRIADOR", numero_equipo, lote, "INICIO DE CICLO")
        sim_state["start_logged_for_cycle"].add(cycle_id)

    if step == CICLO_ENFRIAMIENTO_RESET_STEP - 1 and cycle_id not in sim_state["end_logged_for_cycle"]:
        log_cycle_event("ENFRIADOR", numero_equipo, lote, "FIN DE CICLO")
        sim_state["end_logged_for_cycle"].add(cycle_id)

    if sim_state["last_estado"] != estado_equipo:
        log_state_change("ENFRIADOR", numero_equipo, lote, estado_equipo, paso_actual, step)
        sim_state["last_estado"] = estado_equipo

    peso_producto = random.randint(850, 1050)

    if estado_equipo == ESTADO_PREOPERATIVO:
        carga_agua = True
        amoniaco_on = True
        bomba_centrifuga = True
        filtro_succion = False
        vapor_limpieza_on = False
        temp_agua = max(24 - ((24 - 2) / 30) * step, 2)
        temp_producto = 0.0
        nivel_agua = random.uniform(1560, 1650)
    elif estado_equipo == ESTADO_OPERATIVO:
        carga_agua = False
        amoniaco_on = True
        bomba_centrifuga = True
        filtro_succion = True
        vapor_limpieza_on = False
        temp_agua = random.uniform(1, 5)

        l_value = 70
        k_value = 0.03
        temp_producto = 75 - (l_value * (1 - math.exp(-k_value * step)))
        temp_producto = max(temp_producto, 5)

        if step < 40:
            nivel_agua = 1650 + ((2000 - 1650) / 10) * (step - 30)
        else:
            nivel_agua = random.uniform(1900, 2000)
    else:
        carga_agua = False
        amoniaco_on = False
        bomba_centrifuga = False
        filtro_succion = False
        vapor_limpieza_on = False
        temp_agua = 2.0
        temp_producto = 5.0
        descenso_final = max(0, step - 230)
        nivel_agua = max(1800, 1900 - (descenso_final * random.uniform(1, 3)))

    set_value(nodes["FILTRO_SUCCION_AGUA"], filtro_succion, ua.VariantType.Boolean)
    set_value(nodes["CARGA_AGUA"], carga_agua, ua.VariantType.Boolean)
    set_value(nodes["AMONIACO_ACC"], amoniaco_on, ua.VariantType.Boolean)
    set_value(nodes["VAPOR_LIMPIEZA_ACC"], vapor_limpieza_on, ua.VariantType.Boolean)
    set_value(nodes["BOMBA_CENTRIFUGA"], bomba_centrifuga, ua.VariantType.Boolean)
    set_value(nodes["CICLO_TIPO_FIN"], True, ua.VariantType.Boolean)

    set_value(nodes["TEMP_AGUA"], round(temp_agua, 2), ua.VariantType.Float)
    set_value(nodes["TEMP_PRODUCTO"], round(temp_producto, 2), ua.VariantType.Float)
    set_value(nodes["NIVEL_AGUA"], round(nivel_agua, 2), ua.VariantType.Float)

    set_value(nodes["AMONIACO"], bool_to_int16(amoniaco_on), ua.VariantType.Int16)
    set_value(nodes["VAPOR_LIMPIEZA"], bool_to_int16(vapor_limpieza_on), ua.VariantType.Int16)
    set_value(nodes["PASO_ACTUAL"], paso_actual, ua.VariantType.Int16)
    set_value(nodes["NUMERO_RECETA"], 1, ua.VariantType.Int16)
    set_value(nodes["CANTIDAD_TORRES"], 1, ua.VariantType.Int16)
    set_value(nodes["ESTADO_EQUIPO"], estado_equipo, ua.VariantType.Int16)
    set_value(nodes["PESO_PRODUCTO"], peso_producto, ua.VariantType.Int16)
    set_value(nodes["NUMERO_EQUIPO"], numero_equipo, ua.VariantType.Int16)
    set_value(nodes["TIEMPO_TRANS"], step, ua.VariantType.Int32)
    set_value(nodes["LOTE_CICLO"], lote, ua.VariantType.String)

    sim_state["step"] += 1
    if sim_state["step"] >= CICLO_ENFRIAMIENTO_RESET_STEP:
        sim_state["step"] = 0
        sim_state["cycle_id"] += 1
        sim_state["lote"] = build_lote("ENF", numero_equipo, sim_state["cycle_id"])
        sim_state["start_logged_for_cycle"] = {cid for cid in sim_state["start_logged_for_cycle"] if cid >= sim_state["cycle_id"] - 1}
        sim_state["end_logged_for_cycle"] = {cid for cid in sim_state["end_logged_for_cycle"] if cid >= sim_state["cycle_id"] - 1}


def build_cocina_states():
    return [
        create_cocina_sim_state(numero_equipo=i + 1, start_step=offset)
        for i, offset in enumerate(COCCION_START_OFFSETS)
    ]



def build_enfriador_states():
    return [
        create_enfriador_sim_state(numero_equipo=i + 1, start_step=offset)
        for i, offset in enumerate(ENFRIAMIENTO_START_OFFSETS)
    ]


# =========================
# MAIN
# =========================
def main():
    ip = get_local_ip()

    server = Server()
    server.set_endpoint(f"opc.tcp://{ip}:4841")
    server.set_server_name("Servidor OPC UA PF - Python")
    server.set_security_policy([ua.SecurityPolicyType.NoSecurity])

    uri = "http://pfalimentos.local/opcua/server/"
    idx = server.register_namespace(uri)

    objects = server.get_objects_node()

    server_interfaces = objects.add_object(idx, "ServerInterfaces")

    pf_l1 = server_interfaces.add_object(idx, "PF-L1")

    recetario = pf_l1.add_object(idx, "RECETARIO")
    recetas_nodes = []
    for i in range(11):
        recetas_nodes.append(build_recipe_item(recetario, idx, f"[{i}]", i))

    alarmas_l1 = pf_l1.add_object(idx, "ALARMAS L1")
    alarmas_l1_nodes, alarmas_l1_state = build_alarmas_l1(alarmas_l1, idx)

    cocina_l1 = pf_l1.add_object(idx, "COCINA L1")
    cocina_l1_nodes = []
    for i in range(3):
        cocina_l1_nodes.append(build_cocina_unit(cocina_l1, idx, f"[{i}]", i + 1))

    enfriador_l1 = pf_l1.add_object(idx, "ENFRIADOR L1")
    enfriador_l1_nodes = []
    for i in range(4):
        enfriador_l1_nodes.append(build_enfriador_unit(enfriador_l1, idx, f"[{i}]", i + 1))

    pf_l2 = server_interfaces.add_object(idx, "PF-L2")
    alarmas_l2 = pf_l2.add_object(idx, "ALARMAS L2")
    alarmas_l2_nodes, alarmas_l2_state = build_alarmas_l2(alarmas_l2, idx)

    cocina_l2 = pf_l2.add_object(idx, "COCINA L2")
    cocina_l2_nodes = []
    for i in range(3):
        cocina_l2_nodes.append(build_cocina_unit(cocina_l2, idx, f"[{i}]", i + 1))

    enfriador_l2 = pf_l2.add_object(idx, "ENFRIADOR L2")
    enfriador_l2_nodes = []
    for i in range(4):
        enfriador_l2_nodes.append(build_enfriador_unit(enfriador_l2, idx, f"[{i}]", i + 1))

    cocina_states = build_cocina_states()
    enfriador_states = build_enfriador_states()

    server.start()
    print("Servidor OPC UA iniciado")
    print(f"Endpoint: opc.tcp://{ip}:4841")
    print("Presioná Ctrl + C para detenerlo")

    try:
        while True:
            update_alarmas_l1(alarmas_l1_nodes, alarmas_l1_state)
            update_alarmas_l2(alarmas_l2_nodes, alarmas_l2_state)


            for nodes, sim_state in zip(cocina_l1_nodes, cocina_states[:3]):
                update_cocina(nodes, sim_state)

            #for nodes, sim_state in zip(enfriador_l1_nodes, enfriador_states[:4]):
            #    update_enfriador(nodes, sim_state)

            #for nodes, sim_state in zip(cocina_l2_nodes, cocina_states[3:]):
            #    update_cocina(nodes, sim_state)

            for nodes, sim_state in zip(enfriador_l2_nodes, enfriador_states[4:]):
                update_enfriador(nodes, sim_state)

            time.sleep(1)

    except KeyboardInterrupt:
        print("Deteniendo servidor...")
    finally:
        server.stop()
        print("Servidor detenido")


if __name__ == "__main__":
    main()
