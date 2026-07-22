from typing import Union
from fastapi import FastAPI, WebSocket, WebSocketDisconnect, Query, Depends
from fastapi.responses import StreamingResponse
from datetime import date, datetime, time
from contextlib import asynccontextmanager
from sqlalchemy import text
from starlette.middleware.cors import CORSMiddleware

from config.opc import OPCUAClient
from config.ws import ws_manager
from config import db

from services.opcClienteService import ObtenerNodosOpcUA

from models.ciclo import Ciclo
from models.sensoresIO import SensoresIO
from models.sensoresAA import SensoresAA
from models.sensores import Sensores
from models.equipo import Equipo
from models.receta import Receta
from models.estadoCiclo import EstadoCiclo
from models.alarmas import Alarmas
from models.alarmasL2 import AlarmasL2
from models.historicoAlarma import HistoricoAlarma
from models.usuarios import Usuario

from routers import equiposDatos, historicoGraficos, historicoProductividad, historicoAlarmas

import logging
import asyncio

from dotenv import load_dotenv
import os

load_dotenv()

opc_ip   = os.getenv("OPC_SERVER_IP")
opc_port = os.getenv("OPC_SERVER_PORT")

ruta_principal = os.path.dirname(os.path.abspath(__file__))
logger = logging.getLogger("uvicorn")

URL        = f"opc.tcp://{opc_ip}:{opc_port}"
opc_client = OPCUAClient(URL)

db.Base.metadata.create_all(bind=db.engine)
dGeneral = ObtenerNodosOpcUA(opc_client)

ruta_sql_sensores = os.path.join(ruta_principal, 'data/bdd', 'insert_sensores.sql')
ruta_sql_equipos  = os.path.join(ruta_principal, 'data/bdd', 'insert_equipos.sql')
ruta_sql_alarmas_l1  = os.path.join(ruta_principal, 'data/bdd', 'insert_alarmas_l1.sql')
ruta_sql_alarmas_l2  = os.path.join(ruta_principal, 'data/bdd', 'insert_alarmas_l2.sql')

_reconexion_lock = asyncio.Lock()

def cargar_archivo_sql(file_path: str):
    try:
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding="utf-8") as f:
                sql_script = f.read()
            with db.engine.connect() as conn:
                conn.execute(text(sql_script))
                conn.commit()
                logger.info(f"SQL ejecutado: {file_path}")
        else:
            logger.error(f"Archivo SQL no encontrado: {file_path}")
    except Exception as e:
        logger.error(f"Error al cargar SQL: {e}")


async def central_opc_render():
    
    while True:
        try:
            datos = await dGeneral.datosGenerales()
            await ws_manager.send_message("datos-generales", datos)
        except Exception as e:
            logger.error(f"Error en el loop WebSocket: {e}")
        await asyncio.sleep(1.0)

async def monitor_opc(period_ms: int = 500, check_interval: int = 10):
    logger.info(f"Monitor OPC iniciado (ping cada {check_interval}s).")
    while True:
        await asyncio.sleep(check_interval)

        alive = await opc_client.ping()
        if alive:
            continue

        logger.warning("⚠️ Monitor OPC: ping fallido. Iniciando recuperación...")

        if _reconexion_lock.locked():
            logger.info("🔒 Reconexión ya en curso, saltando este ciclo.")
            continue

        async with _reconexion_lock:
            try:
                await opc_client.reconnect()

                if not opc_client.connected:
                    logger.error("❌ Reconexión fallida. Se reintentará en el próximo ciclo.")
                    continue
                await dGeneral.iniciar_suscripcion(period_ms=period_ms)

                logger.info("✅ Recuperación OPC completada. Sistema reanudado.")

            except Exception as e:
                logger.error(f"❌ Error durante la recuperación OPC: {e}")

@asynccontextmanager
async def lifespan(app: FastAPI):
    session = db.SessionLocal()
    try:
        await opc_client.connect()
        logger.info("Conectado al servidor OPC UA.")

        await dGeneral.iniciar_suscripcion(period_ms=500)

        asyncio.create_task(central_opc_render())
        asyncio.create_task(monitor_opc(period_ms=500, check_interval=10))
        if session.query(Sensores).count() == 0:
            logger.info("Cargando registros BDD [Sensores]")
            cargar_archivo_sql(ruta_sql_sensores)
        if session.query(Equipo).count() == 0:
            logger.info("Cargando registros BDD [Equipos]")
            cargar_archivo_sql(ruta_sql_equipos)
        if session.query(Alarmas).count() == 0:
            logger.info("Cargando registros BDD [Alarmas_l1]")
            cargar_archivo_sql(ruta_sql_alarmas_l1)
        if session.query(AlarmasL2).count() == 0:
            logger.info("Cargando registros BDD [Alarmas_l2]")
            cargar_archivo_sql(ruta_sql_alarmas_l2)
        yield

    finally:
        await opc_client.disconnect()
        session.close()


app = FastAPI(lifespan=lifespan)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(historicoGraficos.RoutersGraficosH)
app.include_router(historicoProductividad.RouterProductividad)
app.include_router(historicoAlarmas.RouterAlarmas)

@app.websocket("/ws/{id}")
async def resumen_desmoldeo(websocket: WebSocket, id: str):
    await websocket.accept()
    await ws_manager.connect(id, websocket)
    try:
        while True:
            await websocket.receive_json()
            await ws_manager.send_message(id, "data")
            await asyncio.sleep(0.2)
    except WebSocketDisconnect:
        await ws_manager.disconnect(id, websocket)


@app.get("/")
def read_root():
    try:
        with db.engine.connect() as connection:
            connection.execute(text("SELECT 1"))
        estado_bdd = "Conectado"
    except Exception as e:
        estado_bdd = "Desconectado"
    
    try:
        if opc_client.connected and opc_client.client:
            root_node = opc_client.client.get_root_node()
            root_node.get_browse_name()
            estado_opc = "Conectado"
        else:
            estado_opc = "Desconectado"
    except Exception as e:
        estado_opc = "Desconectado"
    
    return {
        "Hola Mundo-": " Levanto el server!", 
        "Estado BDD": estado_bdd, 
        "Estado OPC": estado_opc,
        "Fecha actual": datetime.now().strftime("%d-%m-%Y %H-%M")
    }