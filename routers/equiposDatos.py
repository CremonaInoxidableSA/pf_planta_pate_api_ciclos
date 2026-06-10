from fastapi import APIRouter, HTTPException
from config.opc import OPCUAClient
import socket
import os

from dotenv import load_dotenv

opc_ip = os.getenv("OPC_SERVER_IP")
opc_port = os.getenv("OPC_SERVER_PORT")

# Obtener la IP local y definir la URL del servidor OPC
localIp = socket.gethostbyname(socket.gethostname())
URL = f"opc.tcp://{opc_ip}:{opc_port}"

# Inicializar y conectar el cliente OPC
opc_client = OPCUAClient(URL)
opc_client.connect()

RouterLive = APIRouter(prefix="/dev", tags=["PruebaHTTPDatosOPC"], responses={404: {"description": "Sin acceso al servidor OPC"}})
