from fastapi import WebSocket
from typing import Dict, Any, Set
import logging
import json
from datetime import datetime

logger = logging.getLogger("uvicorn")
TMessagePayload = Any
TActiveConnections = Dict[str, Set[WebSocket]]

def datetime_to_string(obj):
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError(f"Object of type {type(obj)} is not JSON serializable")

class WSManager:
    def __init__(self):
        self.active_connections: TActiveConnections = {}

    async def connect(self, poll_id: str, ws: WebSocket):
        self.active_connections.setdefault(poll_id, set()).add(ws)

    async def disconnect(self, poll_id: str, ws: WebSocket):
        self.active_connections[poll_id].remove(ws)

    async def send_message(self, poll_id: str, message: list):
        if poll_id in self.active_connections:
            websockets = self.active_connections[poll_id]
            for websocket in websockets:
                try:
                    # Convertir el mensaje a JSON usando el convertidor personalizado
                    json_str = json.dumps(message, default=datetime_to_string)
                    await websocket.send_text(json_str)
                except Exception as e:
                    logger.error(f"Error al enviar mensaje a WebSocket: {e}")
        

ws_manager = WSManager()