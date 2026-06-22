from fastapi import APIRouter, HTTPException, Depends, Query
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from datetime import datetime, date, time
from config import db

from services.historicoServices import obtener_historico_alarmas, generar_reporte_alarmas_descarga

import logging

logger = logging.getLogger("uvicorn")

RouterAlarmas = APIRouter(prefix="/historico-alarmas", tags=["Alarmas Historico"]) 
@RouterAlarmas.get("/alarmas")
def listar_alarmas_bdd(
    fecha_inicio: date = Query(..., description="Fecha de inicio (YYYY-MM-DD)"),
    fecha_fin: date = Query(..., description="Fecha de fin (YYYY-MM-DD)"),
):
    logger.info(f"Endpoint /alarmas llamado con fecha_inicio={fecha_inicio} y fecha_fin={fecha_fin}")
    if fecha_inicio is not None and fecha_fin is not None:
        fecha_inicio_dt = datetime.combine(fecha_inicio, time.min)
        fecha_fin_dt = datetime.combine(fecha_fin, datetime.max.time())

        return obtener_historico_alarmas(fecha_inicio_dt, fecha_fin_dt, db.SessionLocal())
    
    return obtener_historico_alarmas(fecha_inicio, fecha_fin, db.SessionLocal())

@RouterAlarmas.get("/alarmas/defecto")
def listar_todas_alarmas_bdd():
    logger.info(f"Endpoint /alarmas/defecto llamado")
    return obtener_historico_alarmas(fecha_inicio=None, fecha_fin=None, session=db.SessionLocal())


@RouterAlarmas.get("/alarmas/descargar")
def descargar_alarmas_excel(
    fecha_inicio: date = Query(..., description="Fecha de inicio (YYYY-MM-DD)"),
    fecha_fin: date = Query(..., description="Fecha de fin (YYYY-MM-DD)"),

):
    logger.info(f"Endpoint /alarmas/descargar llamado con fecha_inicio={fecha_inicio} y fecha_fin={fecha_fin}")
    if fecha_inicio is not None and fecha_fin is not None:
        fecha_inicio_dt = datetime.combine(fecha_inicio, time.min)
        fecha_fin_dt = datetime.combine(fecha_fin, datetime.max.time())
        xlms_stream = generar_reporte_alarmas_descarga(db.SessionLocal(), fecha_inicio_dt, fecha_fin_dt)
        fecha_actual = datetime.now().strftime("%Y-%m-%d_%H-%M")
        nombreArchivo = f"informe_alarmas_{fecha_actual}.xlsx"

        return StreamingResponse(
            xlms_stream, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            headers={"Content-Disposition": f"attachment; filename={nombreArchivo}"}
        )
    else:
        return {"error": "Fechas no válidas. Asegúrese de proporcionar fecha_inicio y fecha_fin."}
    