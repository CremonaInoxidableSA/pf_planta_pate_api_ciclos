from fastapi import APIRouter, HTTPException, Depends, Query
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from datetime import datetime, date
from config import db
from services.historicoServices import productividad_equipo, generar_reporte_productividad

import logging

logger = logging.getLogger("uvicorn")

RouterProductividad = APIRouter(prefix="/historico-productividad", tags=["Productividad Historico"]) 
#http://localhost/historico/<id_equipo>/<fecha_inicio>/<fecha_fin>
@RouterProductividad.get("/{id_equipo}")
def obtener_lista_productividad(
    id_equipo: int, 
    fecha_inicio: date = Query(..., description="Fecha de inicio (YYYY-MM-DD)"),
    fecha_fin: date = Query(..., description="Fecha de fin (YYYY-MM-DD)"),
    db : Session = Depends(db.get_db)
):
    datos = productividad_equipo(db, fecha_inicio, fecha_fin, id_equipo)
    return datos


#http://localhost/historico/<linea>/<fecha_inicio>/<fecha_fin>

@RouterProductividad.get("/descargar/{id_equipo}")
def descargar_archivo_xlms(
    id_equipo: int, 
    fecha_inicio: date = Query(..., description="Fecha de inicio (YYYY-MM-DD)"),
    fecha_fin: date = Query(..., description="Fecha de fin (YYYY-MM-DD)"),
    db : Session = Depends(db.get_db)
):
    xlms = generar_reporte_productividad(id_equipo, fecha_inicio, fecha_fin, db)
    fecha_actual = datetime.now().strftime("%Y-%m-%d_%H-%M")
    nombreArchivo = f"Informe_productividad_{fecha_actual}.xlsx"

    return StreamingResponse(
        xlms, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        headers={"Content-Disposition": f"attachment; filename={nombreArchivo}"}
    )