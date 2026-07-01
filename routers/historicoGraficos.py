from fastapi import APIRouter, HTTPException, Depends, Query
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from datetime import datetime, date, time, timedelta

from models.ciclo import Ciclo
from models.equipo import Equipo
from models.sensores import Sensores

from services.historicoServices import obtener_datos_graficos, generar_informe_ciclo

from config import db

#http://localhost/historico/<equipo>/<fecha_inicio>/<fecha_fin> 

RoutersGraficosH = APIRouter(prefix="/historico-graficos", tags=["Graficos Historico"]) 

from fastapi import APIRouter, HTTPException, Depends, Query
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from datetime import datetime, date, time, timedelta

from models.ciclo import Ciclo
from models.equipo import Equipo
from models.sensores import Sensores

from services.historicoServices import obtener_datos_graficos, generar_informe_ciclo

from config import db

#http://localhost/historico/<equipo>/<fecha_inicio>/<fecha_fin> 

RoutersGraficosH = APIRouter(prefix="/historico-graficos", tags=["Graficos Historico"]) 

@RoutersGraficosH.get("/ultimo-ciclo")
def obtener_ultimo_ciclo():
    try:
        ultimo_ciclo = (
            db.SessionLocal().query(Ciclo)
            .filter(Ciclo.fecha_fin != None)
            .order_by(Ciclo.fecha_fin.desc())
            .first()
        )

        if not ultimo_ciclo:
            return {"message": "No se encontraron ciclos en la base de datos."}

        ciclo_data = {
            "id_ciclo": ultimo_ciclo.id,
            "id_equipo": ultimo_ciclo.idEquipo,
            "lote": ultimo_ciclo.lote,
            "fecha_inicio": ultimo_ciclo.fecha_inicio,
            "fecha_fin": ultimo_ciclo.fecha_fin,
            "tiempo_transcurrido": ultimo_ciclo.tiempoTranscurrido
        }
        return ciclo_data

    except Exception as e:
        print("Error:", e)
        raise HTTPException(status_code=500, detail="Error al obtener el último ciclo.")

@RoutersGraficosH.get("/{equipo}")
def obtener_lista_ciclos(
    equipo: str,
    fecha_inicio: date = Query(..., description="Fecha de inicio (YYYY-MM-DD)"),
    fecha_fin: date = Query(..., description="Fecha de fin (YYYY-MM-DD)"),
    db: Session = Depends(db.get_db)
):
    fecha_inicio_dt = datetime.combine(fecha_inicio, time.min)
    fecha_fin_dt = datetime.combine(fecha_fin, datetime.max.time())
    lista_ciclos = []

    try:
        ciclos = (
            db.query(Ciclo)
            .filter(
                Ciclo.idEquipo == int(equipo),
                Ciclo.fecha_fin >= fecha_inicio_dt,
                Ciclo.fecha_fin < fecha_fin_dt
            )
            .all()
        )

        if not ciclos:
            return None

        for elem in ciclos:
            if elem.estadoMaquina == "FINALIZADO":
                ciclo = {
                    "id_ciclo": elem.id,
                    "lote": elem.lote,
                    "fecha_inicio": elem.fecha_inicio,
                    "fecha_fin": elem.fecha_fin,
                    "tiempo_transcurrido": elem.tiempoTranscurrido,
                    "estado_maquina": elem.estadoMaquina
                }
                lista_ciclos.append(ciclo)

        return lista_ciclos if lista_ciclos else None

    except Exception as e:
        print("Error:", e)
        return None


@RoutersGraficosH.get("/{equipo}/{id_ciclo}")
def obtener_datos_sensores(
    equipo: str, 
    id_ciclo: int, 
    db : Session = Depends(db.get_db)
):
    return obtener_datos_graficos(db, id_ciclo)

@RoutersGraficosH.get("/{equipo}/descargar/{id_ciclo}")
def descargar_archivo_xlms(
    id_ciclo:int, 
    equipo: str,
    db: Session = Depends(db.get_db)
):
    fecha_actual = datetime.now().strftime("%Y-%m-%d_%H-%M")
    nombreArchivo = f"informe_ciclo_{id_ciclo}_{fecha_actual}.xlsx"
    try:
        xlms_stream = generar_informe_ciclo(db, id_ciclo, equipo)
        
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))

    return StreamingResponse(
        xlms_stream, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        headers={"Content-Disposition": f"attachment; filename={nombreArchivo}"}
    )