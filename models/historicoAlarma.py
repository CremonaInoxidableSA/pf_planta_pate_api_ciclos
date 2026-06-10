from sqlalchemy import Column, Integer, DateTime, String,ForeignKey, Boolean
from sqlalchemy.orm import relationship
from config.db import Base

class HistoricoAlarma(Base):
    __tablename__ = "historico_alarma"
    
    id = Column(Integer, primary_key=True, index=True)
    idAlarma = Column(Integer, index=True, nullable=True)
    fechaInicio = Column(DateTime, index=True, nullable=False)
    tipoLinea     = Column(String(100), nullable=True)
    fechaFin = Column(DateTime, index=True, nullable=True) 
    valor = Column(Boolean, index=True, nullable=False)
