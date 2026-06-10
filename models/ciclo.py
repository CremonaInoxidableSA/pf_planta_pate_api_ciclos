from sqlalchemy import Column, Integer, String, DateTime, Float,ForeignKey
from sqlalchemy.orm import relationship
from config.db import Base

class Ciclo(Base):
    __tablename__ = "ciclo"
    
    id = Column(Integer, primary_key=True, index=True)
    estadoMaquina = Column(String(50), index=True, nullable=True)
    cantidadTorres = Column(Integer, index=True, nullable=True)
    lote = Column(String(50), index=True, nullable=True)
    cantidadPausas = Column(Integer, index=True, nullable=True)
    tiempoTranscurrido = Column(String(50), index=True, nullable=True)
    fecha_inicio = Column(DateTime, index=True, nullable=True)
    fecha_fin = Column(DateTime, index=True, nullable=True)
    peso = Column(Float, index=True, nullable=True)

    idEquipo = Column(Integer, ForeignKey("equipo.id"), nullable=True)
    equipo = relationship("Equipo", back_populates="ciclo")

    idReceta = Column(Integer, ForeignKey("receta.id"),nullable=True)
    receta = relationship("Receta", back_populates="ciclo")

    estadociclo = relationship("EstadoCiclo", back_populates="ciclo")


