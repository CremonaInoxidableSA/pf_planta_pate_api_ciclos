from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Double
from sqlalchemy.orm import relationship
from config.db import Base


class EstadoCiclo(Base):
    __tablename__ = "estadociclo"
    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String(50), index=True, nullable=False)
    fechaInicio = Column(DateTime, nullable=False)
    fechaFin = Column(DateTime, nullable=False)
    tiempoTranscurrido = Column(String(50), index=True, nullable=True)

    idCiclo = Column(Integer, ForeignKey("ciclo.id"), nullable=False)
    ciclo = relationship("Ciclo", back_populates="estadociclo") 
