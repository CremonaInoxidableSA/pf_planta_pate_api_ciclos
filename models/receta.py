from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Double, Boolean
from sqlalchemy.orm import relationship
from config.db import Base

class Receta(Base):
    __tablename__ = "receta"
    
    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String(50), index=True, nullable=False)
    nroPaso = Column(Integer, index=True, nullable=False)
    tipoFin = Column(String(50), index=True, nullable=False)
    onoffalarma = Column(Boolean, index=True, nullable=False)
    pesoXTorre = Column(Integer, index=True, nullable=False)
    tempAgua = Column(String(50), index=True, nullable=False)
    tempCorteEnfriado = Column(Integer, index=True, nullable=False)
    tempProducto = Column(String(50), index=True, nullable=False)
    tiempoCorte = Column(String(50), index=True, nullable=False)
    tiempoCorteEnfriado = Column(Integer, index=True, nullable=False)
    tiempoParaAlarma = Column(Integer, index=True, nullable=False)
    tipoCorte = Column(String(50), index=True, nullable=False)

    ciclo = relationship("Ciclo", back_populates="receta")