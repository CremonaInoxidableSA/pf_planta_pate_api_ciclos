from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Double
from sqlalchemy.orm import relationship
from config.db import Base

class Receta(Base):
    __tablename__ = "receta"
    
    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String(50), index=True, nullable=False)
    nroPaso = Column(Integer, index=True, nullable=False)
    tipoFin = Column(String(50), index=True, nullable=False)

    ciclo = relationship("Ciclo", back_populates="receta")