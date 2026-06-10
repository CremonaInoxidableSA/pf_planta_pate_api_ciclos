from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Double, Boolean
from sqlalchemy.orm import relationship
from config.db import Base

class SensoresIO(Base):
    __tablename__ = "sensoresio"
    
    id = Column(Integer, primary_key=True, index=True)
    idSensor = Column(Integer, index=True, nullable=False)
    valor = Column(Boolean, index=True, nullable=False)
    fechaInicio = Column(DateTime, index=True, nullable=False)
    fechaFin = Column(DateTime, index=True, nullable=False)
    idCiclo = Column(Integer, index=True, nullable=False)