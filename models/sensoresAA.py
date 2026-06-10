from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Double, Float
from sqlalchemy.orm import relationship
from config.db import Base

class SensoresAA(Base):
    __tablename__ = "sensoresaa"
    
    id = Column(Integer, primary_key=True, index=True)
    idSensor = Column(Integer, index=True, nullable=False)
    valor = Column(Float, index=True, nullable=True)
    idCiclo = Column(Integer, index=True, nullable=False)
    fechaRegistro = Column(DateTime, index=True, nullable=False)