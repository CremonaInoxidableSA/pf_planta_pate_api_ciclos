from sqlalchemy import Column, Integer, String, DateTime,ForeignKey
from sqlalchemy.orm import relationship
from config.db import Base
from datetime import datetime

class Alarmas(Base):
    __tablename__ = "alarmas"

    id = Column(Integer, primary_key=True, index=True, autoincrement=False)
    nombre = Column(String(100), index=True)
    tipoAlarma = Column(String(30), index=True)
    descripcion = Column(String(255), index=True)
    fechaRegistro = Column(DateTime, default=datetime)