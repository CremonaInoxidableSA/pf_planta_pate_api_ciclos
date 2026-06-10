from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Double
from sqlalchemy.orm import relationship
from config.db import Base

class Sensores(Base):
    __tablename__ = "sensores"
    
    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String(50), index=True, nullable=False)
    tipo = Column(String(50), index=True, nullable=False)