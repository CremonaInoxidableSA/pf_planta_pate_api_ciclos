from sqlalchemy import (
    Column,
    Integer,
    String,
    Boolean,
    ForeignKey,
    text,
)
from sqlalchemy.orm import relationship

from config.db import Base


class Receta(Base):
    __tablename__ = "receta"

    id = Column(Integer, primary_key=True, index=True)

    nombre = Column(
        String(50),
        index=True,
        nullable=False,
        default="SIN RECETA",
        server_default="SIN RECETA",
    )

    nroPaso = Column(
        Integer,
        index=True,
        nullable=False,
        default=0,
        server_default=text("0"),
    )

    tipoFin = Column(
        String(50),
        index=True,
        nullable=False,
        default="",
        server_default="",
    )

    onoffalarma = Column(
        Boolean,
        index=True,
        nullable=False,
        default=False,
        server_default=text("0"),
    )

    pesoXTorre = Column(
        Integer,
        index=True,
        nullable=False,
        default=0,
        server_default=text("0"),
    )

    tempAgua = Column(
        String(50),
        index=True,
        nullable=False,
        default="",
        server_default="",
    )

    tempCorteEnfriado = Column(
        Integer,
        index=True,
        nullable=False,
        default=0,
        server_default=text("0"),
    )

    tempProducto = Column(
        String(50),
        index=True,
        nullable=False,
        default="",
        server_default="",
    )

    tiempoCorte = Column(
        String(50),
        index=True,
        nullable=False,
        default="",
        server_default="",
    )

    tiempoCorteEnfriado = Column(
        Integer,
        index=True,
        nullable=False,
        default=0,
        server_default=text("0"),
    )

    tiempoParaAlarma = Column(
        Integer,
        index=True,
        nullable=False,
        default=0,
        server_default=text("0"),
    )

    tipoCorte = Column(
        String(50),
        index=True,
        nullable=False,
        default="",
        server_default="",
    )

    ciclo = relationship("Ciclo", back_populates="receta")