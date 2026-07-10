from enum import Enum

from sqlalchemy import (
    Column,
    Integer,
    String,
    DateTime,
    Boolean,
    Enum as SQLEnum,
    text,
)

from config.db import Base


class RolUsuario(str, Enum):
    superadmin = "superadmin"
    admin = "admin"
    user = "user"


class Usuario(Base):
    __tablename__ = "Usuarios"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)

    email = Column(String(255), nullable=False, unique=True, index=True)
    usuario = Column(String(50), nullable=False, unique=True, index=True)

    nombre = Column(String(100), nullable=False)
    apellido = Column(String(100), nullable=False)

    rol = Column(
        SQLEnum(
            RolUsuario,
            values_callable=lambda enum: [item.value for item in enum],
            native_enum=True,
        ),
        nullable=False,
        default=RolUsuario.user,
        server_default="user",
    )

    password_hash = Column(String(255), nullable=False)

    habilitado = Column(
        Boolean,
        nullable=False,
        default=True,
        server_default=text("1"),
    )

    reporte = Column(
        Boolean,
        nullable=False,
        default=False,
        server_default=text("0"),
    )

    ultimo_envio = Column(DateTime, nullable=True)