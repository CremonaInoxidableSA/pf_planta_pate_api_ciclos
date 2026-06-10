from sqlalchemy import create_engine, MetaData
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from dotenv import load_dotenv
import os
import pymysql
from sqlalchemy.exc import OperationalError

load_dotenv()

db_user = os.getenv("MYSQL_USER")
db_password = os.getenv("MYSQL_PASSWORD")
db_host = os.getenv("MYSQL_HOST", "localhost")  # Por defecto usa localhost si no está configurado
db_port = os.getenv("MYSQL_PORT", "3306")
db_name = os.getenv("MYSQL_DATABASE")
db_url = os.getenv("DATABASE_URL")

DATABASE_URL = f"mysql+pymysql://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}"

print(f"Usuario: {db_user}, Password: {db_password}, Host: {db_host}, Puerto: {db_port}, DB: {db_name}")

def try_connect(db_url):
    print(f"DATO ULR: {db_url}")
    try:
        # Intenta crear una conexión al motor de SQLAlchemy
        engine = create_engine(db_url)
        with engine.connect() as connection:
            print(f"Conectado a la base de datos en {db_url}")
            return engine
    except OperationalError as e:
        print(f"Error al conectar a la base de datos: {e}")
        return None

engine = try_connect(db_url)
if not engine:
    print("Intentando conexión con localhost...")
    db_host = "localhost"
    DATABASE_URL = f"mysql+pymysql://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}"
    engine = try_connect(DATABASE_URL)

if not engine:
    raise Exception("No se pudo conectar a la base de datos ni con el valor de la variable de entorno ni con localhost.")

Base = declarative_base()

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
