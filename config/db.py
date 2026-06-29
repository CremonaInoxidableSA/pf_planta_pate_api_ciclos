from sqlalchemy import create_engine, MetaData, text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from dotenv import load_dotenv
import os
import pymysql
import time
import threading
from sqlalchemy.exc import OperationalError

load_dotenv()

db_user = os.getenv("MYSQL_USER")
db_password = os.getenv("MYSQL_PASSWORD")
db_host = os.getenv("MYSQL_HOST", "localhost")
db_port = os.getenv("MYSQL_PORT", "3306")
db_name = os.getenv("MYSQL_DATABASE")
db_url = os.getenv("DATABASE_URL")

DATABASE_URL = f"mysql+pymysql://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}"

print(f"Usuario: {db_user}, Password: {db_password}, Host: {db_host}, Puerto: {db_port}, DB: {db_name}")

engine = None
engine_lock = threading.Lock()

def try_connect(db_url):
    print(f"DATO ULR: {db_url}")
    attempt = 1
    
    while True:
        try:
            test_engine = create_engine(db_url)
            with test_engine.connect() as connection:
                print(f"Conectado a la base de datos en {db_url}")
                return test_engine
        except OperationalError as e:
            print(f"Intento {attempt} - Error al conectar a la base de datos: {e}")
            print("Reintentando conexión en 3 segundos...")
            time.sleep(3)
            attempt += 1

def check_connection():
    """Verifica si la conexión está activa"""
    global engine
    try:
        if engine:
            with engine.connect() as connection:
                connection.execute(text("SELECT 1"))
                return True
    except Exception as e:
        print(f"Error en verificación de conexión: {e}")
        return False
    return False

def reconnect():
    """Función para reconectar cuando se detecta una falla"""
    global engine
    print("Detectada pérdida de conexión. Iniciando reconexión...")
    
    with engine_lock:
        engine = try_connect(db_url)
        
        if not engine:
            print("Intentando conexión con localhost...")
            local_db_host = "localhost"
            local_DATABASE_URL = f"mysql+pymysql://{db_user}:{db_password}@{local_db_host}:{db_port}/{db_name}"
            engine = try_connect(local_DATABASE_URL)
        
        global SessionLocal
        SessionLocal.configure(bind=engine)

def connection_monitor():
    """Monitor que verifica la conexión cada 10 segundos"""
    while True:
        time.sleep(10)
        if not check_connection():
            reconnect()

# Conexión inicial
engine = try_connect(db_url)

if not engine:
    print("Intentando conexión con localhost...")
    db_host = "localhost"
    DATABASE_URL = f"mysql+pymysql://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}"
    engine = try_connect(DATABASE_URL)

if not engine:
    raise Exception("No se pudo conectar a la base de datos ni con el valor de la variable de entorno ni con localhost.")

monitor_thread = threading.Thread(target=connection_monitor, daemon=True)
monitor_thread.start()
print("Monitor de conexión iniciado...")

Base = declarative_base()

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

def get_session_local():
    """Función para obtener SessionLocal actualizado"""
    global engine
    with engine_lock:
        return sessionmaker(autocommit=False, autoflush=False, bind=engine)

def get_db():
    SessionLocal = get_session_local()
    db = SessionLocal()
    try:
        yield db
    except Exception as e:
        db.rollback()
        raise
    finally:
        db.close()
        