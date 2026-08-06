import csv
import io
import json
import logging
import os

from fastapi import APIRouter, HTTPException, Query
from fastapi.responses import StreamingResponse

from services.opcClienteService import CAMPOS_HISTORIAL

logger = logging.getLogger("uvicorn")

RouterHistoricoCrudo = APIRouter(
    prefix="/historico/crudo",
    tags=["HistorialCrudo"],
    responses={404: {"description": "No existe historial para ese ciclo"}},
)

HISTORIAL_DIR = os.path.abspath(
    os.getenv("OPC_HISTORIAL_DIR", "data/opc_historial")
)

EXTENSIONES = (".done", ".jsonl", ".processing")


# ---------------------------------------------------------------------------
# Catalogo de equipos
#
# Se construye con los mismos offsets que _navegar_arbol_sync. Ojo: el numero
# que aparece en el nombre del directorio es el numero LOCAL dentro del grupo,
# no el id global. La cocina con id_equipo 4 vive en l2_cocina_1, no en
# l2_cocina_4. Por eso el catalogo no puede derivarse de EQUIPOS_MAP.
# ---------------------------------------------------------------------------

_GRUPOS = (
    # (linea_slug, tipo_slug, cantidad, offset_id)
    ("l1", "cocina",    3,  0),
    ("l1", "enfriador", 4,  6),
    ("l2", "cocina",    3,  3),
    ("l2", "enfriador", 4, 10),
)


def _construir_catalogo():
    por_slug = {}
    por_id = {}
    for linea_slug, tipo_slug, cantidad, offset in _GRUPOS:
        for numero_local in range(1, cantidad + 1):
            slug = f"{linea_slug}_{tipo_slug}_{numero_local}"
            id_equipo = numero_local + offset
            info = {
                "slug": slug,
                "id_equipo": id_equipo,
                "linea": linea_slug.upper(),
                "tipo": tipo_slug.upper(),
                "numero_local": numero_local,
            }
            por_slug[slug] = info
            por_id[id_equipo] = info
    return por_slug, por_id


EQUIPOS_POR_SLUG, EQUIPOS_POR_ID = _construir_catalogo()


def _resolver_equipo(equipo: str) -> dict:
    """
    Acepta tres formas de identificar un equipo y devuelve siempre el mismo
    registro del catalogo:

        "3"                -> id_equipo global
        "l1_cocina_3"      -> slug del directorio
        "PF-L1-COCINA-3"   -> nombre visible

    El valor recibido NUNCA se usa para componer una ruta: se resuelve contra
    el catalogo cerrado. Asi un parametro como "../../etc" no puede escapar del
    directorio de historiales.
    """
    if equipo is None:
        raise HTTPException(status_code=400, detail="Falta el equipo")

    texto = str(equipo).strip()
    if not texto:
        raise HTTPException(status_code=400, detail="Falta el equipo")

    # Forma 1: id global
    if texto.isdigit():
        info = EQUIPOS_POR_ID.get(int(texto))
        if info:
            return info

    # Forma 2: slug del directorio
    normalizado = texto.lower().replace("-", "_").replace(" ", "_")
    info = EQUIPOS_POR_SLUG.get(normalizado)
    if info:
        return info

    # Forma 3: nombre visible tipo PF-L1-COCINA-3
    partes = texto.upper().replace("_", "-").split("-")
    if len(partes) == 4 and partes[0] == "PF":
        linea_slug = partes[1].lower()
        tipo_slug = partes[2].lower()
        try:
            numero_global = int(partes[3])
        except ValueError:
            numero_global = None
        if numero_global is not None:
            candidato = EQUIPOS_POR_ID.get(numero_global)
            if (
                candidato
                and candidato["linea"].lower() == linea_slug
                and candidato["tipo"].lower() == tipo_slug
            ):
                return candidato

    raise HTTPException(
        status_code=404,
        detail=(
            f"Equipo desconocido: {equipo!r}. "
            f"Valores validos: {sorted(EQUIPOS_POR_SLUG)} "
            f"o los id {sorted(EQUIPOS_POR_ID)}."
        ),
    )


def _buscar_archivo_ciclo(info: dict, id_ciclo: int):
    """
    Busca el historial de un ciclo dentro del directorio de su equipo.

    Devuelve (ruta, extension). Prioriza .done; si el ciclo todavia no cerro
    acepta .jsonl o .processing. Si el archivado genero variantes con marca de
    tiempo (ciclo_4120_20260806_101500.done) se toma la mas reciente.
    """
    directorio = os.path.join(HISTORIAL_DIR, info["slug"])
    if not os.path.isdir(directorio):
        raise HTTPException(
            status_code=404,
            detail=(
                f"No existe el directorio de historiales del equipo "
                f"{info['slug']}."
            ),
        )

    prefijo = f"ciclo_{int(id_ciclo)}"

    for extension in EXTENSIONES:
        exacto = os.path.join(directorio, prefijo + extension)
        candidatos = [exacto] if os.path.isfile(exacto) else []

        # Variantes con marca de tiempo, solo para .done
        if extension == ".done":
            candidatos += [
                os.path.join(directorio, nombre)
                for nombre in os.listdir(directorio)
                if nombre.startswith(prefijo + "_")
                and nombre.endswith(extension)
            ]

        if not candidatos:
            continue

        candidatos.sort(key=os.path.getmtime, reverse=True)
        elegido = candidatos[0]

        if len(candidatos) > 1:
            logger.warning(
                "[CRUDO] El ciclo %s del equipo %s tiene %s archivos. "
                "Se entrega el mas reciente: %s",
                id_ciclo,
                info["slug"],
                len(candidatos),
                elegido,
            )

        # Defensa en profundidad: el archivo resuelto debe seguir estando
        # debajo del directorio de historiales.
        real = os.path.realpath(elegido)
        if not real.startswith(os.path.realpath(HISTORIAL_DIR) + os.sep):
            raise HTTPException(status_code=400, detail="Ruta invalida")

        return real, extension

    raise HTTPException(
        status_code=404,
        detail=(
            f"No se encontro historial del ciclo {id_ciclo} para el equipo "
            f"{info['slug']}."
        ),
    )


def _filas_csv(ruta: str, separador: str, columnas: tuple):
    """
    Generador que convierte el JSONL a CSV linea por linea.

    No carga el archivo completo en memoria: un ciclo de tres horas tiene unas
    diez mil lineas y se transmiten a medida que se leen.

    Las lineas que no se pueden interpretar se omiten y se registran en el log
    en lugar de abortar la descarga: este endpoint es una herramienta de
    auditoria y conviene que entregue todo lo que sea legible.
    """
    buffer = io.StringIO()
    escritor = csv.writer(
        buffer,
        delimiter=separador,
        quoting=csv.QUOTE_MINIMAL,
        lineterminator="\r\n",
    )

    def _volcar():
        datos = buffer.getvalue()
        buffer.seek(0)
        buffer.truncate(0)
        return datos

    # BOM para que Excel reconozca UTF-8 y no rompa los acentos.
    yield "\ufeff"

    escritor.writerow(columnas)
    yield _volcar()

    descartadas = 0
    leidas = 0

    with open(ruta, "r", encoding="utf-8", newline="") as f:
        for numero_linea, linea in enumerate(f, start=1):
            texto = linea.strip()
            if not texto:
                continue
            try:
                registro = json.loads(texto)
            except json.JSONDecodeError:
                # Tipicamente la ultima linea de un ciclo interrumpido.
                descartadas += 1
                logger.warning(
                    "[CRUDO] Linea %s ilegible en %s. Se omite.",
                    numero_linea,
                    ruta,
                )
                continue

            if not isinstance(registro, dict):
                descartadas += 1
                continue

            escritor.writerow([registro.get(campo) for campo in columnas])
            leidas += 1
            yield _volcar()

    logger.info(
        "[CRUDO] %s -> CSV: %s filas entregadas, %s descartadas.",
        os.path.basename(ruta),
        leidas,
        descartadas,
    )


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@RouterHistoricoCrudo.get("/equipos")
def listar_equipos():
    """Catalogo de equipos con la cantidad de historiales disponibles."""
    salida = []
    for slug in sorted(EQUIPOS_POR_SLUG):
        info = dict(EQUIPOS_POR_SLUG[slug])
        directorio = os.path.join(HISTORIAL_DIR, slug)
        info["historiales"] = (
            len([
                n for n in os.listdir(directorio)
                if n.endswith(EXTENSIONES)
            ])
            if os.path.isdir(directorio)
            else 0
        )
        salida.append(info)
    return {"directorio": HISTORIAL_DIR, "equipos": salida}


@RouterHistoricoCrudo.get("/{equipo}/ciclos")
def listar_ciclos(equipo: str):
    """Ids de ciclo con historial disponible para un equipo."""
    info = _resolver_equipo(equipo)
    directorio = os.path.join(HISTORIAL_DIR, info["slug"])
    if not os.path.isdir(directorio):
        return {"equipo": info, "ciclos": []}

    ciclos = {}
    for nombre in os.listdir(directorio):
        if not nombre.endswith(EXTENSIONES) or not nombre.startswith("ciclo_"):
            continue
        base, extension = os.path.splitext(nombre)
        identificador = base[len("ciclo_"):].split("_")[0]
        if not identificador.isdigit():
            continue
        ruta = os.path.join(directorio, nombre)
        actual = ciclos.get(int(identificador))
        candidato = {
            "id_ciclo": int(identificador),
            "estado_archivo": extension.lstrip("."),
            "bytes": os.path.getsize(ruta),
            "modificado": os.path.getmtime(ruta),
        }
        if actual is None or candidato["modificado"] > actual["modificado"]:
            ciclos[int(identificador)] = candidato

    return {
        "equipo": info,
        "ciclos": sorted(ciclos.values(), key=lambda c: c["id_ciclo"]),
    }


@RouterHistoricoCrudo.get("/{id_ciclo}/{equipo}")
def descargar_historial_crudo(
    id_ciclo: int,
    equipo: str,
    separador: str = Query(
        ";",
        description=(
            "Separador de columnas. ';' abre directo en Excel configurado en "
            "espanol; ',' es el estandar para procesar con pandas o R."
        ),
    ),
    incluir_temp_prod: bool = Query(
        True,
        description=(
            "temp_prod se conserva en el JSONL crudo pero no se publica en los "
            "informes productivos. Poner en false para excluirla tambien aqui."
        ),
    ),
):
    """
    Descarga el historial crudo de un ciclo convertido a CSV.

    El archivo original no se modifica, no se mueve y no se borra: solo se lee.

    Ejemplos:
        GET /historico/crudo/4120/l1_cocina_3
        GET /historico/crudo/4120/3
        GET /historico/crudo/4120/PF-L1-COCINA-3
        GET /historico/crudo/4120/l1_cocina_3?separador=,
    """
    if separador not in (";", ","):
        raise HTTPException(
            status_code=400,
            detail="El separador debe ser ';' o ','.",
        )

    if id_ciclo <= 0:
        raise HTTPException(
            status_code=400,
            detail="El id de ciclo debe ser un entero positivo.",
        )

    info = _resolver_equipo(equipo)
    ruta, extension = _buscar_archivo_ciclo(info, id_ciclo)

    columnas = tuple(
        campo
        for campo in CAMPOS_HISTORIAL
        if incluir_temp_prod or campo != "temp_prod"
    )

    nombre_descarga = f"ciclo_{id_ciclo}_{info['slug']}.csv"

    return StreamingResponse(
        _filas_csv(ruta, separador, columnas),
        media_type="text/csv; charset=utf-8",
        headers={
            "Content-Disposition": (
                f'attachment; filename="{nombre_descarga}"'
            ),
            # Informativo para el frontend: permite avisar que el ciclo
            # todavia no cerro o que quedo un cierre a medio hacer.
            "X-Origen-Archivo": extension.lstrip("."),
            "X-Equipo": info["slug"],
            "X-Id-Ciclo": str(id_ciclo),
        },
    )