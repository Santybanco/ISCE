import os
from config.configuracion import (
    MES_TRABAJO,
    RUTA_ONEDRIVE_BASE,
    CARPETA_INSUMOS_INDICADORES,
    SUBCARPETA_MONITOREO_CONTABLE
)

# ==============================
# RUTAS PRINCIPALES DEL PROYECTO
# ==============================

RUTA_PROYECTO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RUTA_DATOS = os.path.join(RUTA_PROYECTO, "datos")
RUTA_SALIDA = os.path.join(RUTA_DATOS, "salida")
RUTA_TEMPORALES = os.path.join(RUTA_DATOS, "temporales")


# ==============================
# CREAR CARPETAS SI NO EXISTEN
# (solo para salida / temporales)
# ==============================

def crear_carpetas_si_no_existen():
    os.makedirs(RUTA_SALIDA, exist_ok=True)
    os.makedirs(RUTA_TEMPORALES, exist_ok=True)


# ==============================
# RUTA DE INSUMOS (SHAREPOINT)
# ==============================

def obtener_ruta_insumos_mes():
    """
    Retorna la ruta donde están los archivos de entrada
    del mes contable en SharePoint (OneDrive sincronizado).
    """
    ruta = os.path.join(
        RUTA_ONEDRIVE_BASE,
        CARPETA_INSUMOS_INDICADORES,
        MES_TRABAJO,
        SUBCARPETA_MONITOREO_CONTABLE
    )

    if not os.path.isdir(ruta):
        raise Exception(
            f"No se encontró la carpeta de insumos del mes:\n{ruta}"
        )

    return ruta


# ==============================
# BUSCAR ARCHIVOS POR COINCIDENCIA
# ==============================

def obtener_archivo_por_coincidencia(parte_nombre_archivo):
    """
    Busca un archivo dentro de la carpeta de insumos
    del mes (SharePoint) usando coincidencia parcial.
    """

    carpeta = obtener_ruta_insumos_mes()
    archivos = os.listdir(carpeta)

    coincidencias = []

    for archivo in archivos:
        nombre_archivo = archivo.lower()
        parte_buscada = parte_nombre_archivo.lower()

        if parte_buscada in nombre_archivo and archivo.endswith((".xlsx", ".xls")):
            coincidencias.append(archivo)

    if len(coincidencias) == 0:
        raise Exception(
            f"No se encontró ningún archivo que coincida con '{parte_nombre_archivo}' "
            f"en la carpeta:\n{carpeta}"
        )

    if len(coincidencias) > 1:
        raise Exception(
            f"Se encontraron varios archivos que coinciden con '{parte_nombre_archivo}':\n"
            + "\n".join(coincidencias)
        )

    return os.path.join(carpeta, coincidencias[0])


# ==============================
# RUTAS DE SALIDA
# ==============================

def obtener_ruta_salida(nombre_archivo):
    crear_carpetas_si_no_existen()
    return os.path.join(RUTA_SALIDA, nombre_archivo)


def obtener_ruta_temporal(nombre_archivo):
    crear_carpetas_si_no_existen()
    return os.path.join(RUTA_TEMPORALES, nombre_archivo)
