# ==============================
# CONFIGURACIÓN GENERAL DEL PROYECTO ISCE
# ==============================

# --------------------------------------
# MES DE TRABAJO (MES VENCIDO)
# Este será el nombre de la hoja en el archivo de salida
# Ejemplo: "Enero", "Febrero", "Marzo"
# --------------------------------------
MES_TRABAJO = "Enero"


# --------------------------------------
# ARCHIVO DE SALIDA PRINCIPAL
# --------------------------------------
NOMBRE_ARCHIVO_SALIDA = "Indicadores_Operacion.xlsx"


# --------------------------------------
# ARCHIVOS DE ENTRADA - NOMBRE BASE
# Se buscarán por coincidencia parcial
# --------------------------------------
NOMBRE_BASE_ALCON = "Indicadores_ALCON_"
NOMBRE_BASE_CERTIFICACION = "Historico Indicador Certificación Gerentes"


# --------------------------------------
# HOJAS DE LECTURA EN ARCHIVOS DE ENTRADA
# --------------------------------------
HOJA_ALCON = "Detalle_Bancolombia"
# Si luego certificación requiere hoja específica, la definimos aquí


# --------------------------------------
# COLUMNAS A EXTRAER - ALCON
# --------------------------------------
COLUMNAS_ALCON = [
    "Gerencia",
    "Cantidad alertas",
    "Alertas con reproceso",
    "Alertas sin reproceso",
    "Calidad Gerencia"
]


# --------------------------------------
# COLUMNAS A EXTRAER - CERTIFICACIÓN GERENTES
# --------------------------------------
COLUMNAS_CERTIFICACION = [
    "GERENCIA",
    "FECHA CERTIFICACIÓN",
    "FECHA OBJETIVO",
    "INDICADOR"
]


# --------------------------------------
# CELDAS DE DESTINO EN EL ARCHIVO DE SALIDA
# --------------------------------------
CELDA_INICIO_ALCON = "A83"
CELDA_INICIO_CERTIFICACION = "A180"


# --------------------------------------
# MENSAJES DEL SISTEMA
# --------------------------------------
MENSAJE_INICIO = "Ejecutando indicadores de salud contable"
MENSAJE_CONFIRMACION = "Comenzaremos la ejecución de indicadores de salud contable del mes: "


# ==============================
# INFORME CUENTAS TEMPORALES
# ==============================

NOMBRE_BASE_TEMPORALES = "Informe Cuentas Temporales"

HOJA_TEMPORAL_REFERENCIA = "Temporales"
HOJA_SABANA_REFERENCIA = "Detalle Temporales"

COLUMNAS_TEMPORAL = [
    "gerencia_responsable",
    "SALDO CONTABLE",
    "PARTIDAS FUERA DE POLITICA_y"
]

COLUMNAS_SABANA = [
    "gerencia_responsable",
    "FUERA DE POLITICA",
    "VALOR PARTIDA PESOS"
]

CELDA_INICIO_TD_SALDO = "A4"
CELDA_INICIO_TD_SABANA = "E4"

# ==============================
# INFORME CXC
# ==============================

NOMBRE_BASE_CXC = "Informe CxC"

HOJA_CXC_REFERENCIA = "cxc"
HOJA_SABANA_CXC_REFERENCIA = "Detalle CXC"

COLUMNAS_CXC = [
    "gerencia_responsable",
    "SALDO CONTABLE",
    "PARTIDAS FUERA DE POLITICA_y"
]

COLUMNAS_SABANA_CXC = [
    "gerencia_responsable",
    "VALOR PARTIDA PESOS",
    "FUERA DE CICLO"
]

CELDA_INICIO_CXC_SALDO = "A36"
CELDA_INICIO_CXC_SABANA = "E36"

# ==============================
# RUTAS SHAREPOINT / ONEDRIVE
# ==============================

# Ruta base del SharePoint sincronizado por OneDrive
RUTA_ONEDRIVE_BASE = r"C:\Users\santcord\OneDrive - Grupo Bancolombia\Administrativo_M365 - Indicadores Operación"

# Carpeta donde están los insumos por mes
CARPETA_INSUMOS_INDICADORES = "Insumos indicadores"

# Subcarpeta que contiene los archivos a procesar
SUBCARPETA_MONITOREO_CONTABLE = "Monitoreo contable"

# ==================================================
# REPROCESO HISTÓRICO DE MESES (OPCIONAL)
# ==================================================
# - Usar SOLO cuando se necesite poner al día el Excel
# - Si es None -> ejecución normal (un solo mes)
# - Si es lista -> ejecuta ISCE para cada mes de la lista
#
# Ejemplo:
# MESES_REPROCESO = ["Enero", "Febrero"]
# ==================================================

MESES_REPROCESO = ["Enero", "Febrero"]

