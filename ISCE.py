import time
import config.configuracion as config
import config.rutas as rutas
import procesadores.procesador as proc

from utils.mensajes import confirmar_inicio, mostrar_info, mostrar_error

from procesadores.procesador import (
    procesar_alcon,
    procesar_certificacion_gerentes,
    procesar_temporales_td_saldo,
    procesar_temporales_td_sabana,
    procesar_cxc,
    procesar_cxp
)

# ==============================
# FUNCIÓN PRINCIPAL DEL SISTEMA
# ==============================

def ejecutar_indicadores():
    """
    Ejecuta el flujo ISCE para uno o varios meses contables,
    mostrando progreso y tiempo de ejecución.
    """

    inicio = time.time()

    # ------------------------------
    # Determinar meses a ejecutar
    # ------------------------------
    if config.MESES_REPROCESO:
        meses_a_ejecutar = config.MESES_REPROCESO
        mensaje_inicio = (
            "Se ejecutarán los indicadores de salud contable para los meses:\n\n"
            + ", ".join(meses_a_ejecutar)
        )
    else:
        meses_a_ejecutar = [config.MES_TRABAJO]
        mensaje_inicio = f"{config.MENSAJE_CONFIRMACION}{config.MES_TRABAJO}"

    confirmar = confirmar_inicio(mensaje_inicio)
    if not confirmar:
        return

    try:
        # ------------------------------
        # Ejecución por mes
        # ------------------------------
        for mes in meses_a_ejecutar:

            print(f"\n==============================")
            print(f"▶ Iniciando procesamiento mes: {mes}")
            print(f"==============================")

            # Sincronizar mes en todos los módulos
            config.MES_TRABAJO = mes
            proc.MES_TRABAJO = mes
            rutas.MES_TRABAJO = mes

            print("• Procesando ALCON...")
            procesar_alcon()

            print("• Procesando Certificación de Gerentes...")
            procesar_certificacion_gerentes()

            print("• Procesando Temporales - TD Saldo...")
            procesar_temporales_td_saldo()

            print("• Procesando Temporales - TD Sábana...")
            procesar_temporales_td_sabana()

            print("• Procesando CXC...")
            procesar_cxc()

            print("• Procesando CXP...")
            procesar_cxp()

            print(f"✅ Mes {mes} procesado correctamente")

        # ------------------------------
        # Fin de ejecución
        # ------------------------------
        fin = time.time()
        duracion = fin - inicio

        hh_mm_ss = time.strftime("%H:%M:%S", time.gmtime(duracion))

        mostrar_info(
            "Proceso completado",
            "Los indicadores fueron procesados correctamente.\n\n"
            f"Meses procesados: {', '.join(meses_a_ejecutar)}\n"
            f"Tiempo total: {hh_mm_ss}"
        )

    except Exception as e:
        mostrar_error(
            "Error en la ejecución",
            f"Ocurrió un error durante el procesamiento:\n\n{str(e)}"
        )


# ==============================
# EJECUCIÓN DEL PROGRAMA
# ==============================

if __name__ == "__main__":
    ejecutar_indicadores()
