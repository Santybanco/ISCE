import tkinter as tk
from tkinter import messagebox



try:
    import pyttsx3
    engine = pyttsx3.init()
except:
    pyttsx3=None
    engine= None



# ==============================
# CONFIGURACIÓN DE VOZ
# ==============================

motor_voz = engine
if motor_voz:
    motor_voz.setProperty("rate",170)

# ==============================
# FUNCIONES DE MENSAJES
# ==============================

def hablar(texto):
    print(texto)

    if engine:
        try:
            engine.say(texto)
            engine.runAndwait()
        except:
            pass


def confirmar_inicio(mensaje):
    """
    Muestra una ventana de confirmación.
    Retorna True si el usuario acepta, False si cancela.
    """
    ventana = tk.Tk()
    ventana.withdraw()

    respuesta = messagebox.askyesno("Confirmación de ejecución", mensaje)

    ventana.destroy()
    return respuesta


def mostrar_info(titulo, mensaje):
    """
    Muestra una ventana informativa.
    """
    ventana = tk.Tk()
    ventana.withdraw()

    messagebox.showinfo(titulo, mensaje)

    ventana.destroy()


def mostrar_error(titulo, mensaje):
    """
    Muestra una ventana de error.
    """
    ventana = tk.Tk()
    ventana.withdraw()

    messagebox.showerror(titulo, mensaje)

    ventana.destroy()