import tkinter as tk
from tkinter import messagebox
import schedule
import time
import subprocess
import psutil



def tarea_programada(ruta_vbscript,ruta_ingreso):

    try:
        # Ejecutar el VBScript utilizando el programa wscript.exe
        if not sap_activo():
            subprocess.run(["wscript", ruta_ingreso], capture_output=True, text=True)
        subprocess.run(["wscript", ruta_vbscript], capture_output=True, text=True)

    except FileNotFoundError:
        print(f"No se pudo encontrar el archivo {ruta_vbscript}. Verifica la ruta y vuelve a intentarlo.")
    except subprocess.CalledProcessError as e:
        print(f"Ocurrió un error al ejecutar el VBScript: {e}")

def definir_tareas(ruta_ingreso,lista_tareas):

    # Programa la tarea para que se ejecute en el horario especificado
    # schedule.every().day.at(hora_fija).do(tarea_programada)
    for hora, tarea in lista_tareas:
        #Guarda en schedule las tareas programadas
        schedule.every().day.at(hora).do(tarea_programada, tarea, ruta_ingreso)

    while True:
        #Ciclo infinito que valida cada 30 segundos si hay algo pendiente por ejecutar guardado en schedule
        schedule.run_pending()
        time.sleep(30)

def sap_activo():
    process_names = ["sapshcut.exe", "sapgui.exe", "saplogon.exe"]
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] in process_names:
            return True
    return False

