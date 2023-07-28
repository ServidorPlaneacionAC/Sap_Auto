import utils_sap as utils

ruta_ingreso=r"D:\Nutresa\Proyectos\Scripts SAP\ingreso_sap.vbs"
lista_tareas = [
    ("08:40", r"D:\Nutresa\Proyectos\Scripts SAP\Descargar LM.vbs"),
    # ("HH:MM", r"ruta sap scripting")
    # Agrega aquí más tareas en el formato ("HH:MM", "Nombre de la tarea")
]

utils.definir_tareas(ruta_ingreso=ruta_ingreso,lista_tareas=lista_tareas)
