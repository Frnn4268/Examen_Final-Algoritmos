#Importación de librerias necesarias para el correcto funcionamiento del programa
import win32com.client
import os
from random import choice
import time
import json
#Bienvenida al usuario
print("            ______________________")
print("           |Microservicios UMG App|")
print("            ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯")
print("Bienvenido, esta es una simulación de como datos")
print("son ingresados y guardados en una Base de Datos.")
time.sleep(6)

database = open("Colas-DB.txt","a") #Creación del archivo donde se guardarán los JSON

while True: #Creación del "while" que generará archivos infinitos y aleatorios
#Creación de listas de datos aleatorios
    datos_aleatorios = ['Valorant', 'Call of Duty', 'Age of Empires', 'Proyecto', 'Autocad', 'TUF', 'ROG', 'Aura', 'DotMod', 'VGOD', 'KeyboardRGB', 'Chrome', 'Visual Studio', 'Ajustes', 'OPERA', 'Photoshop', 'Word', 'Carpeta', 'Documento1', 'TablaExcel', 'Calendario', 'Epic Launcher', 'FL Studio', 'MSI Afterburner', 'uTorrent', 'Steam', 'Explorer']

    extensiones_aleatorias = ['.docx', '.pdl', '.txt', '.exe', '.pdp', '.pdx', '.pef', '.pdlcp', '.jpg', '.gif', '.avi', '.mp4', '.iso', '.dll', '.img', '.wav', '.mwv', '.zip', '.rar', '.sys', '.bat', '.ibook', '.epub', '.bmp', '.db', '.fon']

    aleatorio = choice(datos_aleatorios) #Captura aleatoria de la primera lista

    aleatorio1 = choice(extensiones_aleatorias) #Captura aleatoria de la segunda lista

    datos = { #Inicio de la creación del JSON
        'Archivo': (aleatorio),
        'Extension': (aleatorio1)
    }

    dato_json = json.dumps(datos) #Creación del JSON
    #Datos para encolar a MSMQ
    qinfo=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
    computer_name = os.getenv('COMPUTERNAME')
    qinfo.FormatName="direct=os:"+computer_name+"\\PRIVATE$\\datos umg"
    queue=qinfo.Open(2,0)   # Open a ref to queue
    msg=win32com.client.Dispatch("MSMQ.MSMQMessage")
    msg.Label= "Archivos Aleatorios"
    msg.Body = (dato_json)
    print(" _____________________________")
    print("|Ingresando dato a la Database|")
    print(" ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯")
    print(" ________________________________")
    print("|Dato ingresado: " + aleatorio + aleatorio1) #Instrucciones para el usuario
    print(" ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯")
    msg.Send(queue)

    database = open("Colas-DB.txt","a") #Ingreso de datos a la Base de datos (Archivos de texto)

    database.write(dato_json + "\n")

    database.close()

    queue.Close() #Cierre de las colas MSMQ
    #Tiempo de espera para el usuario
    print("Dato ingresado con éxito...")
    time.sleep(1)
    print("Esperando...")
    print("")
    time.sleep(2)