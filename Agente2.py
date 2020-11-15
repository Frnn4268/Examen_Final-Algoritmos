#Importe de librerias necesarias para el correcto funcionamiento del programa
import win32com.client
import os
import time

while True: #Creación del "while" que eliminará datos infinitos de MSMQ

    qinfo=win32com.client.Dispatch("MSMQ.MSMQQueueInfo") #Apertura del MSMQ 
    computer_name = os.getenv('COMPUTERNAME')
    qinfo.FormatName="direct=os:"+computer_name+"\\PRIVATE$\\datos umg"
    queue=qinfo.Open(1,0)   # Open a ref to queue to read(1)
    msg=queue.Receive()
    print(" ____________________________________________________")
    print("|Título del dato:",msg.Label) #Tiempo de espera para el usuario
    print("|")
    print("|Cuerpo del dato:",msg.Body)
    print(" ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯")
    time.sleep(2)
    queue.Close()