import win32com.client as win32
import os
from datetime import datetime, timedelta


# Fecha
fecha_sin_hora = str((datetime.today()).strftime("%Y-%m-%d"))
dia = fecha_sin_hora.split("-")[2]
mes = fecha_sin_hora.split("-")[1]

#Folder
path = "C:/Users/z0a018o/Desktop/" + dia + "-" + mes

print("Ventas inventarios")
print("Carpeta creada del dia: " + fecha_sin_hora)
os.mkdir(path)

outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')

#Carpeta de outlook
inbox = outlook.GetDefaultFolder(6).Folders["Ventas inventarios"]

messages = inbox.Items

delete = []

for i in range(0,len(messages)):
	date = str(messages[i].SentOn).split(" ")[0]
	delete.append(messages[i])
	if date == fecha_sin_hora:
		for attachment in messages[i].Attachments:
			print(attachment)
			attachment.SaveAsFile(os.path.join(path,attachment.filename)) 

#Eliminar archivos del correo
for i in delete:
	i.Delete()

