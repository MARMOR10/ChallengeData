import pickle
import os
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
import datetime 
from dataclasses import dataclass
from datetime import datetime, timedelta
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mysql.connector
import mysql
from mimetypes import MimeTypes
from pickle import TRUE
from re import T
from urllib import response
import io
from apiclient import errors
import sys
from google.oauth2 import service_account
from importlib.machinery import FileFinder
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from webbrowser import get
from attr import field
import mysql.connector
import mysql
from numpy import delete, insert, place
from sklearn import tree

#creación interfaz gráfica
windows=Tk()
windows.title("Aplicación Inventario Archivos Drive")
windows.geometry("1700x800")
windows.configure(bg='#0059b3')
label = Label(windows, text="INVENTARIO ARCHIVOS DRIVE")
label.pack(anchor=CENTER)
label.config(fg="white",bg='#0059b3',font=("Verdana",30)) 
fileid=StringVar()
fname=StringVar()
fprop=StringVar()
fext=StringVar()
fvisi=StringVar()
fdate=StringVar()

#código google
def Create_Service(client_secret_file, api_name, api_version, *scopes):
    print(client_secret_file, api_name, api_version, scopes, sep='-')
    CLIENT_SECRET_FILE = client_secret_file
    API_SERVICE_NAME = api_name
    API_VERSION = api_version
    SCOPES = [scope for scope in scopes[0]]
    print(SCOPES)

    cred = None

    pickle_file = f'token_{API_SERVICE_NAME}_{API_VERSION}.pickle'
    # print(pickle_file)

    if os.path.exists(pickle_file):
        with open(pickle_file, 'rb') as token:
            cred = pickle.load(token)

    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            cred.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            cred = flow.run_local_server()

        with open(pickle_file, 'wb') as token:
            pickle.dump(cred, token)

    try:
        service = build(API_SERVICE_NAME, API_VERSION, credentials=cred)
        print(API_SERVICE_NAME, 'service created successfully')
        return service
    except Exception as e:
        print('Unable to connect.')
        print(e)
        return None

def convert_to_RFC_datetime(year=1900, month=1, day=1, hour=0, minute=0):
    dt = datetime.datetime(year, month, day, hour, minute, 0).isoformat() + 'Z'
    return dt

#función para enviar correo:
def enviomail(nombre,propietario):
    CLIENT_SECRET_FILE = 'client_secret.json'
    API_NAME = 'gmail'
    API_VERSION = 'v1'
    SCOPES = ['https://mail.google.com/']
    service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
    emailMsg = 'El archivo'+ ' ' + nombre +' '+ 'se encontraba público'+' '+'por esto se cambiaron los permisos de este a restringido'
    mimeMessage = MIMEMultipart()
    mimeMessage['to'] = propietario
    mimeMessage['subject'] = 'Se restringieron los permisos de su archivo'+ ' ' +nombre
    mimeMessage.attach(MIMEText(emailMsg, 'plain'))
    raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()
    service.users().messages().send(userId='me', body={'raw': raw_string}).execute()

#función para borrar los permisos del archivo:
def borrarpermisos(id):
    CLIENT_SECRET_FILE = 'client_secret.json'
    API_NAME = 'drive'
    API_VERSION = 'v3'
    SCOPES = ['https://www.googleapis.com/auth/drive']
    service_drive = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
    service_drive.permissions().delete(fileId=id, permissionId="anyoneWithLink").execute() 
    

#Función funcionamiento aplicación:

def aplicacion():
    #autentica para enviar correos desde la cuenta gmail
    CLIENT_SECRET_FILE = 'client_secret.json'
    API_NAME = 'drive'
    API_VERSION = 'v3'
    SCOPES = ['https://www.googleapis.com/auth/drive']
    service_drive = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
    
#creación base de datos#:
    #cursorr = conexion.cursor()
    #creo la bd donde se va a guardar los datos asociados a los archivos de drive
    #cursorr.execute("CREATE DATABASE InventarioDrive")

#conexión base de datos
    conexionbdinventario = mysql.connector.connect(
        host = "localhost",
        user = "root",
        password= "",
        database = 'InventarioDrive'   
        #port = 3306,
        )

    print(conexionbdinventario)
    cursorbdinventario = conexionbdinventario.cursor()

#creación tablas bases de datos:
    #sql="""CREATE TABLE Inventario_Historico (ID VARCHAR(255), Archivo VARCHAR(255), Propietario VARCHAR(255), Extension VARCHAR(255), Visibilidad VARCHAR(255), Fecha_Modificacion VARCHAR(255))"""
    #cursorbdinventario.execute(sql)
    #conexionbdinventario.commit()
    #print('creo bd inventario historico')


#tomo los id y fecha de los archivos que se encuentran ya inventariados
    cursorbdinventario.execute("SELECT ID,Fecha_Modificacion FROM Inventario")
    filessavebd=cursorbdinventario.fetchall()
    listidsavebd=[]
    cursorbdinventario.close()

    for archivo in filessavebd:
        idfilesaved = archivo[0]
        datefilesaved= archivo[1]
        listidsavebd.append(archivo[0])

#tomo los ids y nombre de los archivos que se encuentran en drive para recorrer cada uno y poder tomar los campos necesarios a guardar en la BD
    results =  service_drive.files().list( fields = "nextPageToken, files(id,name)").execute()
    items = results.get('files',[])

    for item in items:
        infofile=service_drive.files().get(fileId=item['id'], fields="*",).execute()
        namefile=infofile['name']
#obtengo los datos relacionados a cada archivo dentro de drive:
        idfile=infofile['id']
        namefile=infofile['name']
        extfile=infofile['mimeType']
        extensionfile=(extfile[28 : ])
        datemodificfile=infofile['modifiedTime']
        datospropietario=infofile['owners']
        datoprop=datospropietario[0]
        emailpropietario=datoprop['emailAddress']
        idspermisos=infofile['permissionIds']
#definición de variable para identificar cuando un archivo esta público:
        publico="anyoneWithLink"
#valido que el id no se encuentre los datos que ya fueron guardados
        if item['id'] not in listidsavebd:
#reviso en los permisos del archivo si este se encuentra público
            if publico in idspermisos:
                identificador=item['id']
                enviomail(namefile,emailpropietario)
                borrarpermisos(identificador) 
                cursorbdinventario = conexionbdinventario.cursor()
                cursorbdinventario.execute("INSERT INTO Inventario_Historico (ID, Archivo, Propietario, Extension, Visibilidad, Fecha_Modificacion) VALUES ('{}','{}','{}','{}','{}','{}')".format(idfile,namefile,emailpropietario,extensionfile,'Público',datemodificfile))
                conexionbdinventario.commit()
                cursorbdinventario.close()
            else:            
#inserta los datos de los archivos que se encuentran privados
                cursorbdinventario = conexionbdinventario.cursor()
                cursorbdinventario.execute("INSERT INTO Inventario (ID, Archivo, Propietario, Extension, Visibilidad, Fecha_Modificacion) VALUES ('{}','{}','{}','{}','{}','{}')".format(idfile,namefile,emailpropietario,extensionfile,'Privado',datemodificfile))
                conexionbdinventario.commit()
                cursorbdinventario.close()
#reviso si el los permisos de los archivos ya guardados no se encuentran como público                
        else:
            if publico in idspermisos:
#si es público el archivo se enviara el mensaje al propietario una notificación a su correo
                identificador=item['id']
                enviomail(namefile,emailpropietario)
                borrarpermisos(identificador)
                cursorbdinventario = conexionbdinventario.cursor()
                cursorbdinventario.execute("INSERT INTO Inventario_Historico (ID, Archivo, Propietario, Extension, Visibilidad, Fecha_Modificacion) VALUES ('{}','{}','{}','{}','{}','{}')".format(idfile,namefile,emailpropietario,extensionfile,'Público',datemodificfile))
                conexionbdinventario.commit()
                cursorbdinventario.close()
            else:
#tomo la fecha de modificación del archivo para verificar si esta cambio o no, con el fin de hacer la actualización respectiva en la BD
                cursorbdinventario = conexionbdinventario.cursor()
                idfile=item['id']
                datemodificfile=infofile['modifiedTime']  
#trae el valor de la fecha de modificación para cada archivo y la compara con la actualizada:
                sql = "SELECT Fecha_Modificacion FROM Inventario WHERE ID = %s"
                adr = [(idfile)]
                cursorbdinventario.execute(sql,adr)
                myresult = cursorbdinventario.fetchone()
                datesavedbd=myresult[0]        
                cursorbdinventario.close()
#valido si la fecha actualizada es igual a la fecha que estaba guardada no modifique nada y sino que actualice la fecha en la BD
                if datemodificfile == datesavedbd :
                    print('no guardar')
                else: 
                    cursorbdinventario = conexionbdinventario.cursor()
                    idarchivo=idfile
                    datenew=datemodificfile
                    sqlupdate= ("UPDATE Inventario SET Fecha_Modificacion = %s WHERE ID = %s")
                    valores=(datenew,idarchivo)
                    cursorbdinventario.execute(sqlupdate,valores)
                    conexionbdinventario.commit()
                    cursorbdinventario.close()
                
#función para salir de la interfaz
def exit():
    mensaje=messagebox.askquestion("Salir","¿Desea salir de la aplicación inventario?")
    if mensaje=="yes":
        windows.destroy()
#función para mostrar los datos de la tabla inventario donde se encuentran todos los archivos
def mostrardatos():
    conexionbdinventario = mysql.connector.connect(host = "localhost", user = "root", password= "", database = 'InventarioDrive' )
    print(conexionbdinventario)
    cursorbdinventario = conexionbdinventario.cursor()
    registros=tree.get_children()
    for elemento in registros:
        tree.delete(elemento)
    try:
        cursorbdinventario.execute("SELECT * FROM Inventario ")    
        for row in cursorbdinventario:
            tree.insert("",0,text=row[0],values=(row[1],row[2],row[3],row[4],row[5]))
    except:
        pass

#función para mostrar los datos de la tabla inventario_historico, donde se encuentran todos los archivos que alguna vez fueron públicos
def mostrardatoshistorico():
    conexionbdinventario = mysql.connector.connect(host = "localhost", user = "root", password= "", database = 'InventarioDrive' )
    print(conexionbdinventario)
    cursorbdinventario = conexionbdinventario.cursor()
    registros=tree.get_children()
    for elemento in registros:
        tree.delete(elemento)
    try:
        cursorbdinventario.execute("SELECT * FROM Inventario_Historico ")    
        for row in cursorbdinventario:
            tree.insert("",0,text=row[0],values=(row[1],row[2],row[3],row[4],row[5]))
    except:
        pass    

#creación de botones y campos tabla:
tree=ttk.Treeview(height=70,columns=('#0','#1','#2','#3','#4'))
tree.place(x=0,y=150)
tree.heading('#0', text="ID",anchor=CENTER,)
tree.heading('#1', text="Nombre",anchor=CENTER)
tree.heading('#2', text="Propietario",anchor=CENTER)
tree.heading('#3', text="Extensión",anchor=CENTER)
tree.heading('#4', text="Visibilidad",anchor=CENTER)
tree.heading('#5', text="Fecha Modificación",anchor=CENTER)
tree.column('#4',width=270)
tree.column('#0',width=278)
tree.column('#1',width=280)
tree.column('#3',width=255)
tree.column('#5',width=250)
botonactualizar=Button(windows,text="Actualizar Inventario",bg="white",relief=RAISED, command=aplicacion)
botonactualizar.place(x=330,y=75)
botoninventario=Button(windows,text="Inventario Archivos",bg="white",relief=RAISED, command=mostrardatos)
botoninventario.place(x=580,y=75)
botoninventariohistorico=Button(windows,text="Inventario Historico Archivos Públicos",bg="white",relief=RAISED, command=mostrardatoshistorico)
botoninventariohistorico.place(x=780,y=75)
botonsalir=Button(windows,text="Salir Aplicación", bg="white",relief=RAISED,command=exit)
botonsalir.place(x=1100,y=75)

#ejecución de la interfaz
windows.mainloop()