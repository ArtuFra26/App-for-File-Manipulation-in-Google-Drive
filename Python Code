######################### Install all dependencies and test Google API credentials #########################
#
#
############################################################################################################

# Bibliotecas a instalar mediante pip
!pip install gspread
!pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
!pip install shutil
!pip install pandas

# Importar bibliotecas requeridas para manipular datos
import os
import io
import shutil
import gspread
import json
import pandas as pd

# Bibliotecas para interactuar con la API de Google
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

# Definir los scopes de permisos para interactuar con la sesión de Drive de la cuenta de jurídico
SCOPES = ['https://www.googleapis.com/auth/drive.metadata',
          'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/drive.file',
          'https://www.googleapis.com/auth/spreadsheets']

#Definir función de importación de archivos 
creds = None

# Acceder a la API de Google mediante las credenciales de acceso de la cuenta de jurídico
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# Si las credenciales no son válidas acceder mediante autenticación
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Guardar las credenciales para la siguiente autenticación
        with open('token.json', 'w') as token:token.write(creds.to_json())

# Generar una conexión con la aplicación Drive
service = build('drive', 'v3', credentials=creds)


######################### Import the list of IDs of files stored in Drive to be dowloaded #########################
#
#
###################################################################################################################

# Generar conexión con aplicación de SpreadSheets
client = gspread.authorize(creds)

# Abrir hoja de cálculo que almacena los datos de la importación
spreadsheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1Mu39myPJzlvfzEpH3LefTGsskRAVKmHzVF07hwNVxeY/edit#gid=0')

# Definir primera página como dataset de trabajo
worksheet = spreadsheet.get_worksheet(0)

# Almacenar todos los registros en formato JSON
records_data = worksheet.get_all_records()

#Convertir dataset a objeto de tipo dataframe
records_df = pd.DataFrame.from_dict(records_data)

#Definir lista de IDs para importación
list_urls = records_df["ID"] 


######################### Execute download in local path and store the path in each file in list  #############################
#
#
###############################################################################################################################

# Generar lista para almacenar paths de archivos
adrs = []

# Implementar descarga de forma recursiva
for id_url in list_urls:
    # Recuperar archivo de Drive e iniciar la descarga
    file_id = id_url
    results = service.files().get(fileId=file_id).execute()
    name = results['name']
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    # Informar sobre avance de descarga del archivo
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))

    fh.seek(0)
    # Incorporar los datos de la descarga a contenedor en archivo local
    path = "******************************" + name
    with open(path, 'wb') as f:
        shutil.copyfileobj(fh, f)

    #Incorporar datos de dirección del archivo descargado
    adrs.append(
        {
            "Archivo": file_id,
            "Path": path
        }
    )

# Objeto de tipo dataframe con lista
adrs = pd.DataFrame(adrs)

################################# Write local paths back in Google Sheets so domain user know where to locate files #####################################
#
#
#########################################################################################################################################################

# Definir una función que haga conteo de cuál es la última columna ocupada en la hoja de trabajo
def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

# Obtener todos los valores observados en la hoja de trabajo
values = worksheet.get_all_values()

# Definir cuál es la última columna utilizada
col = colnum_string(max([len(r) for r in values]) + 1)

#Incorporar valores de path y ID en hoja de Drive
worksheet.update(values = adrs.values.tolist(), range_name = col + '2', value_input_option = 'USER_ENTERED')
