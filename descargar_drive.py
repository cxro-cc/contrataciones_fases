import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# Configuración de credenciales
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
creds = service_account.Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
service = build('drive', 'v3', credentials=creds)

def descargar_archivo(file_id, file_name):
    # 1. Averiguamos qué tipo de archivo es
    file_metadata = service.files().get(fileId=file_id, fields='mimeType').execute()
    mime_type = file_metadata.get('mimeType')

    # 2. Si es un Google Sheet, lo exportamos a Excel
    if mime_type == 'application/vnd.google-apps.spreadsheet':
        request = service.files().export_media(
            fileId=file_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    # 3. Si ya es un archivo binario/Excel (como tu archivo de 37.6 MB), lo descargamos directo
    else:
        request = service.files().get_media(fileId=file_id)

    # El proceso de descarga en "chunks" maneja muy bien archivos pesados
    fh = io.FileIO(file_name, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    print(f"✓ Descargado: {file_name}")

# Reemplaza los IDs reales aquí abajo
archivos = {
    "1Axj6z0U1odyFcdDz7Rjdv2jOZgnh9yIxKAXiWzvXKW4": "Candidatos_UAS.xlsx",
    "1iSEdYOU1uDF7Gwnu0o9ABR9hv3pI1INnrJv2Yt9jbN0": "Candidatos_UPE.xlsx",
    "1VKPo1J0T5kF-0me6Yd3tz1-nKC5kI0My": "Unidades_Inauguradas_fases.xlsx"
}

for f_id, f_name in archivos.items():
    descargar_archivo(f_id, f_name)