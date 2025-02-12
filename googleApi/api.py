from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from persistence.Informe1 import llenarInforme1
import os
from util.Pdf import Pdf

class GoogleSheetsAExcel:
    def __init__(self, service_account_file: str, spreadsheet_id: str, range_name: str, plantilla_path: str, drive_folder_id: str) -> None:
        """
        Constructor de la clase.

        """
        self.service_account_file = service_account_file
        self.spreadsheet_id = spreadsheet_id
        self.range_name = range_name
        self.plantilla_path = plantilla_path
        self.drive_folder_id = drive_folder_id
        self.scopes = [
            'https://www.googleapis.com/auth/spreadsheets.readonly',
            'https://www.googleapis.com/auth/drive'
        ]
        self.credentials = None
        self.sheet_service = None
        self.drive_service = None

    def inicializarServicios(self):
        """
        Inicializa los servicios de Google Sheets y Google Drive con las credenciales
        especificadas en el constructor.

        """
        self.credentials = service_account.Credentials.from_service_account_file(
            self.service_account_file, scopes=self.scopes
        )
        self.sheet_service = build('sheets', 'v4', credentials=self.credentials)
        self.drive_service = build('drive', 'v3', credentials=self.credentials)

    def fetchDatos(self):
        """
        Obtiene los datos de la hoja especificada en la definición de la clase.

        """
        result = self.sheet_service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id, range=self.range_name
        ).execute()
        values = result.get('values', [])
        if not values:
            raise ValueError("No se encontraron datos en el rango especificado.")
        columnas = values[0]
        df = pd.DataFrame(values[1:], columns=columnas)
        if 'data-fecha' in df.columns:
            df['data-fecha'] = pd.to_datetime(df['data-fecha'], errors='coerce')
        return df

    def obtenerIdCarpetaFecha(self, nombre_carpeta):
        """
        Obtiene el ID de la carpeta en Google Drive con el nombre especificado.

        """

        query = f"'{self.drive_folder_id}' in parents and name = '{nombre_carpeta}' and mimeType = 'application/vnd.google-apps.folder'"
        respuesta = self.drive_service.files().list(q=query).execute()
        archivos = respuesta.get('files', [])
        return archivos[0]['id'] if archivos else None

    def crearCarpeta(self, nombre_carpeta):
        """
        Crea una carpeta en Google Drive y retorna su ID.

        """

        metadata = {
            'name': nombre_carpeta,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [self.drive_folder_id]
        }
        carpeta = self.drive_service.files().create(body=metadata, fields='id').execute()
        return carpeta.get('id')

    def subirArchivo(self, file_path, nombre_archivo, folder_id):
        """
        Sube un archivo a la carpeta de Google Drive especificada.

        """
        archivo_metadata = {
            'name': nombre_archivo,
            'parents': [folder_id]
        }
        media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.drive_service.files().create(body=archivo_metadata, media_body=media, fields='id').execute()
        print(f"Archivo {nombre_archivo} subido exitosamente a la carpeta {folder_id}.")

    def archivoExiste(self, nombre_archivo, folder_id) -> bool:
        """
        Verifica si un archivo PDF existe en la carpeta de Google Drive.
        """
        query = f"'{folder_id}' in parents and name = '{nombre_archivo}' and mimeType = 'application/pdf'"
        respuesta = self.drive_service.files().list(q=query, fields="files(id, name)").execute()
        archivos = respuesta.get('files', [])
        return len(archivos) > 0


    def llenarYSubirPlantillas(self):
        """
        Llena las plantillas con los datos de Google Sheets y las sube a Drive en formato PDF.
        Verifica antes que no existan duplicados en Drive.
        """
        datos = self.fetchDatos()
        pdfConv = Pdf()

        for index, fila in datos.iterrows():
            fecha = pd.to_datetime(fila['data-info_general-fecha'])
            fecha_str = fecha.strftime('%Y-%m-%d') if pd.notna(fecha) else 'SinFecha'

            # Verificar o crear carpeta
            folder_id = self.obtenerIdCarpetaFecha(fecha_str)
            if not folder_id:
                print(f"Carpeta {fecha_str} no encontrada. Creando...")
                folder_id = self.crearCarpeta(fecha_str)

            # Nombre del archivo PDF esperado
            nombre_pdf = f'plantilla_{index + 1}_{fecha_str}.pdf'

            # Verificar si el archivo PDF ya existe en Drive
            if self.archivoExiste(nombre_pdf, folder_id):
                print(f"El archivo {nombre_pdf} ya existe en la carpeta {fecha_str}. Saltando llenado y subida...")
                continue

            # Si no existe, llenar la plantilla y convertir a PDF
            print(f"Llenando datos para la fila {index + 1}...")
            wb = load_workbook(self.plantilla_path)
            ws = wb.active
            llenarInforme1(ws, fila)

            # Guardar archivo Excel localmente
            nombre_excel = f'plantilla_{index + 1}_{fecha_str}.xlsx'
            ruta_excel = os.path.join('./', nombre_excel)
            wb.save(ruta_excel)
            print(f"Archivo Excel guardado localmente: {nombre_excel}")

            # Convertir Excel a PDF
            ruta_pdf = os.path.join('./', nombre_pdf)
            pdfConv.excelPdf(ruta_excel, ruta_pdf)
            print(f"Archivo convertido a PDF: {ruta_pdf}")

            os.remove(ruta_excel)  # Eliminar el archivo Excel local temporal

            # Subir archivo PDF a Drive
            self.subirArchivo(ruta_pdf, nombre_pdf, folder_id)
            os.remove(ruta_pdf)  # Eliminar el archivo PDF local después de subirlo
            print(f"Archivo PDF {nombre_pdf} subido y eliminado localmente.")





