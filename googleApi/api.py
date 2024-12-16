from googleapiclient.discovery import build
from google.oauth2 import service_account
import pandas as pd
from openpyxl import load_workbook
from persistence.Informe1 import llenarInforme1

class GoogleSheetsToExcel:
    def __init__(self, service_account_file, spreadsheet_id, range_name, plantilla_path):
        self.service_account_file = service_account_file
        self.spreadsheet_id = spreadsheet_id
        self.range_name = range_name
        self.plantilla_path = plantilla_path
        self.scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        self.credentials = None
        self.sheet_service = None

    def inicializarServicio(self):
        
        self.credentials = service_account.Credentials.from_service_account_file(
            self.service_account_file, scopes=self.scopes
        )
        self.sheet_service = build('sheets', 'v4', credentials=self.credentials)

    def fetchDatos(self):
        
        result = self.sheet_service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id, range=self.range_name
        ).execute()
        values = result.get('values', [])
        if not values:
            raise ValueError("No data found in the specified range.")
        columnas = values[0]  # First row contains column names
        df = pd.DataFrame(values[1:], columns=columnas)
        
        # Convertir la columna 'data-fecha' a datetime (ignorar errores si está vacía)
        if 'data-fecha' in df.columns:
            df['data-fecha'] = pd.to_datetime(df['data-fecha'], errors='coerce')
        
        return df

    def llenarPlantillas(self, datos):
        wb = load_workbook(self.plantilla_path)
        ws = wb.active

        for index, fila in datos.iterrows():
            print(f"Llenando datos para la fila {index + 1}...")
            llenarInforme1(ws, fila)  

            archivo_salida = f"plantilla_ejemplo_{index + 1}.xlsx"
            wb.save(archivo_salida)
            print(f"Archivo guardado: {archivo_salida}")


