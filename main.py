from googleApi.api import GoogleSheetsAExcel

if __name__ == "__main__":
    # Configuración
    SERVICE_ACCOUNT_FILE = r'googleApi\censos-maute-d48ff1e9060b.json'
    SPREADSHEET_ID = '1Z2HXcI6iDcO9JLBxB2cil3QC8mLNRuyzp-00uesukkk'
    RANGE_NAME = 'Sheet1!A1:AP10000'
    PLANTILLA_PATH = r'censos\FORMATO 1 IDENTIFICACIÓN - Aprobado.xlsx'
    DRIVE_FOLDER_ID = '1pe3OAni1kG5lO4XrB5juq9SRnjoDVSqT'

    # Inicializar y ejecutar
    servicio = GoogleSheetsAExcel(SERVICE_ACCOUNT_FILE, SPREADSHEET_ID, RANGE_NAME, PLANTILLA_PATH, DRIVE_FOLDER_ID)
    servicio.inicializarServicios()
    servicio.llenarYSubirPlantillas()