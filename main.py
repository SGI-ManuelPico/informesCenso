from googleApi.api import GoogleSheetsAExcel

if __name__ == "__main__":
    # 1. Configuración de credenciales y rutas
    SERVICE_ACCOUNT_FILE = r'googleApi\censos-maute-d48ff1e9060b.json'
    SPREADSHEET_ID       = '1WLZu1vYe8MtihM4kGvRq5Dj3-A5QgtmkSuF151drtZA'
    DRIVE_FOLDER_ID      = '1kq_6eo-_u0fuCUOHDRV_V5bQ1c_dvfMF'

    # 2. Rangos
    RANGE_INFORME1 = "Sheet1!A1:EZ200"   
    RANGE_FICHA1   = "Sheet1!A1:EZ200" 
    RANGE_FICHA2   = "data-start_carac_poblacion-caracteristicas_poblacion!A1:H200" 

    # 3. Rutas a las plantillas
    PLANTILLA_INFORME1 = r"censos\FORMATO 1 IDENTIFICACIÓN - Aprobado.xlsx"
    PLANTILLA_FICHA    = r"censos\FICHA_PREDIAL_FINAL.xlsm"

    # 4. Instanciar el servicio
    servicio = GoogleSheetsAExcel(
        service_account_file=SERVICE_ACCOUNT_FILE,
        spreadsheet_id=SPREADSHEET_ID,
        drive_folder_id=DRIVE_FOLDER_ID,
        range_informe1=RANGE_INFORME1,
        range_ficha1=RANGE_FICHA1,
        range_ficha2=RANGE_FICHA2,
        plantilla_informe1=PLANTILLA_INFORME1,
        plantilla_ficha=PLANTILLA_FICHA
    )

    # 5. Inicializar servicios de Sheets y Drive
    servicio.inicializarServicios()

    # 6. Ejecutar procesos de llenado y subida para cada plantilla
    # servicio.llenarYSubirInforme1()
    servicio.llenarYSubirFichaPredial()
