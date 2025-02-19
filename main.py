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
    RANGE_USOS1    = "Sheet1!A1:EZ200"
    RANGE_USOS2    = "data-start_bienes_serv-informacion_usos!A1:O200"


    # 3. Rutas a las plantillas
    PLANTILLA_INFORME1         = r"censos\FORMATO 1 IDENTIFICACION - Aprobado.xlsx"
    PLANTILLA_FICHA            = r"censos\FICHA_PREDIAL_FINAL.xlsm"
    PLANTILLA_USOS_USUARIOS    = r"censos\Formato_Usos_Usuarios_agua.xlsx"

    # 4. Instanciar el servicio
    servicio = GoogleSheetsAExcel(
        service_account_file=SERVICE_ACCOUNT_FILE,
        spreadsheet_id=SPREADSHEET_ID,
        drive_folder_id=DRIVE_FOLDER_ID,

        range_informe1=RANGE_INFORME1,
        range_ficha1=RANGE_FICHA1,
        range_ficha2=RANGE_FICHA2,
        range_usos1=RANGE_USOS1,
        range_usos2=RANGE_USOS2,

        plantilla_informe1=PLANTILLA_INFORME1,
        plantilla_ficha=PLANTILLA_FICHA,
        plantilla_usos_usuarios=PLANTILLA_USOS_USUARIOS
    )

    servicio.inicializarServicios()

    # servicio.llenarYSubirInforme1()
    servicio.llenarYSubirFichaPredial()
    # servicio.llenarYSubirUsosUsuarios()

