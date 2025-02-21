from googleApi.api import GoogleSheetsAExcel

if __name__ == "__main__":

    # ==========================================================================
    # 1. CONFIGURACIÓN DE CREDENCIALES Y CARPETA DE DRIVE
    # ==========================================================================
    SERVICE_ACCOUNT_FILE = r'googleApi\censos-maute-d48ff1e9060b.json'
    SPREADSHEET_ID1      = '1WLZu1vYe8MtihM4kGvRq5Dj3-A5QgtmkSuF151drtZA'   # Primer spreadsheet (FP, UU, Act. Ec)
    SPREADSHEET_ID2      = '1TBqYQ3i4itD2OVswoAWjrOQd2Pu6TzVYqh1GgyjZmpU'   # Segundo spreadsheet (Agropecuario)
    DRIVE_FOLDER_ID      = '1kq_6eo-_u0fuCUOHDRV_V5bQ1c_dvfMF'

    # ==========================================================================
    # 2. RANGOS Y PLANTILLAS PARA LA PRIMERA INSTANCIA
    #    (Informe1, FichaPredial, UsosUsuarios)
    # ==========================================================================
    # Ejemplo de rangos
    RANGE_INFORME1       = "Sheet1!A1:EZ200"
    RANGE_FICHA1         = "Sheet1!A1:EZ200"
    RANGE_FICHA2         = "data-start_carac_poblacion-caracteristicas_poblacion!A1:H200"
    RANGE_USOS1          = "Sheet1!A1:EZ200"
    RANGE_USOS2          = "data-start_bienes_serv-informacion_usos!A1:O200"

    # Plantillas
    PLANTILLA_INFORME1   = r"censos\FORMATO 1 IDENTIFICACIÓN - Aprobado.xlsx"
    PLANTILLA_FICHA      = r"censos\FICHA_PREDIAL_FINAL.xlsm"
    PLANTILLA_USOS       = r"censos\Formato_Usos_Usuarios_agua.xlsx"

    # ==========================================================================
    # 3. RANGOS Y PLANTILLA PARA LA SEGUNDA INSTANCIA (FORMATO AGROPECUARIO)
    # ==========================================================================
    # -- Rangos para Formato Agropecuario --

    RANGE_FORMATO_AGRO       = "Sheet1!A1:DC300"
    RANGE_INFO_COMERCIAL     = "data-informacion_comercial-insumos_actividad!A1:J200"
    RANGE_EXPLOT_AVICOLA     = "data-explotacion_avicola-tipo_explotacion!A1:J200"
    RANGE_INFO_LABORAL       = "data-informacion_laboral!A1:Q200"
    RANGE_EXPLOT_AGRICOLA    = "data-begin_agricola-cultivo!A1:L200"
    RANGE_EXPLOT_PORCINA     = "data-explotacion_porcina-categoria_porquina!A1:M200"
    RANGE_DETALLE_JORNAL     = "data-detalle_jornal-jornal_detalle!A1:H200"

    PLANTILLA_FORMATO_AGRO   = r"censos\Plantilla Agropecuario.xlsx"

    # ==========================================================================
    # 4. CREAR LA PRIMERA INSTANCIA: ENCUESTAS 1,2,3
    # ==========================================================================
    servicio1 = GoogleSheetsAExcel(
        service_account_file=SERVICE_ACCOUNT_FILE,
        spreadsheet_id=SPREADSHEET_ID1,   # ← Usa la ID del primer spreadsheet
        drive_folder_id=DRIVE_FOLDER_ID,

        range_informe1=RANGE_INFORME1,
        range_ficha1=RANGE_FICHA1,
        range_ficha2=RANGE_FICHA2,
        range_usos1=RANGE_USOS1,
        range_usos2=RANGE_USOS2,

        plantilla_informe1=PLANTILLA_INFORME1,
        plantilla_ficha=PLANTILLA_FICHA,
        plantilla_usos_usuarios=PLANTILLA_USOS
    )

    # Inicializamos servicios y ejecutamos
    servicio1.inicializarServicios()

    servicio1.llenarYSubirInforme1()
    servicio1.llenarYSubirFichaPredial()
    servicio1.llenarYSubirUsosUsuarios()

    # ==========================================================================
    # 5. CREAR LA SEGUNDA INSTANCIA: FORMATO AGROPECUARIO
    # ==========================================================================
    servicio2 = GoogleSheetsAExcel(
        service_account_file=SERVICE_ACCOUNT_FILE,
        spreadsheet_id=SPREADSHEET_ID2,  
        drive_folder_id=DRIVE_FOLDER_ID,

        range_formato_agro=RANGE_FORMATO_AGRO,
        range_info_comercial=RANGE_INFO_COMERCIAL,
        range_explot_avicola=RANGE_EXPLOT_AVICOLA,
        range_info_laboral=RANGE_INFO_LABORAL,
        range_explot_agricola=RANGE_EXPLOT_AGRICOLA,
        range_explot_porcina=RANGE_EXPLOT_PORCINA,
        range_detalle_jornal=RANGE_DETALLE_JORNAL,

        plantilla_formato_agro=PLANTILLA_FORMATO_AGRO
    )

    # Inicializa servicios y llama al método
    servicio2.inicializarServicios()

    servicio2.llenarYSubirFormatoAgropecuario()
