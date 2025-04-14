from googleApi.api import GoogleSheetsAExcel

if __name__ == "__main__":

    # ==========================================================================
    # 1. CONFIGURACIÓN DE CREDENCIALES Y CARPETA DE DRIVE
    # ==========================================================================
    SERVICE_ACCOUNT_FILE = r'googleApi\censos-maute-d48ff1e9060b.json'
    SPREADSHEET_ID1      = '1WLZu1vYe8MtihM4kGvRq5Dj3-A5QgtmkSuF151drtZA'   # Primer spreadsheet (FP, UU, Act. Ec)
    SPREADSHEET_ID2      = '1TBqYQ3i4itD2OVswoAWjrOQd2Pu6TzVYqh1GgyjZmpU'   # Segundo spreadsheet (Agropecuario)
    SPREADSHEET_ID3      = '1nfYCVZLgWTdDJqiu6xGAjT9PEDu6hx-mMRdER3zFYZQ'   # Tercer spreadsheet (Comercial)
    SPREADSHEET_ID4      = '1A45AjJ8UFlebJNW9RlKVGF47l9xk-IjtSmBjfRmCyBk'   # Cuarto spreadsheet (Servicios) 
    SPREADSHEET_ID5      = '1jWdt-s6oTwgWFVn008sJ0Q1HCvc9y3ctetLZp5sNvk4'              # Quinto spreadsheet (Actividad Económica)             
    DRIVE_FOLDER_ID      = '1kq_6eo-_u0fuCUOHDRV_V5bQ1c_dvfMF'              # Carpeta de Drive para guardar los PDF

    # ==========================================================================
    # 2. RANGOS Y PLANTILLAS PARA LA PRIMERA INSTANCIA
    #    (Informe1, FichaPredial, UsosUsuarios)
    # ==========================================================================
    # Rangos
    RANGE_INFORME1       = "Sheet1!A1:EZ2000"
    RANGE_FICHA1         = "Sheet1!A1:EZ2000"
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

    RANGE_FORMATO_AGRO       = "Sheet1!A1:DD300"
    RANGE_INFO_COMERCIAL     = "data-informacion_comercial-insumos_actividad!A1:J200"
    RANGE_EXPLOT_AVICOLA     = "data-explotacion_avicola-tipo_explotacion!A1:J200"
    RANGE_INFO_LABORAL       = "data-informacion_laboral!A1:Q200"
    RANGE_EXPLOT_AGRICOLA    = "data-begin_agricola-cultivo!A1:L200"
    RANGE_EXPLOT_PORCINA     = "data-explotacion_porcina-categoria_porquina!A1:M200"
    RANGE_DETALLE_JORNAL     = "data-detalle_jornal-jornal_detalle!A1:H200"

    PLANTILLA_FORMATO_AGRO   = r"censos\Plantilla Agropecuario.xlsx"

    # ==========================================================================
    # 4. RANGOS Y PLANTILLA PARA LA TERCERA INSTANCIA: FORMATO COMERCIAL
    # ==========================================================================
    # -- Rangos para Formato Comercial --
    RANGE_FORMATO_COMERCIAL = "Sheet1!A1:AS300"
    RANGE_DESCP_ABAST       = "data-descripcion_actividad-abastecimiento!A1:F200"
    RANGE_DESC_ACTIVIDAD     = "data-descripcion_actividad-precio_venta!A1:Z200"
    RANGE_INFO_LABORAL_COMERCIAL = "data-informacion_laboral!A1:Q200"

    PLANTILLA_FORMATO_COMERCIAL = r"censos\Plantilla Comercial.xlsx"

    # ==========================================================================
    # 5. RANGOS Y PLANTILLAS PARA LA CUARTA INSTANCIA: FORMATO SERVICIOS
    # ==========================================================================
    # -- Rangos para Formato Servicios --
    RANGE_FORMATO_SERVICIOS = "Sheet1!A1:AQ300"
    RANGE_DESC_ACTIVIDAD_SERVICIOS = "data-desc_actividad-precio_servicios!A1:D2000"
    RANGE_INSUMOS_ABASTECIMIENTO = "data-begin_insumos-abastecimiento_insumos!A1:G200"
    RANGE_DESC_ACTIVIDAD_SERVICIOS_NUM = "data-desc_actividad-num_servicios!A1:C200"
    RANGE_EQUIPOS_MAQUINARIA = "data-equipos_maquinaria!A1:G200"
    RANGE_INFO_LABORAL_SERVICIOS = "data-informacion_laboral!A1:Q200"

    PLANTILLA_FORMATO_SERVICIOS = r"censos\Plantilla Servicios.xlsx"

    # ==========================================================================
    # 6. RANGOS Y PLANTILLA PARA LA QUINTA INSTANCIA: FORMATO ACTIVIDAD ECONOMICA
    # ==========================================================================
    # -- Rangos para Formato Actividad Económica --
    RANGE_ACTIVIDAD_ECONOMICA = "Sheet1!A1:AP300"

    PLANTILLA_ACTIVIDAD_ECONOMICA = r"censos\FORMATO 1 IDENTIFICACIÓN - Aprobado.xlsx"

    # ==========================================================================
    # 7. CREAR LA PRIMERA INSTANCIA: ENCUESTAS 1,2,3
    # ==========================================================================
    servicio1 = GoogleSheetsAExcel(
        service_account_file=SERVICE_ACCOUNT_FILE,
        spreadsheet_id=SPREADSHEET_ID1,   
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
    # 8. CREAR LA SEGUNDA INSTANCIA: FORMATO AGROPECUARIO
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
    servicio2.llenarYSubirFormatoAgropecuario()

    # ==========================================================================
    # 9. CREAR LA TERCERA INSTANCIA: FORMATO COMERCIAL
    # ==========================================================================

    servicio3 = GoogleSheetsAExcel(
        service_account_file=SERVICE_ACCOUNT_FILE,
        spreadsheet_id=SPREADSHEET_ID3,  
        drive_folder_id=DRIVE_FOLDER_ID,

        range_formato_comercial=RANGE_FORMATO_COMERCIAL,
        range_descripcion_abastecimiento=RANGE_DESCP_ABAST,
       
        range_descripcion_actividad_precio=RANGE_DESC_ACTIVIDAD,
        range_info_laboral2=RANGE_INFO_LABORAL_COMERCIAL,
        plantilla_formato_comercial=PLANTILLA_FORMATO_COMERCIAL
    )

    # Inicializa servicios y llama al método
    servicio3.inicializarServicios()
    servicio3.llenarYSubirFormatoComercial()

    # ==========================================================================
    # 10. CREAR LA CUARTA INSTANCIA: FORMATO SERVICIOS
    # ==========================================================================
    servicio4 = GoogleSheetsAExcel(
        service_account_file=SERVICE_ACCOUNT_FILE,
        spreadsheet_id=SPREADSHEET_ID4,  
        drive_folder_id=DRIVE_FOLDER_ID,

        range_formato_servicios=RANGE_FORMATO_SERVICIOS,
        range_desc_actividad_precio_servicios=RANGE_DESC_ACTIVIDAD_SERVICIOS,
        range_insumos_abastecimiento_servicios=RANGE_INSUMOS_ABASTECIMIENTO,
        range_desc_actividad_servicios=RANGE_DESC_ACTIVIDAD_SERVICIOS_NUM,
        range_equipos_maquinaria_servicios=RANGE_EQUIPOS_MAQUINARIA,
        range_info_laboral_servicios=RANGE_INFO_LABORAL_SERVICIOS,

        plantilla_formato_servicios=PLANTILLA_FORMATO_SERVICIOS
    )

    # Inicializa servicios y llama al método
    servicio4.inicializarServicios()
    servicio4.llenarYSubirFormatoServicios()

    # ==========================================================================
    # 11. CREAR LA QUINTA INSTANCIA: FORMATO ACTIVIDAD ECONOMICA
    # ==========================================================================
    servicio5 = GoogleSheetsAExcel(
        service_account_file=SERVICE_ACCOUNT_FILE,
        spreadsheet_id=SPREADSHEET_ID5,  
        drive_folder_id=DRIVE_FOLDER_ID,

        range_identificacion_actividad=RANGE_ACTIVIDAD_ECONOMICA,

        plantilla_identificacion_actividad=PLANTILLA_ACTIVIDAD_ECONOMICA
    )

    # Inicializa servicios y llama al método
    servicio5.inicializarServicios()
    servicio5.llenarYSubirIdentificacionActEconomica()