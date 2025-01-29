from reportlab.lib.pagesizes import LETTER
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image, PageBreak
from reportlab.lib.units import cm
from reportlab.lib.colors import Color
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
import datetime
import base64
import io

def create_ficha_predial():
    # Configuración del documento
    doc = SimpleDocTemplate("ficha_predial.pdf", pagesize=LETTER)
    elements = []
    
    # Estilos personalizados
    styles = {
        "title": ParagraphStyle(
            name="Title",
            fontSize=14,
            leading=16,
            alignment=1,
            fontName="Helvetica-Bold"
        ),
        "header": ParagraphStyle(
            name="Header",
            fontSize=12,
            leading=14,
            fontName="Helvetica-Bold"
        ),
        "normal": ParagraphStyle(
            name="Normal",
            fontSize=10,
            leading=12,
            fontName="Helvetica"
        )
    }

    # Contenido del documento
        logo = Image("Imagenes/Logo.png", width=4.5*cm, height=1.2*cm)
        
        encabezado_data = [
            [logo, Paragraph("<b>Orden de compra</b>", self.title_style), Paragraph("<b>Fecha</b>", self.small_bold), self.fechaActual],
            ["", "", Paragraph("<b>Página</b>", self.small_bold), "1/1"],
            ["", "", Paragraph("<b>Versión</b>", self.small_bold), "Versión 10"],
            ["", "", Paragraph("<b>Código</b>", self.small_bold), "FO-GR-009"]

    # Tabla de servicios geológicos
    servicio_data = [
        ["FECHA PROYECTO", "DEPARTAMENTO", "VEREDA", "CUENCA"],
        ["MUNICIPIO", "TENENCIA:", "Propia □", "Arriendo □", "Trabajo □", "Familiar □"]
    ]
    
    servicio_table = Table(servicio_data, colWidths=[100, 100, 80, 80, 80, 80])
    servicio_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('SPAN', (0,1), (1,1))
    ]))
    elements.append(servicio_table)
    elements.append(Spacer(1, 12))
    
    # Sección de datos del predio
    predio_data = [
        ["NOMBRE DEL PREDIO", "TELEFONO", "Vive en el predio:", "Sí □", "No □"]
    ]
    predio_table = Table(predio_data, colWidths=[180, 100, 100, 50, 50])
    predio_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey)
    ]))
    elements.append(predio_table)
    elements.append(Spacer(1, 12))
    
    # Sección de escrituración
    elements.append(Paragraph("ESCRITURADA:", styles["header"]))
    escrituracion_data = [["SÍ □", "No □", "No sabe □", "No. de registro de la escritura: ________________"]]
    escrituracion_table = Table(escrituracion_data, colWidths=[60, 60, 80, 200])
    elements.append(escrituracion_table)
    elements.append(Spacer(1, 12))
    
    # Sección de encargado
    elements.append(Paragraph("NOMBRE DEL ENCARGADO O ADMINISTRADOR", styles["header"]))
    elements.append(Paragraph("TELEFONO: _________________________", styles["normal"]))
    
    # Generar PDF
    doc.build(elements)


if __name__ == "__main__":

    create_ficha_predial()