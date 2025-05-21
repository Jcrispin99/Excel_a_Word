import os
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def set_table_borders(table):
    # Tu código existente para configurar bordes
    tbl = table._element
    tblPr = tbl.xpath("w:tblPr")
    if not tblPr:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    else:
        tblPr = tblPr[0]
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '8')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)
    tblPr.append(tblBorders)

def remove_paragraph_spacing(paragraph):
    # Tu código existente para eliminar espacios
    if paragraph:
        p = paragraph._p
        pPr = p.get_or_add_pPr()
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:before'), '20')
        spacing.set(qn('w:after'), '20')
        pPr.append(spacing)

def adjust_cell_spacing(cell):
    # Tu código existente para ajustar espacios en celdas
    tcPr = cell._element.tcPr
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        cell._element.append(tcPr)
    
    tblCellMar = OxmlElement('w:tcMar')
    margins = {'top': '40', 'left': '60', 'bottom': '40', 'right': '60'}
    for margin, value in margins.items():
        mar = OxmlElement(f'w:{margin}')
        mar.set(qn('w:w'), value)
        mar.set(qn('w:type'), 'dxa')
        tblCellMar.append(mar)
    
    existing_mar = tcPr.find(qn('w:tcMar'))
    if existing_mar is not None:
        tcPr.remove(existing_mar)
    tcPr.append(tblCellMar)
    
    tcVAlign = OxmlElement('w:vAlign')
    tcVAlign.set(qn('w:val'), 'top')
    tcPr.append(tcVAlign)
    
    for paragraph in cell.paragraphs:
        remove_paragraph_spacing(paragraph)

def generate_word_document(excel_path, images_dir, output_path):
    # Leer el archivo Excel
    df = pd.read_excel(excel_path)
    doc = Document()
    
    # Configurar márgenes del documento
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.3)
        section.bottom_margin = Inches(0.3)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
    
    # Procesar cada fila del Excel
    for index, row in df.iterrows():
        # Buscar la imagen correspondiente en el directorio de imágenes
        imagen_path = None
        if 'IMAGEN' in row and pd.notna(row['IMAGEN']):
            imagen_nombre = str(row['IMAGEN'])
            for root, dirs, files in os.walk(images_dir):
                for file in files:
                    if file.lower() == imagen_nombre.lower() or file.lower() == os.path.basename(imagen_nombre).lower():
                        imagen_path = os.path.join(root, file)
                        break
        
        # Obtener el número de cajas
        num_cajas = int(row.get('CONTEO_CAJAS', 1))
        
        # Generar una tabla para cada caja
        for num_caja in range(1, num_cajas + 1):
            # Crear tabla principal
            table = doc.add_table(rows=2, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = False
            table.columns[0].width = Inches(3)
            table.columns[1].width = Inches(3)
            set_table_borders(table)
            
            # Ajustar espacios en todas las celdas
            for row_cells in table.rows:
                for cell in row_cells.cells:
                    adjust_cell_spacing(cell)
            
            # Celda izquierda: imagen y datos
            left_cell = table.cell(0, 0)
            
            # Agregar imagen si existe
            if imagen_path and os.path.exists(imagen_path):
                try:
                    p = left_cell.paragraphs[0]
                    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    run = p.add_run()
                    run.add_picture(imagen_path, width=Inches(2.5))
                except Exception:
                    p = left_cell.paragraphs[0]
                    p.text = "Imagen no encontrada"
            
            # Tabla de datos
            data_table = left_cell.add_table(rows=6, cols=1)
            data_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(data_table)
            
            for row_cells in data_table.rows:
                for cell in row_cells.cells:
                    adjust_cell_spacing(cell)
            
            data_table.cell(0, 0).text = f"CÓDIGO: {row['CODIGO']}"
            data_table.cell(1, 0).text = "REVISADO POR:"
            data_table.cell(2, 0).text = f"DESCRIPCIÓN: {row['DESCRIPCION']}"
            data_table.cell(3, 0).text = f"CAJA {num_caja} DE {num_cajas}"
            data_table.cell(4, 0).text = "RECEPCIONADO:"
            data_table.cell(5, 0).text = f"CANTIDAD: {row['CANTIDAD']}"
            
            # Celda derecha: tabla de registro
            right_cell = table.cell(0, 1)
            
            # Eliminar párrafo existente
            if right_cell.paragraphs:
                p = right_cell.paragraphs[0]
                p._element.getparent().remove(p._element)
                p._p = None
            
            # Tabla interna
            inner_table = right_cell.add_table(rows=22, cols=4)
            inner_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(inner_table)
            
            for row_cells in inner_table.rows:
                for cell in row_cells.cells:
                    adjust_cell_spacing(cell)
            
            headers = ['FECHA DD/MM/AA', 'CANT.', 'RP.', 'FIRMA']
            for j, header in enumerate(headers):
                inner_table.cell(0, j).text = header
            
            for i in range(1, 22):
                for j in range(4):
                    inner_table.cell(i, j).text = ""
            
            # Celda de observación
            obs_cell = table.cell(1, 0)
            obs_cell.merge(table.cell(1, 1))
            obs_cell.text = "OBSERVACIÓN:"
            
            # Aumentar altura de la celda de observación
            tr = obs_cell._element.getparent()
            trHeight = OxmlElement('w:trHeight')
            trHeight.set(qn('w:val'), '800')
            trHeight.set(qn('w:hRule'), 'atLeast')
            trPr = tr.get_or_add_trPr()
            trPr.append(trHeight)
            
            # Agregar salto de página excepto en la última tabla
            if not (index == len(df) - 1 and num_caja == num_cajas):
                doc.add_page_break()
    
    # Guardar el documento
    doc.save(output_path)
    return output_path