import io
import os  
import streamlit as st 
from copy import deepcopy

# Importaciones de docxcompose
from docxcompose.composer import Composer

# Importaciones de docxtpl
from docxtpl import DocxTemplate, Subdoc, RichText

# Importaciones de python-docx
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE

# --------------------------------------------------------------------
#  FUNCIONES DE WORD
# --------------------------------------------------------------------

def combinar_con_composer(path_plantilla_o_buffer, path_origen_o_buffer, path_final):
    """Combina dos documentos de Word usando docx-composer."""
    print(f"-> Combinando documentos con docx-composer...")
    try:
        plantilla = Document(path_plantilla_o_buffer)
        composer = Composer(plantilla)
        origen = Document(path_origen_o_buffer)
    except Exception as e:
        st.error(f"   ERROR abriendo documents para combinar: {e}")
        return False

    marcador = None
    marcador_idx = -1
    for i, p in enumerate(composer.doc.paragraphs):
        if '{{INSERTAR_CONTENIDO_AQUI}}' in p.text:
            marcador = p
            marcador_idx = i
            break
            
    if not marcador:
        st.error(f"   ERROR: No se encontró el marcador '{{INSERTAR_CONTENIDO_AQUI}}'.")
        return False
        
    composer.insert(marcador_idx, origen)
    for p in composer.doc.paragraphs:
        if '{{INSERTAR_CONTENIDO_AQUI}}' in p.text:
            p._element.getparent().remove(p._element)
            break
            
    composer.save(path_final)
    
    # Comprueba si 'path_final' es una ruta de archivo (string) o un buffer
    if isinstance(path_final, str):
        nombre_amigable = os.path.basename(path_final)
        print(f"-> Fusión guardada en '{nombre_amigable}'.")
    else:
        print(f"-> Fusión guardada en un buffer en memoria.")

    return True

# --- Funciones Auxiliares de Formato ---

def set_cell_border(cell, **kwargs):
    """
    Función auxiliar para definir los bordes de una celda.
    Uso: set_cell_border(cell, top={"sz": 12, "val": "single", "color": "#000000"})
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            border = tcBorders.find(qn(tag))
            if border is None:
                border = OxmlElement(tag)
                tcBorders.append(border)
            for k, v in edge_data.items():
                border.set(qn('w:{}'.format(k)), str(v))


def set_cell_shading(cell, fill_color):
    """
    Función auxiliar para definir el color de fondo de una celda.
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tcPr.append(shd)

# --- Funciones Constructoras de Contenido ---

def create_table_subdoc(doc_template, headers, data, keys):
    """
    Crea una tabla con formato avanzado (Arial 8, negritas, bordes y centrado total)
    en un subdocumento de Word.
    """
    sub = doc_template.new_subdoc()
    table = sub.add_table(rows=1, cols=len(headers))
    table.style = 'TablaAnexo'

    # --- Formato del Encabezado ---
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(headers):
        p = hdr_cells[i].paragraphs[0]
        run = p.add_run(header_text)
        run.font.name = 'Arial'
        run.font.size = Pt(8)
        run.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_border(hdr_cells[i], top={"sz": 4, "val": "single", "color": "000000"},
                        bottom={"sz": 4, "val": "single", "color": "000000"})

    # --- Formato de las Filas de Datos ---
    for item in data:
        row_cells = table.add_row().cells
        for i, key in enumerate(keys):
            cell_text = str(item.get(key, ''))
            p = row_cells[i].paragraphs[0]
            run = p.add_run(cell_text)
            run.font.name = 'Arial'
            run.font.size = Pt(8)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            row_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_border(row_cells[i], top={"val": "nil"}, bottom={"val": "nil"})

    # --- Formato de la Última Fila (Total) ---
    if data:
        last_row_cells = table.rows[-1].cells
        for cell in last_row_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True
            set_cell_border(cell, top={"sz": 4, "val": "single", "color": "000000"},
                            bottom={"sz": 4, "val": "single", "color": "000000"})

    return sub

def create_main_table_subdoc(doc_template, headers, data, keys):
    """
    Crea una tabla con formato y ahora es lo suficientemente inteligente
    para manejar tanto tablas simples como tablas con superíndices.
    """
    sub = doc_template.new_subdoc()
    table = sub.add_table(rows=1, cols=len(headers))
    table.style = 'TablaCuerpo'
    SHADING_COLOR = "D9D9D9"

    # --- Formato del Encabezado (sin cambios) ---
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(headers):
        # ... (tu código de formato de encabezado se mantiene igual) ...
        p = hdr_cells[i].paragraphs[0]
        run = p.add_run(header_text)
        run.font.name = 'Arial'
        run.font.size = Pt(10)
        run.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_shading(hdr_cells[i], SHADING_COLOR)


    # --- Formato de las Filas de Datos (CON LÓGICA MEJORADA) ---
    for row_idx, item in enumerate(data):
        row_cells = table.add_row().cells
        is_last_row = (row_idx == len(data) - 1)

        for col_idx, key in enumerate(keys):
            p = row_cells[col_idx].paragraphs[0]
            p.clear()

            # --- INICIO DE LA LÓGICA MEJORADA ---
            
            # Comprobamos si es una fila especial con superíndice (solo para la tabla de BI)
            if 'descripcion_superindice' in item and col_idx == 0:
                # Si lo es, aplicamos la lógica de dos partes
                texto_principal = item.get('descripcion_texto', '')
                run_texto = p.add_run(texto_principal)
                run_texto.font.name = 'Arial'
                run_texto.font.size = Pt(10)
                if is_last_row:
                    run_texto.bold = True
                
                superindice = item.get('descripcion_superindice', '')
                if superindice:
                    run_super = p.add_run(superindice)
                    run_super.font.superscript = True
                    run_super.font.name = 'Arial'
                    run_super.font.size = Pt(10)
                    if is_last_row:
                        run_super.bold = True
            
            # Para todas las demás tablas (como la de Multa) y las demás columnas
            else:
                cell_text = str(item.get(key, ''))
                run = p.add_run(cell_text)
                run.font.name = 'Arial'
                run.font.size = Pt(10)
                if is_last_row:
                    run.bold = True

            # --- FIN DE LA LÓGICA MEJORADA ---

            # El resto del formato de celda se aplica a todos los casos
            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            row_cells[col_idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            if is_last_row:
                set_cell_shading(row_cells[col_idx], SHADING_COLOR)
                
    return sub

def create_summary_table_subdoc(doc_template, headers, data, keys):
    """
    Crea la tabla de resumen final de la multa, con celdas combinadas.
    """
    sub = doc_template.new_subdoc()
    table = sub.add_table(rows=1, cols=len(headers))
    table.style = 'TablaCuerpo'
    SHADING_COLOR = "D9D9D9"

    # --- Formato del Encabezado ---
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(headers):
        p = hdr_cells[i].paragraphs[0];
        run = p.add_run(header_text)
        run.font.name = 'Arial';
        run.font.size = Pt(10);
        run.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_shading(hdr_cells[i], SHADING_COLOR)

    # --- Fila de Datos Principal ---
    for item in data[:-1]:
        row_cells = table.add_row().cells
        for col_idx, key in enumerate(keys):
            cell_text = str(item.get(key, ''))
            p = row_cells[col_idx].paragraphs[0];
            run = p.add_run(cell_text)
            run.font.name = 'Arial';
            run.font.size = Pt(10)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            row_cells[col_idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # --- Fila de Total con Celdas Combinadas ---
    total_multa_str = data[-1]['multa'] if data else "0.000 UIT"  # Obtenemos el valor de la multa de los datos
    total_cells = table.add_row().cells

    # Combinar las primeras dos celdas
    merged_cell = total_cells[0].merge(total_cells[1])

    # Añadir y formatear el texto "Total"
    p_total = merged_cell.paragraphs[0];
    run_total = p_total.add_run('Total')
    run_total.font.name = 'Arial';
    run_total.font.size = Pt(10);
    run_total.bold = True
    p_total.alignment = WD_ALIGN_PARAGRAPH.CENTER
    merged_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_shading(merged_cell, SHADING_COLOR)

    # Añadir y formatear el monto total en la tercera celda
    p_multa = total_cells[2].paragraphs[0];
    run_multa = p_multa.add_run(total_multa_str)
    run_multa.font.name = 'Arial';
    run_multa.font.size = Pt(10);
    run_multa.bold = True
    p_multa.alignment = WD_ALIGN_PARAGRAPH.CENTER
    total_cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_shading(total_cells[2], SHADING_COLOR)

    # Asegurar que la celda "fantasma" (la segunda) también tenga el fondo gris
    set_cell_shading(total_cells[1], SHADING_COLOR)

    return sub
