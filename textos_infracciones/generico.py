# textos_infracciones/generico.py
from docx.shared import Cm, Pt, RGBColor

def redactar_beneficio_ilicito(doc, datos_hecho, num_hecho):
    """
    Dibuja el texto por defecto si la infracción aún no tiene un script personalizado.
    """
    p_desc_bi = doc.add_paragraph(
        "El beneficio ilícito proviene del costo evitado del administrado por no cumplir con la normativa "
        "ambiental y/o sus obligaciones fiscalizables. En este caso, el administrado incumplió lo precisado "
        "en el párrafo anterior."
    )
    p_desc_bi.paragraph_format.left_indent = Cm(1.0)
    
    p_marcador_bi = doc.add_paragraph(f"[Aquí se programará la TABLA GENÉRICA para el Hecho {num_hecho}]")
    p_marcador_bi.paragraph_format.left_indent = Cm(1.0)
    p_marcador_bi.runs[0].font.color.rgb = RGBColor(255, 0, 0) # Rojo
