import streamlit as st

# 1. Configuración de página DEBE ir aquí, antes de la navegación
st.set_page_config(page_title="Asistente de Multas OEFA", layout="wide", page_icon="⚖️")

# 2. Definimos las páginas apuntando a tus scripts actuales
pagina_v2 = st.Page("nueva_app.py", title="Versión 2 (Generación Python)", icon="🚀", default=True)
pagina_v1 = st.Page("app.py", title="Versión 1 (Plantillas Drive)", icon="📁")

# 3. Creamos la navegación (menú lateral)
pg = st.navigation(
    {
        "Selector de Versiones": [pagina_v2, pagina_v1]
    }
)

# 4. Ejecutamos la página seleccionada
pg.run()
