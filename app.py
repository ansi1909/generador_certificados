import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from io import BytesIO
import zipfile

# Configuración inicial de la app de Streamlit
title = "🎓 Generador de Certificados"
st.set_page_config(page_title=title, layout="centered")
st.title(title)

# Función que genera un certificado a partir de un nombre y una plantilla
# Sustituye el marcador {{NOMBRE}} directamente en la diapositiva original

def generate_certificate(name, template_bytes):
    prs = Presentation(BytesIO(template_bytes))  # Carga la presentación desde los bytes
    output = BytesIO()

    # Recorre cada diapositiva y reemplaza el texto que contenga {{NOMBRE}}
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if "{{NOMBRE}}" in run.text:
                            run.text = run.text.replace("{{NOMBRE}}", name)

    prs.save(output)
    output.seek(0)
    return output

# Carga de archivos: plantilla PowerPoint y archivo Excel
uploaded_template = st.file_uploader(
    "Sube la plantilla PowerPoint (.pptx) con {{NOMBRE}}", type="pptx"
)
uploaded_excel = st.file_uploader(
    "Sube el archivo Excel con los participantes", type="xlsx"
)

# Verifica que ambos archivos hayan sido subidos
if uploaded_template and uploaded_excel:
    df = pd.read_excel(uploaded_excel)  # Carga el Excel como DataFrame

    # Permite al usuario seleccionar la columna que contiene los nombres
    nombre_columna = st.selectbox("Selecciona la columna con los nombres", df.columns)

    # Botón para generar certificados
    if st.button("✅ Generar certificados"):
        zip_buffer = BytesIO()  # Buffer para crear el archivo ZIP final

        # Leemos el contenido de la plantilla una sola vez y lo reutilizamos
        template_bytes = uploaded_template.read()

        # Crea un archivo ZIP en memoria
        with zipfile.ZipFile(zip_buffer, 'w') as zipf:
            for name in df[nombre_columna]:
                # Genera el certificado individual
                cert = generate_certificate(name, template_bytes)

                # Asigna un nombre de archivo amigable
                filename = f"Certificado_{name.replace(' ', '_')}.pptx"

                # Añade el certificado al ZIP
                zipf.writestr(filename, cert.read())

        zip_buffer.seek(0)  # Prepara el buffer para descarga

        st.success("✨ Certificados generados correctamente")

        # Botón para descargar el archivo ZIP con todos los certificados
        st.download_button(
            label="📦 Descargar ZIP de certificados",
            data=zip_buffer,
            file_name="certificados.zip",
            mime="application/zip"
        )
