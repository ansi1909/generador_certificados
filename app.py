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
# Sustituye el marcador {{NOMBRE}} por el nombre real del participante
def generate_certificate(name, template_bytes):
    template_pptx = Presentation(template_bytes)  # Carga la plantilla PPTX desde los bytes
    output = BytesIO()  # Buffer para guardar el resultado
    prs = Presentation()  # Nueva presentación vacía donde se insertará el certificado

    # Itera sobre las diapositivas de la plantilla original
    for slide in template_pptx.slides:
        slide_copy = prs.slides.add_slide(prs.slide_layouts[5])  # Crea una nueva diapositiva vacía

        # Recorre las formas de la diapositiva original
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text
                # Reemplaza el marcador con el nombre real
                new_text = text.replace("{{NOMBRE}}", name)

                # Crea una nueva caja de texto en la diapositiva copiada
                new_shape = slide_copy.shapes.add_textbox(
                    shape.left, shape.top, shape.width, shape.height
                )
                new_frame = new_shape.text_frame
                new_frame.text = new_text

    prs.save(output)  # Guarda la presentación en el buffer
    output.seek(0)  # Vuelve al inicio del buffer para lectura
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

        # Crea un archivo ZIP en memoria
        with zipfile.ZipFile(zip_buffer, 'w') as zipf:
            for name in df[nombre_columna]:
                # Genera el certificado individual
                cert = generate_certificate(name, uploaded_template.read())

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
