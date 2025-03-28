import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from io import BytesIO
import zipfile
import os
import subprocess
import tempfile

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

# Función para convertir PPT a PDF usando LbreOffice
def convert_to_pdf(pptx_bytes, output_filename):
    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "temp_cert.pptx")
        pdf_dir = tmpdir

        with open(pptx_path, "wb") as f:
            f.write(pptx_bytes.read())

        subprocess.run([
            "soffice", "--headless", "--convert-to", "pdf", pptx_path, "--outdir", pdf_dir
        ], check=True )

        pdf_path = pptx_path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f:
            return f.read()
        
# Carga de archivos: plantilla PowerPoint y archivo Excel
uploaded_template = st.file_uploader(
    "Sube la plantilla PowerPoint (.pptx) con {{NOMBRE}}", type="pptx"
)
uploaded_excel = st.file_uploader(
    "Sube el archivo Excel con los participantes", type="xlsx"
)

# Selección de formato de salida
output_format = st.selectbox("Selecciona el formato de salida", ['PPTX', 'PDF'])

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
                filename = f"Certificado_{name.replace(' ', '_')}"

                if output_format == "PDF":
                    pdf_content = convert_to_pdf(cert, filename_base + ".pdf")
                    zipf.writestr(filename_base + ".pdf", pdf_content)
                else:
                    zipf.writestr(filename_base + ".pptx", cert.read())
            
        zip_buffer.seek(0)  # Prepara el buffer para descarga

        st.success("✨ Certificados generados correctamente")

        # Botón para descargar el archivo ZIP con todos los certificados
        st.download_button(
            label="📦 Descargar ZIP de certificados",
            data=zip_buffer,
            file_name="certificados.zip",
            mime="application/zip"
        )
