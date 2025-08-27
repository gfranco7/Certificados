from flask import Flask, request, render_template
from docxtpl import DocxTemplate
import pandas as pd
import os
import platform
import subprocess
from datetime import datetime
from pathlib import Path
from collections import defaultdict
import pythoncom
import webbrowser
import threading
# Detectar sistema operativo
ON_WINDOWS = platform.system() == "Windows"
if ON_WINDOWS:
    from docx2pdf import convert  # solo en Windows

app = Flask(__name__)

def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000/")

def get_downloads_folder():
    home = Path.home()
    downloads = home / "Downloads"
    if not downloads.exists():
        downloads = home
    return downloads

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/procesar', methods=['POST'])
def procesar():
    pythoncom.CoInitialize()
    if 'excel_file' not in request.files:
        return "No se subió archivo Excel"

    excel_file = request.files['excel_file']
    df = pd.read_excel(excel_file)

    #Se valida que haya al menos un 'no' en la columna 'certificado' para generar certificados
    if not (df["certificado"].astype(str).str.lower() == "no").any():
        return render_template("error.html")
    
    plantilla = DocxTemplate("plantilla.docx")

    # Carpeta principal en Descargas
    fecha_hoy = datetime.today().strftime("%Y-%m-%d")
    downloads_folder = get_downloads_folder()
    output_dir = downloads_folder / f"certificados"
    os.makedirs(output_dir, exist_ok=True)

    certificados_por_compania = defaultdict(list)
    


    for index, row in df.iterrows():
        contexto = {
            "NOMBRE": row["nombre"],
            "CEDULA": row["cedula"],
            "FECHA": row["fecha"].strftime("%d/%m/%Y") if not pd.isna(row["fecha"]) else "",
            "COMPANIA": row["compañia"]
        }

        if row["certificado"] == "no":
            plantilla.render(contexto)

            # Crear subcarpeta por compañía
            compania_folder = output_dir / str(row["compañia"]).replace(" ", "_")
            os.makedirs(compania_folder, exist_ok=True)

            nombre_base = f"certificado_{row['nombre'].replace(' ', '_')}_{row['plantilla']}"
            docx_file = compania_folder / f"{nombre_base}.docx"
            pdf_file = compania_folder / f"{nombre_base}.pdf"

            plantilla.save(docx_file)
            df.at[index, "certificado"] = "si"
            # Comversión a PDF
            if ON_WINDOWS:
                # Word (docx2pdf)
                convert(str(docx_file), str(pdf_file))
                os.remove(docx_file)
            else:
                # LibreOffice (Linux/Mac)
                subprocess.run([
                    "soffice", "--headless", "--convert-to", "pdf", "--outdir", str(compania_folder), str(docx_file)
                ])
                os.remove(docx_file)
    
            certificados_por_compania[row["compañia"]].append(f"{nombre_base}.pdf")
        
        else:
            continue

    excel_actualizado = output_dir / "datos_persona.xlsx"
    df.to_excel(excel_actualizado, index=False)


    pythoncom.CoUninitialize()
    try:
        if ON_WINDOWS:
            os.startfile(output_dir)
        elif platform.system() == "Darwin":  # Mac
            subprocess.run(["open", output_dir])
        else:  # Linux
            subprocess.run(["xdg-open", output_dir])
    except Exception as e:
        print("No se pudo abrir automáticamente la carpeta:", e)

    return render_template("success.html", ruta=output_dir, certificados=certificados_por_compania)


def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000/")


if __name__ == "__main__":
    threading.Timer(1.5, open_browser).start()
    app.run(debug=False)