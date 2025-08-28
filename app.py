from flask import Flask, request, render_template
import pandas as pd
import os
import platform
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from collections import defaultdict
import webbrowser
import threading
import logging
from pptx import Presentation  # NUEVO para manejar pptx

# Detectar sistema operativo
ON_WINDOWS = platform.system() == "Windows"

# Configurar logging
logging.basicConfig(
    level=logging.DEBUG,
    filename="app.log",
    filemode="w",
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Configuración para PyInstaller
def resource_path(relative_path):
    """Obtener ruta absoluta de recursos, funciona tanto en desarrollo como en .exe"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

app = Flask(__name__,
           template_folder=resource_path('templates'),
           static_folder=resource_path('static'))

def open_browser():
    try:
        webbrowser.open_new("http://127.0.0.1:5000/")
    except Exception as e:
        logger.error(f"Error abriendo navegador: {e}")

def get_downloads_folder():
    try:
        home = Path.home()
        downloads = home / "Downloads"
        if not downloads.exists():
            downloads = home / "Descargas"
            if not downloads.exists():
                downloads = home
        return downloads
    except Exception as e:
        logger.error(f"Error obteniendo carpeta de descargas: {e}")
        return Path.cwd()

def get_plantilla_path_pptx():
    possible = [
        resource_path("plantilla.pptx"),
        "plantilla.pptx",
        Path.cwd() / "plantilla.pptx",
        get_downloads_folder() / "certificados" / "plantilla.pptx"
    ]
    for p in possible:
        if os.path.exists(p):
            logger.info(f"Plantilla PPTX encontrada en: {p}")
            return str(p)
    raise FileNotFoundError("No se encontró plantilla.pptx en ninguna ubicación")

# ---------------- NUEVAS FUNCIONES PPTX ----------------

def safe_filename(s: str) -> str:
    keep = (" ", ".", "_", "-")
    return "".join(c for c in s if c.isalnum() or c in keep).rstrip()

def build_placeholder_map(context: dict) -> dict:
    mapping = {}
    for k, v in context.items():
        vstr = "" if v is None else str(v)
        mapping[f"{{{{{k.upper()}}}}}"] = vstr
        mapping[f"{{{{{k.lower()}}}}}"] = vstr
        mapping[k.upper()] = vstr
        mapping[k.lower()] = vstr
    return mapping

def replace_placeholders_in_presentation(prs: Presentation, mapping: dict):
    for slide in prs.slides:
        for shape in slide.shapes:
            if not hasattr(shape, "text"):
                continue
            try:
                tf = shape.text_frame
            except Exception:
                continue

            for paragraph in tf.paragraphs:
                for run in paragraph.runs:
                    orig = run.text or ""
                    new = orig
                    for ph, val in mapping.items():
                        if ph in new:
                            new = new.replace(ph, val)
                    if new != orig:
                        run.text = new

def render_pptx_template(template_path: str, context: dict, out_pptx_path: str):
    logger.info(f"Renderizando PPTX: {template_path} -> {out_pptx_path}")
    prs = Presentation(template_path)
    mapping = build_placeholder_map(context)
    replace_placeholders_in_presentation(prs, mapping)
    prs.save(out_pptx_path)

def convert_pptx_to_pdf(pptx_path: str, pdf_path: str, timeout: int = 30) -> bool:
    pptx_path = str(pptx_path)
    pdf_path = str(pdf_path)
    outdir = os.path.dirname(pdf_path)

    # --- LibreOffice ---
    try:
        logger.info("Intentando conversión con LibreOffice...")
        cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, pptx_path]
        subprocess.run(cmd, check=True, timeout=timeout,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        gen = os.path.join(outdir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
        if os.path.exists(gen):
            if os.path.abspath(gen) != os.path.abspath(pdf_path):
                os.replace(gen, pdf_path)
            logger.info(f"PDF generado: {pdf_path}")
            return True
    except Exception as e:
        logger.warning(f"LibreOffice no disponible: {e}")

    # --- PowerPoint COM (solo Windows) ---
    if ON_WINDOWS:
        try:
            import win32com.client
            logger.info("Intentando conversión con PowerPoint COM...")
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            ppt_app.Visible = 0
            presentation = ppt_app.Presentations.Open(pptx_path, WithWindow=False)
            presentation.SaveAs(pdf_path, 32)  # 32 = PDF
            presentation.Close()
            ppt_app.Quit()
            if os.path.exists(pdf_path):
                logger.info(f"PDF generado con PowerPoint COM: {pdf_path}")
                return True
        except Exception as e:
            logger.error(f"Error en PowerPoint COM: {e}")

    logger.error("No se pudo convertir PPTX a PDF.")
    return False

# ---------------- RUTAS ----------------

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/procesar', methods=['POST'])
def procesar():
    try:
        if 'excel_file' not in request.files:
            return "No se subió archivo Excel", 400

        excel_file = request.files['excel_file']
        if excel_file.filename == '':
            return "No se seleccionó archivo", 400

        df = pd.read_excel(excel_file)
        logger.info(f"Excel leído correctamente. Filas: {len(df)}")

        # columnas requeridas
        required = ["nombre", "cedula", "horas", "compañia", "certificado"]
        for col in required:
            if col not in [c.lower() for c in df.columns]:
                return f"Columna requerida no encontrada: {col}", 400
        df.columns = [c.lower() for c in df.columns]

        certificados_pendientes = df["certificado"].astype(str).str.lower().str.strip() == "no"
        if not certificados_pendientes.any():
            return render_template("error.html")

        plantilla_path = get_plantilla_path_pptx()
        downloads_folder = get_downloads_folder()
        output_dir = downloads_folder / "Certificados"
        os.makedirs(output_dir, exist_ok=True)

        certificados_por_compania = defaultdict(list)
        certificados_creados = 0

        for index, row in df.iterrows():
            try:
                if str(row["certificado"]).lower().strip() == "no":
                    contexto = {
                        "NOMBRE": str(row["nombre"]),
                        "CEDULA": str(row["cedula"]),
                        "HORAS": str(row["horas"])
                    }

                    compania_folder = output_dir / str(row["compañia"]).replace(" ", "_").replace("/", "_")
                    os.makedirs(compania_folder, exist_ok=True)

                    base_name = f"certificado_{row['plantilla']}_{row['nombre'].replace(' ', '_')}"
                    base_name = safe_filename(base_name)

                    pptx_file = compania_folder / f"{base_name}.pptx"
                    pdf_file = compania_folder / f"{base_name}.pdf"

                    render_pptx_template(plantilla_path, contexto, str(pptx_file))
                    converted = convert_pptx_to_pdf(str(pptx_file), str(pdf_file))
                    if converted:
                        os.remove(pptx_file)
                        logger.info(f"Certificado PDF creado: {pdf_file}")
                    else:
                        logger.error(f"No se pudo convertir a PDF, se deja PPTX: {pptx_file}")

                    certificados_por_compania[row["compañia"]].append(os.path.basename(str(pdf_file)))
                    certificados_creados += 1
                    df.at[index, "certificado"] = "si"

            except Exception as e:
                logger.error(f"Error procesando fila {index}: {e}")
                continue

        excel_actualizado = output_dir / "datos_actualizados.xlsx"
        df.to_excel(excel_actualizado, index=False)
        logger.info(f"Excel actualizado guardado: {excel_actualizado}")
        logger.info(f"Procesamiento completado. Certificados creados: {certificados_creados}")

        return render_template("success.html")

    except Exception as e:
        logger.error(f"Error general en procesamiento: {e}")
        return f"Error interno: {str(e)}", 500

@app.errorhandler(500)
def internal_error(error):
    return f"Error interno del servidor: {str(error)}", 500

if __name__ == "__main__":
    threading.Timer(2.0, open_browser).start()
    app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False)
