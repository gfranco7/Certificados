from flask import Flask, request, render_template
from docxtpl import DocxTemplate
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

log_path = Path.cwd() / "app.log"
sys.stdout = open(log_path, "w", buffering=1)
sys.stderr = open(log_path, "w", buffering=1)

import logging
log = logging.getLogger('werkzeug')
log.disabled = True


# Detectar sistema operativo
ON_WINDOWS = platform.system() == "Windows"
if ON_WINDOWS:
    try:
        from docx2pdf import convert  # solo en Windows
        import pythoncom
    except ImportError as e:
        logging.error(f"Error importando docx2pdf o pythoncom: {e}")

# Configurar logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Configuración para PyInstaller
def resource_path(relative_path):
    """Obtener ruta absoluta de recursos, funciona tanto en desarrollo como en .exe"""
    try:
        # PyInstaller crea una carpeta temporal y almacena la ruta en _MEIPASS
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
            downloads = home / "Descargas"  # Para sistemas en español
            if not downloads.exists():
                downloads = home
        return downloads
    except Exception as e:
        logger.error(f"Error obteniendo carpeta de descargas: {e}")
        return Path.cwd()

def get_plantilla_path():
    """Buscar plantilla.docx en diferentes ubicaciones"""
    possible_paths = [
        resource_path("plantilla_final.docx"),  # Empaquetada con exe
        "plantilla.docx",  # Directorio actual
        Path.cwd() / "plantilla.docx",  # Directorio de trabajo
        get_downloads_folder() / "certificados" / "plantilla.docx"  # En carpeta certificados
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            logger.info(f"Plantilla encontrada en: {path}")
            return str(path)
    
    raise FileNotFoundError("No se encontró plantilla.docx en ninguna ubicación")

@app.route('/')
def index():
    try:
        return render_template('index.html')
    except Exception as e:
        logger.error(f"Error en ruta index: {e}")
        return f"Error cargando página: {str(e)}", 500

@app.route('/procesar', methods=['POST'])
def procesar():
    try:
        # Inicializar COM solo en Windows
        if ON_WINDOWS:
            try:
                pythoncom.CoInitialize()
            except:
                pass  # Ya está inicializado
        
        logger.info("Iniciando procesamiento...")
        
        if 'excel_file' not in request.files:
            logger.error("No se subió archivo Excel")
            return "No se subió archivo Excel", 400

        excel_file = request.files['excel_file']
        if excel_file.filename == '':
            return "No se seleccionó archivo", 400
    
        logger.info(f"Archivo Excel recibido: {excel_file.name}")
        # Leer Excel
        try:
            df = pd.read_excel(excel_file)
            logger.info(f"Excel leído correctamente. Filas: {len(df)}")
            excel_filename = excel_file.filename
            logger.info(f"Archivo Excel recibido: {excel_filename}")
        except Exception as e:
            logger.error(f"Error leyendo Excel: {e}")
            return f"Error leyendo archivo Excel: {str(e)}", 400

        required_mapping = {
            'item': ['item', 'Item', 'ITEM'],
            'nombre': ['nombre', 'Nombre', 'NOMBRE'],
            'cedula': ['cedula', 'Cedula', 'CÉDULA', 'cédula', 'Cédula', 'CEDULA'],
            'fecha': ['fecha', 'Fecha', 'FECHA'],
            'compañia': ['compañia', 'Compañia', 'COMPAÑÍA', 'compania', 'Compania', 'empresa', 'Empresa', 'COMPAÑIA'],
            'certificado': ['certificado', 'Certificado', 'CERTIFICADO'],
            'horas': ['horas', 'Horas', 'HORAS'],
            'id_formacion': ['id_formacion', 'Id_Formacion', 'ID_FORMACION', 'id formación', 'Id Formación', 'ID FORMACIÓN']
        }
        
        # Normalizar nombres de columnas
        column_mapping = {}
        for standard_name, variations in required_mapping.items():
            found = False
            for variation in variations:
                if variation in df.columns:
                    column_mapping[variation] = standard_name
                    found = True
                    break
            if not found:
                return f"No se encontró la columna '{standard_name}' en el Excel. Columnas disponibles: {list(df.columns)}", 400
        
        # Renombrar columnas para estandarizar
        df = df.rename(columns=column_mapping)

        # Se valida que haya al menos un 'no' en la columna 'certificado'
        certificados_pendientes = df["certificado"].astype(str).str.lower().str.strip() == "no"
        if not certificados_pendientes.any():
            return render_template("error.html")
        
        # Obtener plantilla
        try:
            plantilla_path = get_plantilla_path()
        except FileNotFoundError as e:
            logger.error(str(e))
            return f"Error: {str(e)}", 400

        # Crear carpetas de salida
        downloads_folder = get_downloads_folder()
        output_dir = downloads_folder / "Certificados"
        os.makedirs(output_dir, exist_ok=True)
        logger.info(f"Carpeta de salida: {output_dir}")

        certificados_por_compania = defaultdict(list)
        certificados_creados = 0

        # Procesar cada fila
        for index, row in df.iterrows():
            try:
                if str(row["certificado"]).lower().strip() == "no":
                    logger.info(f"Procesando certificado para: {row['nombre']}")
                    
                    # Preparar contexto
                    meses = {
                        "01": "Enero", "02": "Febrero", "03": "Marzo",
                        "04": "Abril", "05": "Mayo", "06": "Junio",
                        "07": "Julio", "08": "Agosto", "09": "Septiembre",
                        "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
                    }

                    fecha = row["fecha"] if not pd.isna(row["fecha"]) else None
                    if fecha:
                        dia = fecha.strftime("%d")
                        mes = meses[fecha.strftime("%m")]
                        año = fecha.strftime("%Y")
                    else:
                        dia = mes = año = ""

                    contexto = {
                        "ITEM": str(row["item"]),    
                        "NOMBRE": str(row["nombre"]),
                        "CEDULA": str(row["cedula"]),
                        "DIA": dia,
                        "MES": mes,
                        "AÑO": año,
                        "COMPANIA": str(row["compañia"]),
                        "HORAS": str(row.get("horas", "")),
                        "ID_FORMACION": str(row["id_formacion"])
                        
                    }

                    # Crear subcarpeta por compañía
                    compania_folder = output_dir / str(row["compañia"]).replace(" ", "_").replace("/", "_")
                    os.makedirs(compania_folder, exist_ok=True)

                    # Generar nombres de archivo
                    plantilla_name = str(row.get("horas", "general"))
                    nombre_base = f"certificado_{plantilla_name}_horas_{row['nombre'].replace(' ', '_')}"
                    nombre_base = "".join(c for c in nombre_base if c.isalnum() or c in (' ', '-', '_')).rstrip()
                    
                    docx_file = compania_folder / f"{nombre_base}.docx"
                    pdf_file = compania_folder / f"{nombre_base}.pdf"

                    # Generar certificado
                    plantilla = DocxTemplate(plantilla_path)
                    plantilla.render(contexto)
                    plantilla.save(str(docx_file))
                    
                    # Actualizar DataFrame
                    df.at[index, "certificado"] = "si"

                    # Conversión a PDF 
                    if ON_WINDOWS:
                        try:
                            convert(str(docx_file), str(pdf_file))
                            os.remove(docx_file)
                            logger.info(f"Certificado PDF creado: {pdf_file}")
                        except Exception as e:
                            logger.error(f"Error convirtiendo a PDF: {e}")
                            # Mantener DOCX si falla la conversión a PDF
                    else:
                        # LibreOffice (Linux/Mac)
                        try:
                            subprocess.run([    
                                "soffice", "--headless", "--convert-to", "pdf", 
                                "--outdir", str(compania_folder), str(docx_file)
                            ], check=True)
                            os.remove(docx_file)
                        except subprocess.CalledProcessError as e:
                            logger.error(f"Error con LibreOffice: {e}")

                    certificados_por_compania[row["compañia"]].append(f"{nombre_base}.pdf")
                    certificados_creados += 1

            except Exception as e:
                logger.error(f"Error procesando fila {index}: {e}")
                continue

        # Se guarda el Excel actualizado

# Guardar Excel actualizado en carpeta Certificados con el mismo nombre del archivo subido
        try:
            original_name = Path(excel_file.filename).name  # nombre original del archivo
            excel_actualizado = output_dir / original_name  # lo guardamos con el mismo nombre
            df.to_excel(excel_actualizado, index=False)
            logger.info(f"Excel actualizado guardado: {excel_actualizado}")
        except Exception as e:
            logger.error(f"Error guardando Excel: {e}")

        # Limpiar COM
        if ON_WINDOWS:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
        try:
            if ON_WINDOWS:
                os.startfile(str(output_dir))
            elif platform.system() == "Darwin":  # Mac
                subprocess.run(["open", str(output_dir)])
            else:  # Linux
                subprocess.run(["xdg-open", str(output_dir)])
        except Exception as e:
            logger.error(f"No se pudo abrir automáticamente la carpeta: {e}")

        logger.info(f"Procesamiento completado. Certificados creados: {certificados_creados}")

        return render_template("success.html")

       

    except Exception as e:
        logger.error(f"Error general en procesamiento: {e}")
        if ON_WINDOWS:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
        return f"Error interno: {str(e)}", 500

@app.errorhandler(500)
def internal_error(error):
    logger.error(f"Error 500: {error}")
    return f"Error interno del servidor: {str(error)}", 500



if __name__ == "__main__":
    try:
        # Verificar archivos necesarios
        template_path = resource_path('templates')
        if not os.path.exists(template_path):
            logger.warning(f"ADVERTENCIA: No se encuentra carpeta templates en {template_path}")
        
        # Abrir navegador después de un delay
        threading.Timer(2.0, open_browser).start()
        
        # Ejecutar Flask
        app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False)
        
    except Exception as e:
        logger.error(f"Error iniciando aplicación: {e}")