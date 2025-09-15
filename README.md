# Generador de Certificados (Aplicación Local)

Esta aplicación genera certificados en **PDF** a partir de un archivo Excel cargado por el usuario.  
Está construida con **Flask** y empaquetada como ejecutable (`app.exe`) mediante PyInstaller, para usarse en entornos locales sin necesidad de instalar dependencias.

---

## 🚀 Características principales

- Interfaz web local que se abre automáticamente en el navegador (`http://127.0.0.1:5000/`).
- Carga de un archivo Excel con la lista de participantes.
- Lectura de los campos requeridos: nombre, cédula, fecha, compañía, horas, id_formación, item, certificado.
- Generación de certificados personalizados usando una plantilla Word (`plantilla.docx`).
- Conversión automática a PDF (usa Microsoft Word en Windows o LibreOffice en Linux/Mac).
- Organización automática de certificados en carpetas por compañía dentro de la carpeta **Descargas/Certificados**.
- Actualización del Excel marcando los registros procesados con `certificado = si`.

---

## 📂 Estructura del repositorio

```
.
├── app.py                 # Aplicación principal Flask
├── templates/             # Plantillas HTML de la interfaz web
├── static/                # Archivos estáticos (CSS, JS, imágenes)
├── plantilla.docx         # Plantilla base para los certificados
└── README.md              # Este archivo
```

---

## ⚙️ Requisitos

- Windows, Linux o macOS
- Python 3.10+ (para desarrollo)
- Microsoft Word (Windows) o LibreOffice (Linux/Mac) para conversión a PDF
- Excel de entrada con las siguientes columnas mínimas:
  - `nombre`
  - `cedula`
  - `fecha`
  - `compañia`
  - `horas`
  - `id_formacion`
  - `item`
  - `certificado`

---

## ▶️ Uso

### Como desarrollador (código fuente)
```bash
# Clonar el repositorio
git clone https://github.com/tuusuario/generador-certificados.git
cd generador-certificados

# Crear entorno virtual
python -m venv venv
venv\Scripts\activate  # Windows
source venv/bin/activate # Linux/Mac

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar aplicación
python app.py
```

Luego abre en el navegador: [http://127.0.0.1:5000/](http://127.0.0.1:5000/)

### Como usuario final (ejecutable)
- Descarga el archivo `app.exe` desde la sección de releases.
- Haz doble clic en `app.exe`.
- El navegador se abrirá automáticamente en [http://127.0.0.1:5000/](http://127.0.0.1:5000/).
- Carga tu archivo Excel y espera a que los certificados se generen en **Descargas/Certificados**.

---

## 🧪 Notas importantes

- La plantilla debe contener las variables delimitadas con `{{ }}`, por ejemplo:

```docx
Certificado No: {{CERTIFICADO}}
Otorgado a: {{NOMBRE}}
Identificación: {{CEDULA}}
Fecha: {{DIA}} de {{MES}} de {{AÑO}}
Compañía: {{COMPANIA}}
Horas: {{HORAS}} HORAS CERTIFICADAS
```

- Los certificados se generan en PDF dentro de subcarpetas por compañía.
- El Excel original se actualiza y se guarda en la misma carpeta `Descargas/Certificados` con el mismo nombre del archivo cargado.

---

## 📜 Licencia

Este proyecto está bajo la licencia MIT. Puedes usarlo, modificarlo y distribuirlo libremente.
