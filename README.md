# 📄 Herramienta OCR-SF-UCD

**Extracción automática de planillas de seguridad social colombianas (PILA)**

Esta aplicación web automatiza la lectura, extracción y validación de datos desde documentos PDF (planillas de seguridad social). Fue diseñada para agilizar los procesos de la Subdirección Financiera - Unidad de Costos y Dependencias (DNP), procesando múltiples archivos, aplicando reglas de negocio y generando reportes estructurados listos para su descarga.

## ✨ Características Principales

* **Procesamiento por Lotes:** Permite la carga masiva de múltiples archivos PDF empaquetados en un archivo `.zip`.
* **OCR y Extracción Inteligente:** Integra **Azure AI Document Intelligence** para la lectura de estructuras y **Azure OpenAI (GPT-4o)** para la extracción precisa de entidades (operador, identificaciones, montos y fechas).
* **Validaciones Automáticas:** Evalúa reglas de negocio en tiempo real (ej. montos inusuales, fechas caducadas o identificaciones inválidas).
* **Dashboard Interactivo:** Panel de control con KPIs y gráficos dinámicos generados para visualizar el estado de validación y los montos por operador.
* **Exportación Consolidada:** Descarga de todos los datos extraídos y validados en un formato estructurado de Excel (`.xlsx`).

## 🛠️ Tecnologías y Dependencias

El proyecto está construido principalmente con las siguientes tecnologías:
* **Interfaz y Backend:** Python 3.x, Streamlit.
* **Procesamiento PDF y OCR:** PyMuPDF, Azure AI Document Intelligence.
* **Modelos de Lenguaje:** Azure OpenAI.
* **Manejo de Datos y Visualización:** Pandas, Plotly.

## 📂 Estructura del Repositorio

```text
herramienta-ocr-sf-ucd/
├── app/
│   └── main.py                 # Lógica principal de la aplicación Streamlit
├── requirements.txt            # Dependencias de Python del proyecto
├── .env.example                # Plantilla de variables de entorno
├── .gitignore                  # Archivos y carpetas ignorados por Git
├── LICENSE                     # Licencia del proyecto
└── README.md                   # Documentación principal
```

## 🚀 Guía de Instalación y Uso Local

### 1. Clonar el repositorio
```bash
git clone [https://github.com/drambaut/Herramienta-OCR--SF-UCD.git](https://github.com/drambaut/Herramienta-OCR--SF-UCD.git)
cd Herramienta-OCR--SF-UCD
```

### 2. Crear y activar entorno virtual (Recomendado)
```bash
python -m venv venv
# En Windows:
venv\Scripts\activate
# En macOS/Linux:
source venv/bin/activate
```

### 3. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 4. Configurar Variables de Entorno
Crea un archivo llamado `.env` en la raíz del proyecto basándote en el `.env.example`. Debes completar las credenciales con tus recursos de Azure:

```env
AZURE_FORM_RECOGNIZER_ENDPOINT=https://<tu-recurso>[.cognitiveservices.azure.com/](https://.cognitiveservices.azure.com/)
AZURE_FORM_RECOGNIZER_KEY=<tu_api_key_de_document_intelligence>
AZURE_OPENAI_ENDPOINT=https://<tu-recurso>[.openai.azure.com/](https://.openai.azure.com/)
AZURE_OPENAI_API_KEY=<tu_api_key_de_azure_openai>
AZURE_OPENAI_API_VERSION=2024-02-15-preview
AZURE_OPENAI_ASSISTANT_ID=<asst_id_del_assistant>
AZURE_OPENAI_DEPLOYMENT_NAME=gpt-4o
```
> ⚠️ **IMPORTANTE:** NUNCA subas el archivo `.env` al control de versiones (Git). Asegúrate de que permanezca listado en tu `.gitignore`.

### 5. Ejecutar la aplicación
```bash
streamlit run app/main.py
```

## ☁️ Despliegue en Render (Paso a paso)

Si deseas subir esta aplicación a un entorno de producción como Render, sigue estos pasos:

1. **Sube los cambios a tu repositorio de GitHub:**
   Asegúrate de que tus archivos `app/main.py`, `requirements.txt`, `.env.example` y `.gitignore` estén confirmados en la rama `main`.
2. **Crea un nuevo Web Service en Render** conectado a tu repositorio de GitHub.
3. **Configuración del entorno de Render:**
   * **Build Command:** `pip install -r requirements.txt`
   * **Start Command:** `streamlit run app/main.py --server.port $PORT`
4. **Variables de entorno (Environment Variables):**
   Dentro de la configuración de Render, ve a la sección "Environment" y añade manualmente las variables definidas en el paso de instalación (Endpoint y Keys de Form Recognizer y OpenAI). No subas el archivo `.env`.

## 📄 Licencia

Este proyecto está distribuido bajo la licencia MIT. Copyright (c) 2026 Daniel Felipe Rambaut Lemus.
