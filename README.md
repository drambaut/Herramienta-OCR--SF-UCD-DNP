# ──────────────────────────────────────────────────────────────
# Herramienta OCR-SF-UCD · Estructura del repositorio
# ──────────────────────────────────────────────────────────────

Esta herramienta permite extraer información de documentos PDF de planillas de seguridad social. Se procesan los archivos PDF y se extrae la información más importante y se inserta en un excel para descargar.
herramienta-ocr-sf-ucd/
├── app/
│ └── main.py 
│
├── requirements.txt 
├── .env.example
├── .gitignore 
├── LICENSE 
└── README.md

# ──────────────────────────────────────────────────────────────
# DESPLIEGUE EN RENDER · Paso a paso
# ──────────────────────────────────────────────────────────────

## 1. Preparar el repositorio en GitHub
   git init
   git add main.py requirements.txt .env.example .gitignore
   git commit -m "feat: initial OCR-SF-UCD app"
   git remote add origin https://github.com/drambaut/Herramienta-OCR--SF-UCD.git
   git push -u origin main

## 2. Variables de entorno
   En el .env.example reemplazar con las credenciales de Azure:

   AZURE_FORM_RECOGNIZER_ENDPOINT   = https://<tu-recurso>.cognitiveservices.azure.com/
   AZURE_FORM_RECOGNIZER_KEY        = <tu-clave>
   AZURE_OPENAI_ENDPOINT            = https://<tu-recurso>.openai.azure.com/
   AZURE_OPENAI_API_KEY             = <tu-clave>
   AZURE_OPENAI_API_VERSION         = 2024-05-01-preview
   AZURE_OPENAI_DEPLOYMENT_NAME     = gpt-4o

   ⚠ NUNCA pongas estas variables en el código ni en archivos que suban a Git.

