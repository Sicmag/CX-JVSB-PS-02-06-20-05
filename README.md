# EscribIA

Prototipo funcional que digitaliza apuntes manuscritos usando IA.

## Funciones
- Convierte fotos de apuntes a documentos de Word
- Genera presentaciones de PowerPoint
- Crea resúmenes de estudio con preguntas de repaso
- Sistema de doble proveedor de IA (Gemini + Groq) para alta disponibilidad

## Archivos
- `config.py`: claves de API (privado)
- `escribia.py`: versión terminal
- `app_streamlit.py`: versión web

## Cómo usar
1. Instalar librerías: `pip install -r requirements.txt`
2. Configurar claves en `config.py`
3. Ejecutar versión terminal: `python escribia.py`
4. Ejecutar versión web: `streamlit run app_streamlit.py`

## Tecnologías
- Python 3.10+
- Google Gemini (proveedor principal)
- Groq (proveedor de respaldo)
- Streamlit (interfaz web)
- python-docx, python-pptx (generación de documentos)