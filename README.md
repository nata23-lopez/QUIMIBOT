# QuimiBot web

Adaptación a Streamlit del QuimiBot creado en Google Colab. Conserva la lógica de lectura de PDF, PowerPoint, Excel y Word, búsqueda TF-IDF/coseno, respuestas con Gemini, indicación de fuentes y explicación de ejercicios paso a paso.

## 1. Agregar los materiales

Copia los documentos originales del curso a la carpeta `documentos/`. En el cuaderno original aparecen estos 11 archivos:

- Taller repaso.pdf
- IPQ - Ejercicios de presión MOTT.pdf
- 3 PRESIÓN Y SU MEDICIÓN v2 ok.pptx
- IPQ - GUIA DENSIDAD, PE, GE (1).xlsx
- PROPIEDADES DE FLUIDOS Y SÓLIDOS (1).pptx
- TALLER - CONVERSIONES 2016-1.pptx
- 7 LONGITUD MASA TIEMPO MOL.pptx
- 8 LONGITUD MASA TIEMPO MOL.pptx
- 5 CONCEPTO DE VALOR NUMÉRICO, UNIDAD, DIMENSIÓN, FACTORES DE CONVERSIÓN ok (1).pptx
- 6 SISTEMA DE UNIDADES FUNDAMENTALES, UNIDADES, MÚLTIPLOS Y DERIVADOS.pptx
- 2 PRESENTACION CPIQ ok.pptx

Los nombres son una guía: debes copiar los archivos reales desde tu computador o Google Drive. El código que pegaste muestra los nombres, pero no contiene los archivos mismos.

## 2. Subir a GitHub

Crea un repositorio. Sube el contenido de esta carpeta al nivel principal del repositorio: `app.py`, `requirements.txt`, `README.md`, `.gitignore` y las carpetas `documentos/` y `.streamlit/`. No subas `secrets.toml`, claves API ni carpetas `__pycache__`.

## 3. Publicar con Streamlit Community Cloud

En https://share.streamlit.io/ conecta GitHub, crea una app, selecciona el repositorio y escribe `app.py` como archivo principal.

En la configuración de Secrets de Streamlit agrega:

```toml
GEMINI_API_KEY = "AQ.Ab8RN6IeDH3RG_I3cvee6VKI5s9VXmPjyxRLOQ07aoQzNgIsyQ"
```

No publiques la clave en GitHub ni la pegues en el código. Después de desplegar, Streamlit entregará el enlace de la aplicación.

## Notas

- La aplicación no usa Google Colab ni pide a los estudiantes ejecutar celdas.
- Los documentos se leen desde `documentos/` al iniciar la aplicación.
- Para que los estudiantes puedan consultar, el enlace de Streamlit debe quedar accesible para ellos; el repositorio de GitHub puede mantenerse privado si la configuración de Streamlit lo permite.
- Los PDF escaneados como imagen pueden necesitar OCR; este proyecto extrae texto seleccionable.
- Revisa el consumo y los límites de tu cuenta/API de Gemini antes de compartir el enlace ampliamente.
