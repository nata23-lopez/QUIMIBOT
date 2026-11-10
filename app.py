import os
from pathlib import Path

import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from pypdf import PdfReader
from pptx import Presentation
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

BASE_DIR = Path(__file__).parent
DOCS_DIR = BASE_DIR / "documentos"
MODELO_PRINCIPAL = "gemini-3.5-flash"
MODELOS_RESPALDO = ["gemini-3.8-flash", "gemini-3-flash-preview", "gemini-2.5-flash", "gemini-flash-latest"]
TAMANO_FRAGMENTO = 2000

INSTRUCCIONES = """
Eres QuimiBot, un tutor universitario especializado en Introducción a los Procesos Químicos.
Ayudas a estudiantes de Ingeniería Industrial. Responde siempre en español, con palabras sencillas,
explicaciones claras y ejemplos. Utiliza primero los documentos del curso. No inventes datos ni
resultados. Indica el nombre del archivo y la página, diapositiva u hoja consultada. Si los materiales
no contienen información suficiente, dilo claramente y diferencia el conocimiento general.

Cuando el estudiante envíe un ejercicio, desarrolla la solución en este orden:
1. Interpretación del enunciado.
2. Datos conocidos con unidades.
3. Incógnitas.
4. Fórmulas y explicación de variables.
5. Procedimiento paso a paso, sin saltarte operaciones.
6. Conversiones y manejo de unidades.
7. Resultado final claro.
8. Comprobación del resultado.

Temas: balance de materia, balance de energía, sistemas y conversión de unidades, masa, volumen,
densidad, mol, número de Avogadro, presión, concentraciones, operaciones y procesos químicos.
Si faltan datos, explica qué información hace falta.
"""

PALABRAS_VACIAS = [
    "a", "al", "algo", "como", "con", "cual", "cuales", "cuando", "de", "del", "el",
    "ella", "en", "entre", "es", "esta", "este", "esto", "fue", "ha", "hay", "la",
    "las", "le", "lo", "los", "mas", "me", "mi", "muy", "no", "o", "para", "pero",
    "por", "porque", "que", "se", "si", "sin", "sobre", "son", "su", "sus", "tambien",
    "te", "tu", "un", "una", "uno", "y", "ya", "explica", "explicame", "dime", "favor",
    "puedes", "quiero", "necesito"
]


def dividir_texto(texto, tamano=TAMANO_FRAGMENTO):
    partes, actual = [], ""
    for linea in texto.split("\n"):
        if len(actual) + len(linea) > tamano and actual:
            partes.append(actual)
            actual = ""
        actual += linea + "\n"
    if actual.strip():
        partes.append(actual)
    return partes


def leer_pdf(ruta):
    salida = []
    for numero, pagina in enumerate(PdfReader(str(ruta)).pages, start=1):
        texto = pagina.extract_text() or ""
        if texto.strip():
            salida.append({"archivo": ruta.name, "ubicacion": f"Página {numero}", "texto": texto})
    return salida


def leer_pptx(ruta):
    salida = []
    for numero, diapositiva in enumerate(Presentation(str(ruta)).slides, start=1):
        textos = []
        for forma in diapositiva.shapes:
            if getattr(forma, "has_text_frame", False) and forma.text.strip():
                textos.append(forma.text)
            if getattr(forma, "has_table", False):
                for fila in forma.table.rows:
                    textos.append(" | ".join(celda.text for celda in fila.cells))
        try:
            if diapositiva.has_notes_slide:
                notas = diapositiva.notes_slide.notes_text_frame.text
                if notas.strip():
                    textos.append("Notas: " + notas)
        except Exception:
            pass
        texto = "\n".join(textos)
        if texto.strip():
            salida.append({"archivo": ruta.name, "ubicacion": f"Diapositiva {numero}", "texto": texto})
    return salida


def leer_excel(ruta):
    salida = []
    hojas = pd.read_excel(ruta, sheet_name=None)
    for nombre_hoja, tabla in hojas.items():
        tabla = tabla.dropna(how="all").dropna(axis=1, how="all")
        if tabla.empty:
            continue
        for i, parte in enumerate(dividir_texto(tabla.to_string(index=False)), start=1):
            ubicacion = f"Hoja '{nombre_hoja}'" + (f" (parte {i})" if len(dividir_texto(tabla.to_string(index=False))) > 1 else "")
            salida.append({"archivo": ruta.name, "ubicacion": ubicacion, "texto": parte})
    return salida


def leer_docx(ruta):
    doc = Document(str(ruta))
    textos = [p.text for p in doc.paragraphs if p.text.strip()]
    for tabla in doc.tables:
        for fila in tabla.rows:
            textos.append(" | ".join(celda.text for celda in fila.cells))
    partes = dividir_texto("\n".join(textos))
    return [{"archivo": ruta.name,
             "ubicacion": f"Sección {i}" if len(partes) > 1 else "Documento completo",
             "texto": parte} for i, parte in enumerate(partes, start=1)]


LECTORES = {".pdf": leer_pdf, ".pptx": leer_pptx, ".xlsx": leer_excel,
            ".xls": leer_excel, ".docx": leer_docx}


@st.cache_resource(show_spinner="Leyendo los materiales de la asignatura...")
def cargar_materiales():
    documentos = []
    errores = []
    if not DOCS_DIR.exists():
        return documentos, errores
    for ruta in sorted(DOCS_DIR.iterdir()):
        if not ruta.is_file() or ruta.name.startswith("~$"):
            continue
        lector = LECTORES.get(ruta.suffix.lower())
        if lector is None:
            continue
        try:
            fragmentos = lector(ruta)
            documentos.extend(fragmentos)
            if not fragmentos:
                errores.append(f"{ruta.name}: no se encontró texto; si es escaneado, necesita OCR.")
        except Exception as exc:
            errores.append(f"{ruta.name}: no se pudo leer ({exc}).")
    return documentos, errores


@st.cache_resource(show_spinner=False)
def crear_buscador(textos):
    vectorizador = TfidfVectorizer(
        lowercase=True, strip_accents="unicode", stop_words=PALABRAS_VACIAS,
        ngram_range=(1, 2), sublinear_tf=True
    )
    matriz = vectorizador.fit_transform(textos)
    return vectorizador, matriz


def buscar_informacion(pregunta, documentos, vectorizador, matriz, cantidad=5):
    similitudes = cosine_similarity(vectorizador.transform([pregunta]), matriz).flatten()
    resultados = []
    for indice in similitudes.argsort()[::-1]:
        if similitudes[indice] < 0.05 or len(resultados) == cantidad:
            break
        resultados.append(documentos[indice])
    return resultados


def generar_respuesta(pregunta, historial, documentos, vectorizador, matriz, client):
    resultados = buscar_informacion(pregunta, documentos, vectorizador, matriz)
    if resultados:
        contexto = "\n".join(
            f"ARCHIVO: {r['archivo']}\nUBICACIÓN: {r['ubicacion']}\nCONTENIDO:\n{r['texto']}\n-------------------------"
            for r in resultados
        )
    else:
        contexto = "No se encontró información relacionada en los materiales de clase."

    conversacion = "\n".join(
        f"{m['role']}: {m['content']}" for m in historial[-8:]
    ) or "(Esta es la primera pregunta.)"
    prompt = f"""INFORMACIÓN DE LOS MATERIALES DE CLASE:\n{contexto}

CONVERSACIÓN ANTERIOR:\n{conversacion}

NUEVA PREGUNTA DEL ESTUDIANTE:\n{pregunta}

Responde teniendo en cuenta la conversación anterior. Si los materiales no tienen información
suficiente, acláralo y no inventes datos. Al final indica las fuentes consultadas (archivo y ubicación)."""

    modelos = [MODELO_PRINCIPAL] + [m for m in MODELOS_RESPALDO if m != MODELO_PRINCIPAL]
    ultimo_error = None
    for modelo in modelos:
        try:
            respuesta = client.models.generate_content(
                model=modelo,
                contents=prompt,
                config=types.GenerateContentConfig(system_instruction=INSTRUCCIONES),
            )
            if respuesta.text:
                return respuesta.text
        except Exception as exc:
            ultimo_error = exc
    raise RuntimeError(f"Gemini no pudo generar una respuesta en este momento. Detalle: {ultimo_error}")


st.set_page_config(page_title="QuimiBot | Procesos Químicos", page_icon="🧪", layout="centered")
st.title("🧪 Bienvenido a QuimiBot")
st.caption("Tutor virtual · Introducción a los Procesos Químicos")
st.write("Escribe una pregunta o comparte un ejercicio. QuimiBot te ayudará a resolverlo paso a paso usando primero los materiales de la asignatura.")

api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))
if not api_key:
    st.error("La aplicación todavía no tiene configurada la clave de Gemini. El administrador debe agregarla en Secrets.")
    st.stop()

if not DOCS_DIR.exists() or not any(p.is_file() and p.suffix.lower() in LECTORES for p in DOCS_DIR.iterdir()):
    st.error("Todavía no se han agregado los materiales de la asignatura. El administrador debe colocar los archivos PDF, PPTX, XLSX, XLS o DOCX en la carpeta 'documentos'.")
    st.stop()

documentos, errores = cargar_materiales()
if errores:
    with st.expander("Avisos al leer documentos"):
        for aviso in errores:
            st.warning(aviso)
if not documentos:
    st.error("No se pudo extraer texto de los documentos. Comprueba los archivos de la carpeta 'documentos'.")
    st.stop()

try:
    vectorizador, matriz = crear_buscador(tuple(d["texto"] for d in documentos))
except ValueError:
    st.error("Los documentos no contienen texto suficiente para crear el buscador.")
    st.stop()

st.caption(f"Materiales disponibles: {len({d['archivo'] for d in documentos})} archivos · {len(documentos)} fragmentos de texto")

if "messages" not in st.session_state:
    st.session_state.messages = []

with st.sidebar:
    st.subheader("QuimiBot")
    st.write("Asistente de estudio para Introducción a los Procesos Químicos.")
    if st.button("↻ Nueva conversación", use_container_width=True):
        st.session_state.messages = []
        st.rerun()
    st.divider()
    st.caption("Las respuestas se apoyan primero en los documentos cargados por el docente.")

for mensaje in st.session_state.messages:
    with st.chat_message(mensaje["role"]):
        st.markdown(mensaje["content"])

pregunta = st.chat_input("Escribe tu pregunta o ejercicio...")
if pregunta:
    st.session_state.messages.append({"role": "user", "content": pregunta})
    with st.chat_message("user"):
        st.markdown(pregunta)
    with st.chat_message("assistant"):
        try:
            cliente = genai.Client(api_key=api_key)
            with st.spinner("QuimiBot está preparando la explicación..."):
                respuesta = generar_respuesta(
                    pregunta, st.session_state.messages[:-1], documentos,
                    vectorizador, matriz, cliente
                )
            st.markdown(respuesta)
            st.session_state.messages.append({"role": "assistant", "content": respuesta})
        except Exception as exc:
            st.error(str(exc))
