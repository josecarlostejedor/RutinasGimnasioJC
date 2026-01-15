import streamlit as st
import pandas as pd
import random
import os
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from io import BytesIO
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Rutinas Pro", layout="wide")

# --- CARGAR BASE DE DATOS ROBUSTA ---
@st.cache_data
def cargar_ejercicios():
    try:
        if os.path.exists("DB_EJERCICIOS.xlsx"):
            df = pd.read_excel("DB_EJERCICIOS.xlsx")
            
            # 1. Limpiar nombres de columnas (quitar espacios y poner minúsculas)
            df.columns = df.columns.str.strip().str.lower()
            
            # 2. Verificar columnas obligatorias
            if 'nombre' not in df.columns:
                # Intentar arreglar si se llama 'ejercicio'
                if 'ejercicio' in df.columns:
                    df.rename(columns={'ejercicio': 'nombre'}, inplace=True)
                else:
                    return f"Error: Falta la columna 'nombre' en el Excel. Columnas encontradas: {list(df.columns)}"
            
            # 3. Rellenar columnas opcionales si faltan
            if 'tipo' not in df.columns:
                df['tipo'] = "General" # Valor por defecto si no existe la columna
            if 'imagen' not in df.columns:
                df['imagen'] = ""
            if 'desc' not in df.columns:
                df['desc'] = ""

            # Limpiar datos
            df['tipo'] = df['tipo'].fillna("General").astype(str)
            df['imagen'] = df['imagen'].fillna("").astype(str)
            
            return df.to_dict('records')
        else:
            return None
    except Exception as e:
        return f"Error crítico leyendo Excel: {str(e)}"

DB_EJERCICIOS = cargar_ejercicios()

# --- LÓGICA DE OBJETIVOS ---
def obtener_parametros(objetivo):
    if objetivo == "Hipertrofia":
        return {"reps": "6-12", "int_min": 0.60, "int_max": 0.85, "descanso": "1-3 min"}
    elif objetivo == "Fuerza Máxima":
        return {"reps": "1-5", "int_min": 0.85, "int_max": 0.95, "descanso": "3-5 min"}
    elif objetivo == "Resistencia":
        return {"reps": "15-20", "int_min": 0.40, "int_max": 0.60, "descanso": "30-60 s"}

# --- GENERADOR WORD CON IMÁGENES ---
def generar_word(rutina_df, objetivo, alumno, tipo_rutina):
    doc = Document()
    
    # Configurar página Horizontal
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)

    # Encabezado
    header = doc.add_table(rows=1, cols=3)
    header.autofit = False
    header.columns[0].width = Inches(4)
    header.columns[2].width = Inches(2.5)
    
    header.cell(0,0).text = f"RUTINA: {tipo_rutina.upper()}"
    header.cell(0,1).text = f"OBJETIVO: {objetivo.upper()}"
    c3 = header.cell(0,2)
    c3.text = f"Fecha: {datetime.now().strftime('%d/%m/%Y')}\nIES / CLUB"
    c3.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_paragraph(f"Alumno/a: {alumno}")
    doc.add_paragraph("_" * 115)

    # Estructura Principal
    main_table = doc.add_table(rows=1, cols=2)
    main_table.autofit = False
    main_table.columns[0].width = Inches(5.5)
    main_table.columns[1].width = Inches(5.0)

    # --- IZQUIERDA: IMÁGENES ---
    left_cell = main_table.cell(0,0)
    rows_needed = (len(rutina_df) + 1) // 2
    visual_table = left_cell.add_table(rows=rows_needed, cols=2)
    visual_table.style = 'Table Grid'
    
    for idx, row_data in rutina_df.iterrows():
        r = idx // 2
        c = idx % 2
        cell = visual_table.cell(r, c)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Intentar insertar imagen
        img_name = str(row_data['Imagen']).strip()
        img_path = os.path.join("images", img_name)
        
        if img_name and os.path.exists(img_path):
            try:
                run = p.add_run()
                run.add_picture(img_path, width=Inches(1.8))
                p.add_run("\n")
            except:
                p.add_run("[Error img]\n")
        else:
            p.add_run("\n[FOTO]\n")

        run_txt = p.add_run(row_data['Ejercicio'])
        run_txt.font.bold = True
        run_txt.font.size = Pt(9)

    # --- DERECHA: DATOS ---
    right_cell = main_table.cell(0,1)
    reg_table = right_cell.add_table(rows=1, cols=3)
    reg_table.style = 'Table Grid'
    
    hdr = reg_table.rows[0].cells
    hdr[0].text = "Ejercicio"
    hdr[1].text = "Series/Reps"
    hdr[2].text = "Kg"
    
    for idx, row_data in rutina_df.iterrows():
        row = reg_table.add_row().cells
        row[0].text = f"{idx+1}. {row_data['Ejercicio']} ({row_data['Tipo']})"
        row[1].text = f"{row_data['Series']} x {row_data['Reps']}"
        row[2].text = str(row_data['Peso'])
        
        row_d = reg_table.add_row().cells
        m = row_d[0].merge(row_d[2])
        m.text = f"Descanso: {row_data['Descanso']} | Notas: _________________"
        m.paragraphs[0].runs[0].font.size = Pt(8)

    # Footer
    doc.add_paragraph("\n")
    footer = doc.add_table(rows=1, cols=1)
    footer.style = 'Table Grid'
    footer.cell(0,0).text = "Escala de Borg: 6-7 (Muy ligero) | 12-13 (Algo duro) | 18-20 (Agotamiento)"
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ DE USUARIO ---
st.title("🏋️ Generador de Rutinas por Material")

# Validaciones iniciales
if DB_EJERCICIOS is None:
    st.error("No se encuentra el archivo 'DB_EJERCICIOS.xlsx'. Súbelo al repositorio.")
    st.stop()
elif isinstance(DB_EJERCICIOS, str):
    st.error(DB_EJERCICIOS) # Muestra el error detallado de columnas
    st.stop()

# --- SIDEBAR: FILTROS ---
st.sidebar.header("Configuración de Rutina")
alumno = st.sidebar.text_input("Nombre Alumno:", "Atleta")
objetivo = st.sidebar.selectbox("Objetivo", ["Hipertrofia", "Fuerza Máxima", "Resistencia"])

# 1. Obtener tipos únicos del Excel para el filtro
tipos_disponibles = sorted(list(set([e['tipo'] for e in DB_EJERCICIOS])))

# 2. Selector de Tipo
tipos_seleccionados = st.sidebar.multiselect(
    "Selecciona Material / Tipo:",
    options=tipos_disponibles,
    default=tipos_disponibles # Por defecto selecciona todos
)

# 3. Filtrar la base de datos según selección
if not tipos_seleccionados:
    st.warning("Selecciona al menos un tipo de material.")
    st.stop()

ejercicios_filtrados = [e for e in DB_EJERCICIOS if e['tipo'] in tipos_seleccionados]
num_disponibles = len(ejercicios_filtrados)

if num_disponibles == 0:
    st.warning("No hay ejercicios que coincidan con ese filtro.")
    st.stop()

st.sidebar.markdown(f"**Ejercicios disponibles:** {num_disponibles}")

# 4. Slider dinámico
num_ej = st.sidebar.slider("Nº Ejercicios", 1, min(10, num_disponibles), min(6, num_disponibles))

# --- ÁREA PRINCIPAL ---
st.subheader(f"Selección de Ejercicios ({', '.join(tipos_seleccionados)})")

lista_nombres_filtrados = [e['nombre'] for e in ejercicios_filtrados]
seleccion = st.multiselect("Elige ejercicios específicos:", lista_nombres_filtrados, max_selections=num_ej)

# Lógica de relleno automático
ejercicios_finales_data = []
nombres_elegidos = seleccion.copy()

if len(nombres_elegidos) < num_ej:
    restantes = [e for e in ejercicios_filtrados if e['nombre'] not in nombres_elegidos]
    faltan = num_ej - len(nombres_elegidos)
    if faltan > 0:
        extra = random.sample(restantes, faltan)
        nombres_elegidos.extend([e['nombre'] for e in extra])

# Recuperar datos completos
for nombre in nombres_elegidos:
    # Buscamos en la lista filtrada
    item = next(x for x in ejercicios_filtrados if x['nombre'] == nombre)
    ejercicios_finales_data.append(item)

# Inputs de RM
st.markdown("---")
st.subheader("Cargas (RM)")
cols = st.columns(3)
inputs_rm = {}

for i, item in enumerate(ejercicios_finales_data):
    with cols[i % 3]:
        # Mostramos también el tipo para guiar al usuario
        label = f"{item['nombre']} ({item['tipo']})"
        inputs_rm[item['nombre']] = st.number_input(f"RM {label}", value=50, step=5, key=item['nombre'])

# Botón Generar
if st.button("Generar PDF (Word)"):
    params = obtener_parametros(objetivo)
    rutina_export = []
    
    for item in ejercicios_finales_data:
        rm = inputs_rm[item['nombre']]
        intensidad = random.uniform(params['int_min'], params['int_max'])
        peso = round(rm * intensidad)
        
        rutina_export.append({
            "Ejercicio": item['nombre'],
            "Tipo": item['tipo'],
            "Imagen": item['imagen'],
            "Series": 4,
            "Reps": params['reps'],
            "Peso": peso,
            "Descanso": params['descanso']
        })
        
    df_export = pd.DataFrame(rutina_export)
    
    # Texto para el título del Word (ej: "Multipower + TRX")
    titulo_rutina = " + ".join(tipos_seleccionados) if len(tipos_seleccionados) < 3 else "MIXTA"
    
    docx = generar_word(df_export, objetivo, alumno, titulo_rutina)
    
    st.success("¡Rutina generada!")
    st.download_button(
        "📥 Descargar Word", 
        docx, 
        f"rutina_{alumno}_{datetime.now().strftime('%Y%m%d')}.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
