import streamlit as st
import pandas as pd
import random
import os
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Entrenador Pro v2", layout="wide")

# --- FUNCIONES AUXILIARES WORD (ESTILOS Y COLORES) ---
def set_cell_bg_color(cell, hex_color):
    """Pinta el fondo de una celda de Word con un color HEX"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def style_header_cell(cell, text):
    cell.text = text
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255) # Texto blanco
    set_cell_bg_color(cell, "2E4053") # Fondo gris oscuro profesional

# --- CARGAR BASE DE DATOS ---
@st.cache_data
def cargar_ejercicios():
    try:
        if os.path.exists("DB_EJERCICIOS.xlsx"):
            df = pd.read_excel("DB_EJERCICIOS.xlsx")
            df.columns = df.columns.str.strip().str.lower()
            
            # Normalización de columnas
            if 'nombre' not in df.columns:
                if 'ejercicio' in df.columns: df.rename(columns={'ejercicio': 'nombre'}, inplace=True)
            
            campos_opcionales = ['tipo', 'imagen', 'desc']
            for campo in campos_opcionales:
                if campo not in df.columns: df[campo] = ""
            
            # Rellenar vacíos
            df = df.fillna("")
            return df.to_dict('records')
        else:
            return None
    except Exception as e:
        return f"Error: {str(e)}"

DB_EJERCICIOS = cargar_ejercicios()

# --- LÓGICA DE PARAMETROS ---
def obtener_parametros(objetivo):
    if objetivo == "Hipertrofia":
        return {"reps": "8-12", "descanso": "90 seg", "notas": "Controlar excéntrica"}
    elif objetivo == "Fuerza Máxima":
        return {"reps": "3-5", "descanso": "3-5 min", "notas": "Explosivo concéntrico"}
    elif objetivo == "Resistencia":
        return {"reps": "15-20", "descanso": "45 seg", "notas": "Ritmo constante"}

# --- GENERADOR DE WORD PRO ---
def generar_word_pro(rutina_df, objetivo, alumno, tipo_rutina):
    doc = Document()
    
    # 1. Configuración Página (Vertical suele ser mejor para lista larga, pero mantenemos Horizontal para layout visual)
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    # 2. Encabezado Profesional
    head_tbl = doc.add_table(rows=1, cols=2)
    head_tbl.autofit = False
    head_tbl.columns[0].width = Inches(8)
    head_tbl.columns[1].width = Inches(3)
    
    # Celda Izq: Título
    c1 = head_tbl.cell(0,0)
    p = c1.paragraphs[0]
    r1 = p.add_run(f"PROGRAMA: {tipo_rutina.upper()}\n")
    r1.font.bold = True
    r1.font.size = Pt(16)
    r1.font.color.rgb = RGBColor(41, 128, 185) # Azul bonito
    p.add_run(f"OBJETIVO: {objetivo} | ALUMNO: {alumno}")

    # Celda Der: Fecha
    c2 = head_tbl.cell(0,1)
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run(f"FECHA: {datetime.now().strftime('%d/%m/%Y')}\n").bold = True
    p2.add_run("Entrenamiento Funcional")

    doc.add_paragraph("_" * 90)

    # 3. SECCIÓN 1: GALERÍA VISUAL (Arriba)
    doc.add_heading('1. Guía Visual de Ejercicios', level=2)
    
    # Calculamos filas necesarias para 4 columnas
    num_ej = len(rutina_df)
    cols_visual = 4 
    rows_visual = (num_ej + cols_visual - 1) // cols_visual
    
    vis_table = doc.add_table(rows=rows_visual, cols=cols_visual)
    vis_table.style = 'Table Grid'
    vis_table.autofit = False
    
    # Ajustar ancho de columnas visuales
    for col in vis_table.columns:
        col.width = Inches(2.5)

    for i, row_data in enumerate(rutina_df.to_dict('records')):
        r = i // cols_visual
        c = i % cols_visual
        cell = vis_table.cell(r, c)
        
        # Párrafo centrado
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Insertar Imagen
        img_name = str(row_data['Imagen']).strip()
        img_path = os.path.join("images", img_name)
        
        if img_name and os.path.exists(img_path):
            try:
                run = p.add_run()
                run.add_picture(img_path, width=Inches(1.8), height=Inches(1.5))
                p.add_run("\n")
            except:
                p.add_run("⚠️ Error Fichero\n")
        else:
            # Placeholder bonito si falla
            p.add_run("\n[FOTO]\n")
        
        # Nombre ejercicio
        run_nom = p.add_run(row_data['Ejercicio'])
        run_nom.font.bold = True
        run_nom.font.size = Pt(10)

    doc.add_paragraph("\n") # Espacio

    # 4. SECCIÓN 2: TABLA TÉCNICA (Abajo)
    doc.add_heading('2. Rutina Detallada', level=2)
    
    tech_table = doc.add_table(rows=1, cols=6)
    tech_table.style = 'Table Grid'
    
    # Encabezados
    headers = ["#", "Ejercicio", "Series x Reps", "Carga (Kg)", "Descanso", "Notas"]
    for i, h in enumerate(headers):
        style_header_cell(tech_table.rows[0].cells[i], h)
        
    for idx, row_data in rutina_df.iterrows():
        row_cells = tech_table.add_row().cells
        row_cells[0].text = str(idx + 1)
        row_cells[1].text = row_data['Ejercicio']
        row_cells[2].text = f"4 x {row_data['Reps']}"
        row_cells[3].text = str(row_data['Peso'])
        row_cells[4].text = row_data['Descanso']
        row_cells[5].text = "" # Espacio para escribir a mano

    doc.add_paragraph("\n")

    # 5. SECCIÓN 3: ESCALA DE BORG VISUAL (Bonita)
    doc.add_heading('3. Percepción de Esfuerzo (RPE)', level=3)
    
    borg_table = doc.add_table(rows=2, cols=5)
    borg_table.style = 'Table Grid'
    borg_table.autofit = True
    
    # Definición de la escala visual
    borg_data = [
        {"val": "6-8", "txt": "Muy Ligero", "icon": "🙂", "color": "A9DFBF"}, # Verde claro
        {"val": "9-11", "txt": "Ligero", "icon": "😌", "color": "D4EFDF"},    # Verde muy claro
        {"val": "12-14", "txt": "Algo Duro", "icon": "😐", "color": "F9E79F"}, # Amarillo
        {"val": "15-17", "txt": "Duro", "icon": "😓", "color": "F5CBA7"},    # Naranja
        {"val": "18-20", "txt": "Máximo", "icon": "🥵", "color": "E6B0AA"}    # Rojo claro
    ]
    
    # Rellenar tabla Borg
    row_icons = borg_table.rows[0]
    row_text = borg_table.rows[1]
    
    for i, data in enumerate(borg_data):
        # Icono y Valor
        c1 = row_icons.cells[i]
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run1 = p1.add_run(f"{data['icon']}\n{data['val']}")
        run1.font.size = Pt(14)
        set_cell_bg_color(c1, data['color'])
        
        # Texto descriptivo
        c2 = row_text.cells[i]
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.add_run(data['txt']).font.bold = True
        set_cell_bg_color(c2, data['color'])

    # Guardar
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ STREAMLIT ---
st.title("🏋️ Generador de Rutinas V2 (Layout Vertical)")

# --- DEBUGGER DE IMÁGENES ---
# Esto te ayudará a saber si las imágenes están bien subidas
with st.expander("🛠️ Diagnóstico de Imágenes (Abrir si salen errores)"):
    if os.path.exists("images"):
        archivos = os.listdir("images")
        st.write(f"✅ Carpeta 'images' encontrada. Contiene {len(archivos)} archivos:")
        st.write(archivos)
    else:
        st.error("❌ NO existe la carpeta 'images' en el repositorio. Crea la carpeta y sube las fotos.")

# --- VALIDACIÓN EXCEL ---
if DB_EJERCICIOS is None:
    st.error("Sube el archivo DB_EJERCICIOS.xlsx")
    st.stop()
elif isinstance(DB_EJERCICIOS, str):
    st.error(DB_EJERCICIOS)
    st.stop()

# --- SIDEBAR ---
st.sidebar.header("Datos Sesión")
alumno = st.sidebar.text_input("Alumno:", "Atleta")
objetivo = st.sidebar.selectbox("Objetivo", ["Hipertrofia", "Fuerza Máxima", "Resistencia"])

# Filtro Material
tipos = sorted(list(set([e['tipo'] for e in DB_EJERCICIOS if e['tipo']])))
sel_tipos = st.sidebar.multiselect("Material:", options=tipos, default=tipos)

if not sel_tipos: st.stop()

ej_filtrados = [e for e in DB_EJERCICIOS if e['tipo'] in sel_tipos]
num_ej = st.sidebar.slider("Nº Ejercicios", 1, min(10, len(ej_filtrados)), 6)

# Selección
st.subheader(f"Selección ({', '.join(sel_tipos)})")
nombres_fil = [e['nombre'] for e in ej_filtrados]
seleccion = st.multiselect("Ejercicios:", nombres_fil, max_selections=num_ej)

# Relleno auto
seleccionados_data = []
nombres_finales = seleccion.copy()
if len(nombres_finales) < num_ej:
    pool = [x for x in ej_filtrados if x['nombre'] not in nombres_finales]
    needed = num_ej - len(nombres_finales)
    if needed <= len(pool):
        extras = random.sample(pool, needed)
        nombres_finales.extend([x['nombre'] for x in extras])
        
for nom in nombres_finales:
    seleccionados_data.append(next(x for x in ej_filtrados if x['nombre'] == nom))

# Inputs Cargas
st.markdown("---")
cols = st.columns(3)
rm_inputs = {}
for i, ej in enumerate(seleccionados_data):
    with cols[i%3]:
        rm_inputs[ej['nombre']] = st.number_input(f"RM {ej['nombre']}", value=50, step=5)

if st.button("Generar Word Profesional"):
    params = obtener_parametros(objetivo)
    rutina_export = []
    
    for item in seleccionados_data:
        rm = rm_inputs[item['nombre']]
        # Lógica de peso
        factor = 0.75 if objetivo == "Hipertrofia" else (0.90 if objetivo == "Fuerza Máxima" else 0.50)
        peso = int(rm * factor)
        
        rutina_export.append({
            "Ejercicio": item['nombre'],
            "Imagen": item['imagen'],
            "Reps": params['reps'],
            "Peso": peso,
            "Descanso": params['descanso']
        })
        
    df = pd.DataFrame(rutina_export)
    
    docx = generar_word_pro(df, objetivo, alumno, " + ".join(sel_tipos))
    
    st.success("Documento generado. Revisa la sección de descargas.")
    st.download_button("📥 Descargar Rutina .docx", docx, f"Rutina_{alumno}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
