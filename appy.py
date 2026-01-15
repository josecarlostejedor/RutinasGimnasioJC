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
st.set_page_config(page_title="Entrenador Pro Científico", layout="wide")

# --- ESTILOS WORD ---
def set_cell_bg_color(cell, hex_color):
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
    run.font.color.rgb = RGBColor(255, 255, 255)
    set_cell_bg_color(cell, "2E4053")

# --- BUSCADOR DE IMÁGENES ---
def encontrar_imagen_recursiva(nombre_objetivo):
    if not nombre_objetivo or pd.isna(nombre_objetivo):
        return None, "Celda Vacía"

    nombre_limpio = str(nombre_objetivo).strip().lower()
    nombre_base_excel = os.path.splitext(nombre_limpio)[0]

    for root, dirs, files in os.walk("."):
        for filename in files:
            if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                filename_base = os.path.splitext(filename)[0].lower()
                if filename.lower() == nombre_limpio:
                    return os.path.join(root, filename), "Exacta"
                if filename_base == nombre_base_excel:
                    return os.path.join(root, filename), "Por Nombre"
    return None, f"No encontrado"

# --- CARGAR EXCEL ---
@st.cache_data
def cargar_ejercicios():
    try:
        if os.path.exists("DB_EJERCICIOS.xlsx"):
            df = pd.read_excel("DB_EJERCICIOS.xlsx")
            df.columns = df.columns.str.strip().str.lower()
            if 'nombre' not in df.columns:
                if 'ejercicio' in df.columns: df.rename(columns={'ejercicio': 'nombre'}, inplace=True)
            for col in ['tipo', 'imagen', 'desc']:
                if col not in df.columns: df[col] = ""
            df = df.fillna("")
            return df.to_dict('records')
        else:
            return None
    except Exception as e:
        return f"Error: {str(e)}"

DB_EJERCICIOS = cargar_ejercicios()

# --- GENERADOR WORD ---
def generar_word_final(rutina_df, objetivo, alumno, titulo_material, intensidad_str):
    doc = Document()
    
    # Configuración Página
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)

    # Encabezado
    head_tbl = doc.add_table(rows=1, cols=2)
    head_tbl.autofit = False
    head_tbl.columns[0].width = Inches(8)
    head_tbl.columns[1].width = Inches(3)
    
    c1 = head_tbl.cell(0,0)
    p = c1.paragraphs[0]
    r1 = p.add_run(f"PROGRAMA DE ENTRENAMIENTO DE: {titulo_material.upper()}\n")
    r1.font.bold = True
    r1.font.size = Pt(16)
    r1.font.color.rgb = RGBColor(41, 128, 185)
    
    nombre_mostrar = alumno if alumno.strip() else "ALUMNO"
    # Añadimos la intensidad seleccionada al encabezado
    p.add_run(f"OBJETIVO: {objetivo} ({intensidad_str}) | ALUMNO: {nombre_mostrar.upper()}")

    c2 = head_tbl.cell(0,1)
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run(f"FECHA: {datetime.now().strftime('%d/%m/%Y')}\n").bold = True
    p2.add_run("Entrenamiento Funcional")
    doc.add_paragraph("_" * 95)

    # 1. GUÍA VISUAL
    doc.add_heading('1. Guía Visual de Ejercicios', level=2)
    num_ej = len(rutina_df)
    cols_visual = 4
    rows_visual = (num_ej + cols_visual - 1) // cols_visual
    
    vis_table = doc.add_table(rows=rows_visual, cols=cols_visual)
    vis_table.style = 'Table Grid'
    
    for i, row_data in enumerate(rutina_df.to_dict('records')):
        r = i // cols_visual
        c = i % cols_visual
        cell = vis_table.cell(r, c)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        ruta_img, msg = encontrar_imagen_recursiva(row_data['Imagen'])
        
        if ruta_img:
            try:
                run = p.add_run()
                run.add_picture(ruta_img, width=Inches(2.2), height=Inches(1.5))
                p.add_run("\n")
            except Exception:
                p.add_run(f"[Error Fichero]\n")
        else:
            p.add_run(f"\n[FALTA FOTO]\n")
            
        run_nom = p.add_run(row_data['Ejercicio'])
        run_nom.font.bold = True
        run_nom.font.size = Pt(10)

    doc.add_paragraph("\n")

    # 2. TABLA TÉCNICA
    doc.add_heading('2. Rutina Detallada', level=2)
    tech_table = doc.add_table(rows=1, cols=6)
    tech_table.style = 'Table Grid'
    
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
        # En notas ponemos el % usado
        row_cells[5].text = f"Intensidad: {row_data['Intensidad_Real']}" 

    doc.add_paragraph("\n")

    # 3. BORG
    doc.add_heading('3. Percepción de Esfuerzo (RPE)', level=3)
    borg_table = doc.add_table(rows=2, cols=5)
    borg_table.style = 'Table Grid'
    borg_table.autofit = True
    
    borg_data = [
        {"val": "6-8", "txt": "Muy Ligero", "icon": "🙂", "color": "A9DFBF"},
        {"val": "9-11", "txt": "Ligero", "icon": "😌", "color": "D4EFDF"},
        {"val": "12-14", "txt": "Algo Duro", "icon": "😐", "color": "F9E79F"},
        {"val": "15-17", "txt": "Duro", "icon": "😓", "color": "F5CBA7"},
        {"val": "18-20", "txt": "Máximo", "icon": "🥵", "color": "E6B0AA"}
    ]
    
    for i, data in enumerate(borg_data):
        c1 = borg_table.rows[0].cells[i]
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run1 = p1.add_run(f"{data['icon']}\n{data['val']}")
        run1.font.size = Pt(14)
        set_cell_bg_color(c1, data['color'])
        
        c2 = borg_table.rows[1].cells[i]
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.add_run(data['txt']).font.bold = True
        set_cell_bg_color(c2, data['color'])

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ STREAMLIT ---
st.title("🏋️ Generador Científico de Rutinas")

# Diagnóstico Sidebar
st.sidebar.markdown("### 📂 Estado")
imagenes_encontradas = []
for root, dirs, files in os.walk("."):
    for file in files:
        if file.lower().endswith(('.png', '.jpg', '.jpeg')):
            imagenes_encontradas.append(file)
if imagenes_encontradas:
    st.sidebar.success(f"✅ {len(imagenes_encontradas)} imágenes detectadas.")
else:
    st.sidebar.error("❌ No hay imágenes en GitHub.")

# Carga DB
if DB_EJERCICIOS is None:
    st.error("Error: DB_EJERCICIOS.xlsx no encontrado.")
    st.stop()
elif isinstance(DB_EJERCICIOS, str):
    st.error(DB_EJERCICIOS)
    st.stop()

# --- CONFIGURACIÓN DE LA SESIÓN (LÓGICA NUEVA) ---
col1, col2 = st.columns(2)
with col1:
    alumno = st.text_input("Nombre del Alumno:", "")
    tipos = sorted(list(set([e['tipo'] for e in DB_EJERCICIOS if e['tipo']])))
    sel_tipos = st.multiselect("Material:", options=tipos, default=tipos)

with col2:
    # SELECCIÓN DE OBJETIVO Y PARÁMETROS CIENTÍFICOS
    objetivo = st.selectbox("Objetivo:", ["Hipertrofia Muscular", "Definición Muscular", "Resistencia Muscular"])
    
    # Variables de salida de esta sección
    intensidad_seleccionada = 0  # Entero (ej: 85)
    reps_seleccionadas = ""      # String (ej: "6")
    descanso_seleccionado = ""   # String (ej: "3 min")
    
    if objetivo == "Hipertrofia Muscular":
        st.info("Rango: 1-6 Reps | Intensidad ≥ 85%")
        col_h1, col_h2 = st.columns(2)
        with col_h1:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [85, 90, 95, 100])
        with col_h2:
            descanso_seleccionado = st.selectbox("Descanso:", ["3 min", "4 min", "5 min"])
        
        # CÁLCULO AUTOMÁTICO DE REPS SEGÚN INTENSIDAD (Relación Inversa)
        if intensidad_seleccionada == 100:
            reps_seleccionadas = "1"
        elif intensidad_seleccionada == 95:
            reps_seleccionadas = "2"
        elif intensidad_seleccionada == 90:
            reps_seleccionadas = "3-4"
        else: # 85
            reps_seleccionadas = "5-6"
            
    elif objetivo == "Definición Muscular":
        st.info("Rango: 6-12 Reps | Intensidad 60-85%")
        col_d1, col_d2, col_d3 = st.columns(3)
        with col_d1:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [60, 65, 70, 75, 80, 85])
        with col_d2:
            val_reps = st.slider("Nº Repeticiones:", 6, 12, 10)
            reps_seleccionadas = str(val_reps)
        with col_d3:
            descanso_seleccionado = st.selectbox("Descanso:", ["1 min", "2 min", "3 min"])
            
    elif objetivo == "Resistencia Muscular":
        st.info("Rango: >13 Reps | Intensidad < 60%")
        col_r1, col_r2, col_r3 = st.columns(3)
        with col_r1:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [60, 55, 50, 45, 40])
        with col_r2:
            # Input numérico libre como pidió el usuario
            val_reps = st.number_input("Nº Repeticiones:", min_value=13, max_value=100, value=15, step=1)
            reps_seleccionadas = str(val_reps)
        with col_r3:
            # Generar lista de 60 a 0 bajando de 5 en 5
            opciones_segundos = [f"{s} seg" for s in range(60, -1, -5)]
            descanso_seleccionado = st.selectbox("Descanso:", opciones_segundos)

# Filtro de ejercicios
if sel_tipos:
    ej_filtrados = [e for e in DB_EJERCICIOS if e['tipo'] in sel_tipos]
    num_ej = st.slider("Cantidad de Ejercicios:", 1, min(10, len(ej_filtrados)), 6)
else:
    st.warning("Selecciona material.")
    st.stop()

# Selección
st.subheader("Selección de Ejercicios")
nombres_fil = [e['nombre'] for e in ej_filtrados]
seleccion = st.multiselect("Elige:", nombres_fil, max_selections=num_ej)

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

# Visor previo
st.markdown("---")
st.caption("Vista previa de imágenes:")
cols_prev = st.columns(6)
for i, item in enumerate(seleccionados_data):
    with cols_prev[i % 6]:
        ruta, msg = encontrar_imagen_recursiva(item['imagen'])
        if ruta:
            st.image(ruta, caption=item['nombre'], use_container_width=True)
        else:
            st.error(f"❌ {item['imagen']}")

# Inputs RM
st.subheader("Cargas de Entrenamiento")
st.write(f"Introduce el 1RM actual. Se calculará el **{intensidad_seleccionada}%** automáticamente.")
cols = st.columns(3)
rm_inputs = {}
for i, ej in enumerate(seleccionados_data):
    with cols[i%3]:
        rm_inputs[ej['nombre']] = st.number_input(f"1RM {ej['nombre']} (kg)", value=100, step=5)

# --- BOTONES DE ACCIÓN ---
col_gen, col_reset = st.columns([3, 1])

with col_gen:
    if st.button("📄 GENERAR DOCUMENTO CIENTÍFICO", type="primary", use_container_width=True):
        rutina_export = []
        
        for item in seleccionados_data:
            rm = rm_inputs[item['nombre']]
            
            # CÁLCULO CIENTÍFICO DE LA CARGA
            factor = intensidad_seleccionada / 100.0
            peso_real = int(rm * factor)
            
            rutina_export.append({
                "Ejercicio": item['nombre'],
                "Imagen": item['imagen'],
                "Reps": reps_seleccionadas,
                "Peso": peso_real,
                "Descanso": descanso_seleccionado,
                "Intensidad_Real": f"{intensidad_seleccionada}%"
            })
            
        df = pd.DataFrame(rutina_export)
        
        # TÍTULO DINÁMICO
        if len(sel_tipos) > 1:
            titulo_doc = "MIXTO"
        elif len(sel_tipos) == 1:
            titulo_doc = sel_tipos[0]
        else:
            titulo_doc = "GENERAL"
            
        docx = generar_word_final(df, objetivo, alumno, titulo_doc, f"{intensidad_seleccionada}%")
        
        st.success(f"Rutina generada: {objetivo} al {intensidad_seleccionada}%")
        st.download_button("📥 Descargar Rutina .docx", docx, f"Rutina_{alumno if alumno else 'Alumno'}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with col_reset:
    st.write("")
    if st.button("🔄 Reiniciar", use_container_width=True):
        st.rerun()
