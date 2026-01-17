import streamlit as st
import pandas as pd
import random
import os
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement, ns
from docx.oxml.ns import qn
from io import BytesIO
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Entrenador Pro Científico", layout="wide")

# --- FUNCIONES AUXILIARES PARA WORD ---

def create_element(name):
    return OxmlElement(name)

def create_attribute(element, name, value):
    element.set(ns.qn(name), value)

def add_page_number(run):
    """Agrega el campo dinámico de número de página"""
    fldChar1 = create_element('w:fldChar')
    create_attribute(fldChar1, 'w:fldCharType', 'begin')

    instrText = create_element('w:instrText')
    create_attribute(instrText, 'xml:space', 'preserve')
    instrText.text = "PAGE"

    fldChar2 = create_element('w:fldChar')
    create_attribute(fldChar2, 'w:fldCharType', 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

def set_cell_bg_color(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def style_header_cell(cell, text, width_inches=None):
    cell.text = text
    if width_inches:
        cell.width = Inches(width_inches)
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

# --- GENERADOR WORD (RENOVADO) ---
def generar_word_final(rutina_df, objetivo, alumno, titulo_material, intensidad_str):
    doc = Document()
    
    # 1. Configuración Página A4 Horizontal
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)

    # --- PIE DE PÁGINA (FOOTER) ---
    footer = section.footer
    p_foot = footer.paragraphs[0]
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_foot = p_foot.add_run("Programa creado por Jose Carlos Tejedor. Página ")
    run_foot.font.size = Pt(9)
    add_page_number(run_foot) # Inserta número de página automático

    # ================= PÁGINA 1: PORTADA VISUAL =================

    # Encabezado Principal
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
    p.add_run(f"OBJETIVO: {objetivo} ({intensidad_str}) | ALUMNO: {nombre_mostrar.upper()}")

    c2 = head_tbl.cell(0,1)
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run(f"FECHA: {datetime.now().strftime('%d/%m/%Y')}\n").bold = True
    
    # Subtítulo Situación de Aprendizaje
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run("Situación de Aprendizaje: Trabajo en Salas de Musculación 1º de Bachillerato IES Lucía de Medrano")
    run_sub.font.bold = True
    run_sub.font.italic = True
    run_sub.font.size = Pt(11)

    doc.add_paragraph("_" * 95)

    # Título Sección 1 (Grande)
    h1 = doc.add_heading(level=1)
    run_h1 = h1.add_run('1. Guía Visual de Ejercicios')
    run_h1.font.size = Pt(18) # Letra más grande
    run_h1.font.color.rgb = RGBColor(44, 62, 80)

    # Grid de Imágenes
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
            p.add_run(f"\n[FOTO NO DISPONIBLE]\n")
            
        run_nom = p.add_run(row_data['Ejercicio'])
        run_nom.font.bold = True
        run_nom.font.size = Pt(10)

    # SALTO DE PÁGINA OBLIGATORIO
    doc.add_page_break()

    # ================= PÁGINA 2: RUTINA DETALLADA =================

    # Título Sección 2 (Grande)
    h2 = doc.add_heading(level=1)
    run_h2 = h2.add_run('2. Rutina Detallada')
    run_h2.font.size = Pt(18)
    run_h2.font.color.rgb = RGBColor(44, 62, 80)

    # Tabla Técnica con anchos personalizados
    tech_table = doc.add_table(rows=1, cols=6)
    tech_table.style = 'Table Grid'
    tech_table.autofit = False # Importante para que respete los anchos
    
    # Definición de anchos (Suma total aprox 10.6 pulgadas)
    widths = [0.7, 3.5, 1.5, 1.0, 1.5, 2.4] 
    headers = ["Orden", "Ejercicio", "Series x Reps", "Carga", "Descanso", "Notas"]
    
    # Configurar cabecera
    row_hdr = tech_table.rows[0]
    for i, h in enumerate(headers):
        style_header_cell(row_hdr.cells[i], h, widths[i])
        
    # Rellenar filas
    for idx, row_data in rutina_df.iterrows():
        row_cells = tech_table.add_row().cells
        
        # Aplicar anchos a las celdas de datos también
        for i in range(6):
            row_cells[i].width = Inches(widths[i])

        row_cells[0].text = str(idx + 1)
        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        row_cells[1].text = row_data['Ejercicio']
        row_cells[2].text = f"4 x {row_data['Reps']}"
        row_cells[3].text = f"{row_data['Peso']} kg"
        row_cells[4].text = row_data['Descanso']
        row_cells[5].text = f"Int: {row_data['Intensidad_Real']}" 

    doc.add_paragraph("\n")

    # Título Sección 3 (Grande)
    h3 = doc.add_heading(level=1)
    run_h3 = h3.add_run('3. Percepción del Esfuerzo (RPE)')
    run_h3.font.size = Pt(18)
    run_h3.font.color.rgb = RGBColor(44, 62, 80)

    # Tabla Borg
    borg_table = doc.add_table(rows=3, cols=5) # 3 filas: Iconos, Texto, Casillas vacías
    borg_table.style = 'Table Grid'
    borg_table.autofit = True
    
    borg_data = [
        {"val": "6-8", "txt": "Muy Ligero", "icon": "🙂", "color": "A9DFBF"},
        {"val": "9-11", "txt": "Ligero", "icon": "😌", "color": "D4EFDF"},
        {"val": "12-14", "txt": "Algo Duro", "icon": "😐", "color": "F9E79F"},
        {"val": "15-17", "txt": "Duro", "icon": "😓", "color": "F5CBA7"},
        {"val": "18-20", "txt": "Máximo", "icon": "🥵", "color": "E6B0AA"}
    ]
    
    # Fila 1: Iconos y números
    row_icons = borg_table.rows[0]
    for i, data in enumerate(borg_data):
        c = row_icons.cells[i]
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"{data['icon']}\n{data['val']}")
        run.font.size = Pt(14)
        set_cell_bg_color(c, data['color'])

    # Fila 2: Texto descriptivo
    row_text = borg_table.rows[1]
    for i, data in enumerate(borg_data):
        c = row_text.cells[i]
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(data['txt']).font.bold = True
        set_cell_bg_color(c, data['color'])

    # Fila 3: Casillas para marcar con X (Manteniendo color)
    row_check = borg_table.rows[2]
    # Hacer esta fila un poco más alta
    tr = row_check._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), "600") # Altura en twips
    trPr.append(trHeight)

    for i, data in enumerate(borg_data):
        c = row_check.cells[i]
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Dejamos vacío o ponemos un guion bajo sutil
        set_cell_bg_color(c, data['color'])

    # Añadir instrucción pequeña debajo
    p_note = doc.add_paragraph("Marca con una X la sensación global al terminar el entrenamiento.")
    p_note.style = "Caption"

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

# --- CONFIGURACIÓN DE LA SESIÓN ---
col1, col2 = st.columns(2)
with col1:
    alumno = st.text_input("Nombre del Alumno:", "")
    tipos = sorted(list(set([e['tipo'] for e in DB_EJERCICIOS if e['tipo']])))
    sel_tipos = st.multiselect("Material:", options=tipos, default=tipos)

with col2:
    objetivo = st.selectbox("Objetivo:", ["Hipertrofia Muscular", "Definición Muscular", "Resistencia Muscular"])
    
    intensidad_seleccionada = 0
    reps_seleccionadas = ""
    descanso_seleccionado = ""
    
    if objetivo == "Hipertrofia Muscular":
        st.info("Rango: 1-6 Reps | Intensidad ≥ 85%")
        col_h1, col_h2, col_h3 = st.columns(3)
        with col_h1:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [85, 90, 95, 100])
        with col_h2:
            val_reps = st.selectbox("Repeticiones:", [1, 2, 3, 4, 5, 6])
            reps_seleccionadas = str(val_reps)
        with col_h3:
            descanso_seleccionado = st.selectbox("Descanso:", ["3 min", "4 min", "5 min"])
            
    elif objetivo == "Definición Muscular":
        st.info("Rango: 6-12 Reps | Intensidad 60-85%")
        col_d1, col_d2, col_d3 = st.columns(3)
        with col_d1:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [60, 65, 70, 75, 80, 85])
        with col_d2:
            val_reps = st.selectbox("Repeticiones:", [6, 7, 8, 9, 10, 11, 12])
            reps_seleccionadas = str(val_reps)
        with col_d3:
            descanso_seleccionado = st.selectbox("Descanso:", ["1 min", "2 min", "3 min"])
            
    elif objetivo == "Resistencia Muscular":
        st.info("Rango: 13-20 Reps | Intensidad < 60%")
        col_r1, col_r2, col_r3 = st.columns(3)
        with col_r1:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [60, 55, 50, 45, 40])
        with col_r2:
            val_reps = st.selectbox("Repeticiones:", [13, 14, 15, 16, 17, 18, 19, 20])
            reps_seleccionadas = str(val_reps)
        with col_r3:
            opciones_segundos = [f"{s} seg" for s in range(60, -1, -5)]
            descanso_seleccionado = st.selectbox("Descanso:", opciones_segundos)

if sel_tipos:
    ej_filtrados = [e for e in DB_EJERCICIOS if e['tipo'] in sel_tipos]
    num_ej = st.slider("Cantidad de Ejercicios:", 1, min(10, len(ej_filtrados)), 6)
else:
    st.warning("Selecciona material.")
    st.stop()

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

st.subheader("Cargas de Entrenamiento")
st.write(f"Introduce el 1RM actual. Se calculará el **{intensidad_seleccionada}%** automáticamente.")
cols = st.columns(3)
rm_inputs = {}
for i, ej in enumerate(seleccionados_data):
    with cols[i%3]:
        rm_inputs[ej['nombre']] = st.number_input(f"1RM {ej['nombre']} (kg)", value=100, step=5)

col_gen, col_reset = st.columns([3, 1])

with col_gen:
    if st.button("📄 GENERAR DOCUMENTO CIENTÍFICO", type="primary", use_container_width=True):
        rutina_export = []
        
        for item in seleccionados_data:
            rm = rm_inputs[item['nombre']]
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
        
        if len(sel_tipos) > 1:
            titulo_doc = "MIXTO"
        elif len(sel_tipos) == 1:
            titulo_doc = sel_tipos[0]
        else:
            titulo_doc = "GENERAL"
            
        docx = generar_word_final(df, objetivo, alumno, titulo_doc, f"{intensidad_seleccionada}%")
        
        st.success(f"Rutina generada: {objetivo} ({reps_seleccionadas} reps al {intensidad_seleccionada}%)")
        st.download_button("📥 Descargar Rutina .docx", docx, f"Rutina_{alumno if alumno else 'Alumno'}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with col_reset:
    st.write("")
    if st.button("🔄 Reiniciar", use_container_width=True):
        st.rerun()
