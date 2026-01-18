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

# --- CARGAR EXCEL (CON CORRECCIÓN DE TILDES) ---
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
            
            # CORRECCIÓN AUTOMÁTICA DE "Olimpica" -> "Olímpica"
            df['tipo'] = df['tipo'].astype(str).str.replace('Olimpica', 'Olímpica', regex=False)
            df['tipo'] = df['tipo'].str.replace('olimpica', 'Olímpica', regex=False, case=False)
            df['tipo'] = df['tipo'].str.strip()
            
            df = df.fillna("")
            return df.to_dict('records')
        else:
            return None
    except Exception as e:
        return f"Error: {str(e)}"

DB_EJERCICIOS = cargar_ejercicios()

# --- GENERADOR WORD ---
def generar_word_final(rutina_df, lista_estiramientos, objetivo, alumno, titulo_material, intensidad_str, cardio_tipo, cardio_tiempo):
    doc = Document()
    
    # 1. Configuración Página A4 Horizontal
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)

    # --- PIE DE PÁGINA (FOOTER) ---
    footer = section.footer
    p_foot = footer.paragraphs[0]
    p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT 
    
    run_autor = p_foot.add_run("Programa creado por José Carlos Tejedor Lorenzo.            Página ")
    run_autor.font.size = Pt(10)
    
    run_num = p_foot.add_run()
    run_num.font.size = Pt(10)
    add_page_number(run_num)

    # ================= PÁGINA 1: PORTADA VISUAL =================

    # Encabezado Principal
    head_tbl = doc.add_table(rows=1, cols=2)
    head_tbl.autofit = False
    
    head_tbl.columns[0].width = Inches(9.4) 
    head_tbl.columns[1].width = Inches(1.2)
    
    c1 = head_tbl.cell(0,0)
    p = c1.paragraphs[0]
    
    # Título
    r1 = p.add_run(f"PROGRAMA DE ENTRENAMIENTO DE: {titulo_material.upper()}\n")
    r1.font.bold = True
    r1.font.size = Pt(14) 
    r1.font.color.rgb = RGBColor(41, 128, 185)
    
    nombre_mostrar = alumno if alumno.strip() else "ALUMNO"
    
    # Datos alumno
    r_obj_label = p.add_run("OBJETIVO: ")
    r_obj_label.font.bold = True
    p.add_run(f"{objetivo}")
    p.add_run("\t   ") 
    
    r_int_label = p.add_run("INTENSIDAD DE TRABAJO: ")
    r_int_label.font.bold = True
    p.add_run(f"({intensidad_str})")
    p.add_run("\t   ") 
    
    r_alu_label = p.add_run("ALUMNO/A: ")
    r_alu_label.font.bold = True
    p.add_run(f"{nombre_mostrar.upper()}")

    # Fecha
    c2 = head_tbl.cell(0,1)
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run(f"FECHA:\n{datetime.now().strftime('%d/%m/%Y')}").bold = True
    
    # Subtítulo
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.LEFT 
    run_sub = p_sub.add_run("Situación de Aprendizaje: Trabajo en Salas de Musculación 1º de Bachillerato IES Lucía de Medrano")
    run_sub.font.bold = True
    run_sub.font.name = 'Cambria'
    run_sub.font.size = Pt(16)    
    
    rPr = run_sub._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Cambria')
    rFonts.set(qn('w:hAnsi'), 'Cambria')
    rPr.append(rFonts)

    doc.add_paragraph("_" * 95)

    # Título Sección 1
    h1 = doc.add_heading(level=1)
    run_h1 = h1.add_run('1. Guía Visual de Ejercicios')
    run_h1.font.size = Pt(18)
    run_h1.font.color.rgb = RGBColor(44, 62, 80)

    # TABLA CARDIO
    cardio_table = doc.add_table(rows=1, cols=2)
    cardio_table.style = 'Table Grid'
    
    c_warm = cardio_table.cell(0,0)
    p_warm = c_warm.paragraphs[0]
    p_warm.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_w = p_warm.add_run("A) Calentamiento de 5 minutos de Duración")
    run_w.font.bold = True
    run_w.font.size = Pt(10)
    set_cell_bg_color(c_warm, "EAEDED") 
    
    c_card = cardio_table.cell(0,1)
    p_card = c_card.paragraphs[0]
    p_card.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_c = p_card.add_run(f"B) Cardio: {cardio_tipo} -> {cardio_tiempo}")
    run_c.font.bold = True
    run_c.font.size = Pt(10) 
    set_cell_bg_color(c_card, "EAEDED") 
    
    doc.add_paragraph("")

    # Grid de Imágenes (Rutina Principal)
    num_ej = len(rutina_df)
    cols_visual = 4
    rows_visual = (num_ej + cols_visual - 1) // cols_visual
    
    vis_table = doc.add_table(rows=rows_visual, cols=cols_visual)
    vis_table.style = 'Table Grid'
    
    # Altura forzada
    TR_HEIGHT_TWIPS = 2600 
    for row in vis_table.rows:
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        trHeight = OxmlElement('w:trHeight')
        trHeight.set(qn('w:val'), str(TR_HEIGHT_TWIPS))
        trHeight.set(qn('w:hRule'), "atLeast")
        trPr.append(trHeight)

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
                run.add_picture(ruta_img, width=Inches(2.0), height=Inches(1.35))
                p.add_run("\n")
            except Exception:
                p.add_run(f"[Error Fichero]\n")
        else:
            p.add_run(f"\n[FOTO NO DISPONIBLE]\n")
            
        run_nom = p.add_run(row_data['Ejercicio'])
        run_nom.font.bold = True
        run_nom.font.size = Pt(10)

    # SALTO DE PÁGINA
    doc.add_page_break()

    # ================= PÁGINA 2: RUTINA DETALLADA =================

    # Título Sección 2
    h2 = doc.add_heading(level=1)
    run_h2 = h2.add_run('2. Rutina Detallada')
    run_h2.font.size = Pt(18)
    run_h2.font.color.rgb = RGBColor(44, 62, 80)

    # Tabla Técnica
    tech_table = doc.add_table(rows=1, cols=6)
    tech_table.style = 'Table Grid'
    tech_table.autofit = False 
    
    widths = [0.7, 3.5, 1.5, 1.0, 1.5, 2.4] 
    headers = ["Orden", "Ejercicio", "Series x Reps", "Carga", "Descanso", "Notas"]
    
    row_hdr = tech_table.rows[0]
    for i, h in enumerate(headers):
        style_header_cell(row_hdr.cells[i], h, widths[i])
        
    for idx, row_data in rutina_df.iterrows():
        row_cells = tech_table.add_row().cells
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

    # ================= SECCIÓN 3: ESTIRAMIENTOS =================
    
    if lista_estiramientos:
        h3 = doc.add_heading(level=1)
        run_h3 = h3.add_run('3. Ejercicios de Estiramientos')
        run_h3.font.size = Pt(18)
        run_h3.font.color.rgb = RGBColor(44, 62, 80)

        num_est = len(lista_estiramientos)
        cols_est = 4
        rows_est = (num_est + cols_est - 1) // cols_est
        
        est_table = doc.add_table(rows=rows_est, cols=cols_est)
        est_table.style = 'Table Grid'
        
        for i, item_est in enumerate(lista_estiramientos):
            r = i // cols_est
            c = i % cols_est
            cell = est_table.cell(r, c)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            ruta_img, msg = encontrar_imagen_recursiva(item_est['imagen'])
            
            if ruta_img:
                try:
                    run = p.add_run()
                    run.add_picture(ruta_img, width=Inches(1.8), height=Inches(1.2))
                    p.add_run("\n")
                except Exception:
                    p.add_run(f"[Error Fichero]\n")
            else:
                p.add_run(f"\n[FOTO NO DISPONIBLE]\n")
                
            run_nom = p.add_run(item_est['nombre'])
            run_nom.font.bold = True
            run_nom.font.size = Pt(9)
        
        doc.add_paragraph("\n")

    # ================= SECCIÓN 4: BORG =================

    # Título Sección 4
    h4 = doc.add_heading(level=1)
    run_h4 = h4.add_run('4. Percepción del Esfuerzo (RPE)')
    run_h4.font.size = Pt(18)
    run_h4.font.color.rgb = RGBColor(44, 62, 80)

    # Tabla Borg
    borg_table = doc.add_table(rows=3, cols=5)
    borg_table.style = 'Table Grid'
    borg_table.autofit = True
    
    borg_data = [
        {"val": "6-8", "txt": "Muy Ligero", "icon": "🙂", "color": "A9DFBF"},
        {"val": "9-11", "txt": "Ligero", "icon": "😌", "color": "D4EFDF"},
        {"val": "12-14", "txt": "Algo Duro", "icon": "😐", "color": "F9E79F"},
        {"val": "15-17", "txt": "Duro", "icon": "😓", "color": "F5CBA7"},
        {"val": "18-20", "txt": "Máximo", "icon": "🥵", "color": "E6B0AA"}
    ]
    
    # Fila 1 (Iconos Grandes)
    row_icons = borg_table.rows[0]
    for i, data in enumerate(borg_data):
        c = row_icons.cells[i]
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        run_icon = p.add_run(f"{data['icon']}\n")
        run_icon.font.size = Pt(26) 
        
        run_val = p.add_run(f"{data['val']}")
        run_val.font.size = Pt(14)
        
        set_cell_bg_color(c, data['color'])

    # Fila 2 (Texto)
    row_text = borg_table.rows[1]
    for i, data in enumerate(borg_data):
        c = row_text.cells[i]
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(data['txt']).font.bold = True
        set_cell_bg_color(c, data['color'])

    # Fila 3 (Checks)
    row_check = borg_table.rows[2]
    tr = row_check._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), "600")
    trPr.append(trHeight)

    for i, data in enumerate(borg_data):
        c = row_check.cells[i]
        set_cell_bg_color(c, data['color'])

    p_note = doc.add_paragraph("Marca con una X la sensación global al terminar el entrenamiento.")
    p_note.style = "Caption"

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ STREAMLIT ---

# Título Principal
st.markdown("""
<style>
.big-font {
    font-size:30px !important;
    font-weight: bold;
}
.sub-font {
    font-size:20px !important;
    font-style: italic;
    color: #555;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="big-font">Generador Científico de Rutinas creado por José Carlos Tejedor Lorenzo</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-font">Situación de aprendizaje: Trabajo en Salas de Musculación 1º de Bachillerato IES Lucía de Medrano.</p>', unsafe_allow_html=True)
st.markdown("---")

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
    
    # 1. OBTENER TIPOS
    tipos_todos = sorted(list(set([e['tipo'] for e in DB_EJERCICIOS if e['tipo']])))
    # 2. FILTRAR ESTIRAMIENTOS PARA QUE NO SALGAN EN MATERIAL
    tipos_entreno = [t for t in tipos_todos if 'estiramiento' not in t.lower()]
    
    sel_tipos = st.multiselect("Material de Entrenamiento:", options=tipos_entreno, default=tipos_entreno)
    
    cardio_seleccion = st.selectbox(
        "Todos los días cardio:", 
        ["Bicicleta", "Cinta de Correr", "Step", "Remo de cardio"]
    )

with col2:
    objetivo = st.selectbox("Objetivo:", ["Hipertrofia Muscular", "Definición Muscular", "Resistencia Muscular"])
    
    intensidad_seleccionada = 0
    reps_seleccionadas = ""
    descanso_seleccionado = ""
    cardio_duracion = "" 
    
    if objetivo == "Hipertrofia Muscular":
        st.info("Rango: 1-6 Reps | Intensidad ≥ 85%")
        cardio_duracion = "10-15 min de cardio" 
        
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
        cardio_duracion = "50-55 min de cardio" 
        
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
        cardio_duracion = "Más de 30 min de cardio" 
        
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

# --- SELECCIÓN DE EJERCICIOS ---
st.subheader("Selección de Ejercicios")

with st.expander(f"📸 Ver Galería Visual de ejercicios disponibles ({', '.join(sel_tipos)})"):
    cols_galeria = st.columns(6)
    for i, ej in enumerate(ej_filtrados):
        with cols_galeria[i % 6]:
            ruta, msg = encontrar_imagen_recursiva(ej['imagen'])
            if ruta:
                st.image(ruta, caption=ej['nombre'], use_container_width=True)
            else:
                st.caption(f"❌ {ej['nombre']}")

nombres_fil = [e['nombre'] for e in ej_filtrados]
seleccion = st.multiselect("Elige los ejercicios:", nombres_fil, max_selections=num_ej)

# --- CHECKBOX PARA EVITAR RELLENO AUTOMÁTICO ---
rellenar_auto = st.checkbox(f"Rellenar automáticamente hasta llegar a {num_ej} ejercicios (si no seleccionas suficientes)", value=True)

seleccionados_data = []
nombres_finales = seleccion.copy()

# Lógica de relleno (Solo si el checkbox está activo)
if rellenar_auto and len(nombres_finales) < num_ej:
    pool = [x for x in ej_filtrados if x['nombre'] not in nombres_finales]
    needed = num_ej - len(nombres_finales)
    if needed <= len(pool):
        extras = random.sample(pool, needed)
        nombres_finales.extend([x['nombre'] for x in extras])

# Reconstruir la lista de objetos completa basada en nombres_finales
seleccionados_data = []
for nom in nombres_finales:
    obj_ejercicio = next((x for x in ej_filtrados if x['nombre'] == nom), None)
    if obj_ejercicio:
        seleccionados_data.append(obj_ejercicio)

st.markdown("---")
# Cambio de texto para que tenga sentido con o sin relleno
if rellenar_auto:
    st.caption("Has seleccionado (o se ha completado automáticamente):")
else:
    st.caption("Has seleccionado estrictamente:")
    
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

# --- NUEVA SECCIÓN: ESTIRAMIENTOS ---
st.markdown("---")
st.subheader("Vuelta a la Calma: Estiramientos")

pool_estiramientos = [e for e in DB_EJERCICIOS if 'estiramiento' in str(e['tipo']).lower()]
nombres_est = [e['nombre'] for e in pool_estiramientos]

if pool_estiramientos:
    with st.expander("🧘 Ver Galería Visual de Estiramientos disponibles"):
        cols_est_gal = st.columns(6)
        for i, ej in enumerate(pool_estiramientos):
            with cols_est_gal[i % 6]:
                ruta, msg = encontrar_imagen_recursiva(ej['imagen'])
                if ruta:
                    st.image(ruta, caption=ej['nombre'], use_container_width=True)
                else:
                    st.caption(f"❌ {ej['nombre']}")

    num_est_select = st.slider("Cantidad de estiramientos:", 1, 8, 4)
    seleccion_est = st.multiselect("Elige estiramientos:", nombres_est, max_selections=num_est_select)
    
    # Relleno automático también para estiramientos (opcional, pero consistente)
    estiramientos_finales_nombres = seleccion_est.copy()
    if len(estiramientos_finales_nombres) < num_est_select:
        pool_est = [x['nombre'] for x in pool_estiramientos if x['nombre'] not in estiramientos_finales_nombres]
        needed_est = num_est_select - len(estiramientos_finales_nombres)
        if needed_est <= len(pool_est):
             estiramientos_finales_nombres.extend(random.sample(pool_est, needed_est))

    estiramientos_finales = []
    for nom in estiramientos_finales_nombres:
         estiramientos_finales.append(next(x for x in pool_estiramientos if x['nombre'] == nom))

else:
    st.warning("⚠️ No se han encontrado ejercicios marcados como 'Estiramientos' en el Excel.")
    estiramientos_finales = []


# --- BOTONES DE ACCIÓN ---
col_gen, col_reset = st.columns([3, 1])

with col_gen:
    st.write("")
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
        
        docx = generar_word_final(
            rutina_df=df, 
            lista_estiramientos=estiramientos_finales, 
            objetivo=objetivo, 
            alumno=alumno, 
            titulo_material=titulo_doc, 
            intensidad_str=f"{intensidad_seleccionada}%", 
            cardio_tipo=cardio_seleccion, 
            cardio_tiempo=cardio_duracion
        )
        
        st.success(f"Rutina generada: {objetivo} ({reps_seleccionadas} reps al {intensidad_seleccionada}%) + {len(estiramientos_finales)} Estiramientos")
        st.download_button("📥 Descargar Rutina .docx", docx, f"Rutina_{alumno if alumno else 'Alumno'}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with col_reset:
    st.write("")
    if st.button("🔄 Reiniciar", use_container_width=True):
        st.rerun()
