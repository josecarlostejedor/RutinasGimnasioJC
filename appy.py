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
st.set_page_config(page_title="Generador Rutinas V4 (Final)", layout="wide")

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

# --- BUSCADOR "RASTREADOR" DE IMÁGENES (NUEVO) ---
def encontrar_imagen_recursiva(nombre_objetivo):
    """
    Busca una imagen en TODO el directorio actual y subcarpetas.
    Coincide aunque en Excel no pongas la extensión.
    Ej: Excel 'press' -> Encuentra 'images/Press_Banca.jpg' si contiene 'press'
    """
    if not nombre_objetivo or pd.isna(nombre_objetivo):
        return None, "Celda Vacía"

    nombre_limpio = str(nombre_objetivo).strip().lower()
    # Quitamos extensión del nombre del Excel si la tuviera para comparar solo nombres
    nombre_base_excel = os.path.splitext(nombre_limpio)[0]

    # Recorrer todos los archivos del servidor
    for root, dirs, files in os.walk("."):
        for filename in files:
            # Filtrar solo imagenes para no confundir con otros archivos
            if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                filename_base = os.path.splitext(filename)[0].lower()
                
                # 1. Coincidencia Exacta (con extension)
                if filename.lower() == nombre_limpio:
                    return os.path.join(root, filename), "Exacta"
                
                # 2. Coincidencia de Nombre (sin extension)
                # Si Excel dice "banca" y el archivo es "banca.jpg"
                if filename_base == nombre_base_excel:
                    return os.path.join(root, filename), "Por Nombre"

    return None, f"No encontrado: {nombre_limpio}"

# --- CARGAR EXCEL ---
@st.cache_data
def cargar_ejercicios():
    try:
        if os.path.exists("DB_EJERCICIOS.xlsx"):
            df = pd.read_excel("DB_EJERCICIOS.xlsx")
            df.columns = df.columns.str.strip().str.lower()
            
            if 'nombre' not in df.columns:
                if 'ejercicio' in df.columns: df.rename(columns={'ejercicio': 'nombre'}, inplace=True)
            
            # Asegurar columnas
            for col in ['tipo', 'imagen', 'desc']:
                if col not in df.columns: df[col] = ""
            
            df = df.fillna("")
            return df.to_dict('records')
        else:
            return None
    except Exception as e:
        return f"Error: {str(e)}"

DB_EJERCICIOS = cargar_ejercicios()

# --- LÓGICA PARAMETROS ---
def obtener_parametros(objetivo):
    if objetivo == "Hipertrofia": return {"reps": "8-12", "descanso": "90 seg"}
    elif objetivo == "Fuerza Máxima": return {"reps": "3-5", "descanso": "3-5 min"}
    elif objetivo == "Resistencia": return {"reps": "15-20", "descanso": "45 seg"}

# --- GENERADOR WORD ---
def generar_word_final(rutina_df, objetivo, alumno, tipo_rutina):
    doc = Document()
    
    # Configuración Página Horizontal
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
    r1 = p.add_run(f"PROGRAMA: {tipo_rutina.upper()}\n")
    r1.font.bold = True
    r1.font.size = Pt(16)
    r1.font.color.rgb = RGBColor(41, 128, 185)
    p.add_run(f"OBJETIVO: {objetivo} | ALUMNO: {alumno}")

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
        
        # --- BÚSQUEDA ROBUSTA DE IMAGEN ---
        ruta_img, msg = encontrar_imagen_recursiva(row_data['Imagen'])
        
        if ruta_img:
            try:
                run = p.add_run()
                # Forzamos tamaño fijo para que no rompa la tabla
                run.add_picture(ruta_img, width=Inches(2.2), height=Inches(1.5))
                p.add_run("\n")
            except Exception as e:
                p.add_run(f"[Error Fichero: {msg}]\n")
        else:
            # Si falla, mostramos qué nombre buscó para depurar
            p.add_run(f"\n[FALTA: {row_data['Imagen']}]\n")
            
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
        row_cells[5].text = ""

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
st.title("🏋️ Generador Final con Diagnóstico de Imágenes")

# --- DIAGNÓSTICO EN BARRA LATERAL (CLAVE PARA SOLUCIONAR EL PROBLEMA) ---
st.sidebar.markdown("### 📂 Archivos en el Servidor")
imagenes_encontradas = []
for root, dirs, files in os.walk("."):
    for file in files:
        if file.lower().endswith(('.png', '.jpg', '.jpeg')):
            imagenes_encontradas.append(file)

if len(imagenes_encontradas) > 0:
    st.sidebar.success(f"✅ Se han detectado {len(imagenes_encontradas)} imágenes.")
    with st.sidebar.expander("Ver lista de imágenes"):
        st.write(imagenes_encontradas)
else:
    st.sidebar.error("❌ NO SE DETECTAN IMÁGENES. Revisa que las hayas subido a GitHub.")

# --- CUERPO PRINCIPAL ---
if DB_EJERCICIOS is None:
    st.error("Sube DB_EJERCICIOS.xlsx")
    st.stop()
elif isinstance(DB_EJERCICIOS, str):
    st.error(DB_EJERCICIOS)
    st.stop()

col1, col2 = st.columns(2)
with col1:
    alumno = st.text_input("Alumno:", "Atleta")
    tipos = sorted(list(set([e['tipo'] for e in DB_EJERCICIOS if e['tipo']])))
    sel_tipos = st.multiselect("Material:", options=tipos, default=tipos)
with col2:
    objetivo = st.selectbox("Objetivo", ["Hipertrofia", "Fuerza Máxima", "Resistencia"])
    if sel_tipos:
        ej_filtrados = [e for e in DB_EJERCICIOS if e['tipo'] in sel_tipos]
        num_ej = st.slider("Nº Ejercicios", 1, min(10, len(ej_filtrados)), 6)
    else:
        st.stop()

# Selección
st.subheader(f"Selección de ejercicios")
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

# --- VISOR PREVIO PARA QUE NO FALLES ---
st.info("👇 **Comprobación antes de generar Word:**")
cols_prev = st.columns(6)
for i, item in enumerate(seleccionados_data):
    with cols_prev[i % 6]:
        ruta, msg = encontrar_imagen_recursiva(item['imagen'])
        if ruta:
            st.image(ruta, caption=item['nombre'], use_container_width=True)
        else:
            st.error(f"❌ Falta: {item['imagen']}")

# Inputs RM
st.markdown("---")
cols = st.columns(3)
rm_inputs = {}
for i, ej in enumerate(seleccionados_data):
    with cols[i%3]:
        rm_inputs[ej['nombre']] = st.number_input(f"RM {ej['nombre']}", value=50, step=5)

if st.button("Generar Word"):
    params = obtener_parametros(objetivo)
    rutina_export = []
    
    for item in seleccionados_data:
        rm = rm_inputs[item['nombre']]
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
    docx = generar_word_final(df, objetivo, alumno, " + ".join(sel_tipos))
    
    st.download_button("📥 Descargar Word", docx, f"Rutina_{alumno}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
