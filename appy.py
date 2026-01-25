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

# --- GESTIÓN DE ESTADO ---
if 'reset_counter' not in st.session_state:
    st.session_state.reset_counter = 0

def get_key(base_name):
    return f"{base_name}_{st.session_state.reset_counter}"

# --- DATOS TEÓRICOS DE LOS OBJETIVOS ---
INFO_OBJETIVOS = {
    "Fuerza Máxima": """1️⃣ FUERZA MÁXIMA
🎯 Objetivo
Aumentar la capacidad máxima de producción de fuerza (adaptación neural).

🏋️‍♂️ Trabajo de fuerza
Intensidad: 85–100 % RM
Repeticiones: 1–5
Series: 4–6
Descanso: 3–6 min
Ejercicios: multiarticulares (sentadilla, peso muerto, press banca, press militar)

❤️ Trabajo cardiovascular
Tipo: aeróbico extensivo
Intensidad: 60–70 % FCmáx
Duración: 15–25 min
Frecuencia: 1–2 días/semana
Objetivo: recuperación, no interferir con la fuerza""",

    "Hipertrofia Muscular": """2️⃣ HIPERTROFIA MUSCULAR
🎯 Objetivo
Aumentar el tamaño muscular (hipertrofia miofibrilar y sarcoplasmática).

🏋️‍♂️ Trabajo de fuerza
Intensidad: 65–85 % RM
Repeticiones: 6–12
Series: 3–6
Descanso: 60–120 s
RIR: 0–2 (cerca del fallo)

❤️ Trabajo cardiovascular
Tipo: aeróbico moderado
Intensidad: 65–75 % FCmáx
Duración: 20–30 min
Frecuencia: 2–3 días/semana
Objetivo: salud cardiovascular sin comprometer ganancias musculares""",

    "Definición Muscular": """3️⃣ DEFINICIÓN MUSCULAR
🎯 Objetivo
Mantener masa muscular + reducir grasa corporal.

🏋️‍♂️ Trabajo de fuerza
Intensidad: 60–75 % RM
Repeticiones: 10–15
Series: 3–5
Descanso: 30–60 s
Métodos: superseries, circuitos, alta densidad

❤️ Trabajo cardiovascular
Tipo: HIIT + aeróbico
HIIT: 85–95 % FCmáx | 10–20 min | 1–2 días/sem
Aeróbico: 65–75 % FCmáx | 30–45 min | 2–3 días/sem""",

    "Resistencia Muscular": """4️⃣ RESISTENCIA MUSCULAR
🎯 Objetivo
Mejorar la capacidad de sostener esfuerzos prolongados.

🏋️‍♂️ Trabajo de fuerza
Intensidad: 30–60 % RM
Repeticiones: 15–30+
Series: 2–4
Descanso: 15–45 s
Formato: circuitos o estaciones

❤️ Trabajo cardiovascular
Tipo: aeróbico extensivo
Intensidad: 65–80 % FCmáx
Duración: 30–60 min
Frecuencia: 3–5 días/semana
Objetivo: base aeróbica y resistencia general""",

    "Mantenimiento Muscular": """5️⃣ MANTENIMIENTO MUSCULAR
🎯 Objetivo
Conservar masa muscular, fuerza y salud con bajo volumen.

🏋️‍♂️ Trabajo de fuerza
Intensidad: 60–75 % RM
Repeticiones: 8–12
Series: 2–3
Descanso: 60–90 s
Frecuencia: 2–3 días/semana

❤️ Trabajo cardiovascular
Tipo: aeróbico saludable
Intensidad: 60–75 % FCmáx
Duración: 20–40 min
Frecuencia: 2–4 días/semana""",

    "Rehabilitación Muscular y Articular": """5️⃣ REHABILITACIÓN MUSCULAR Y ARTICULAR
🔁 Progresión recomendada (por fases)

🟢 Fase 1 – Readaptación
20–30 % RM
Isométricos + movilidad
Cardio muy suave

🟡 Fase 2 – Reacondicionamiento
30–50 % RM
Concéntrico + excéntrico lento
Propiocepción dinámica

🔵 Fase 3 – Transición al entrenamiento
50–60 % RM
Patrones básicos
Integración progresiva con mantenimiento muscular""",

    "Programa de Pérdida de Peso": """🔥 PROGRAMA DE PÉRDIDA DE PESO

🎯 Objetivo
Reducir grasa corporal
Mantener o minimizar la pérdida de masa muscular
Aumentar el gasto energético total
Mejorar la salud metabólica y cardiovascular
Crear hábitos de actividad física sostenibles

🏋️‍♂️ Fuerza (entrenamiento principal)
🔹 Intensidad
50–70 % de 1RM
🔹 Repeticiones
12–20 repeticiones
🔹 Series
3–4 series
🔹 Descanso
20–45 segundos

🔹 Organización del trabajo
Circuitos
Superseries
Ejercicios multiarticulares prioritarios
Ritmo continuo, intensidad alta

Objetivo de la fuerza
Mantener masa muscular
Aumentar gasto calórico
Mejorar tono muscular

❤️ Entrenamiento cardiovascular
🔹 Aeróbico continuo
Intensidad: 60–75 % FCmáx
Duración: 30–60 min
Frecuencia: 3–5 días/semana
Ejemplos: caminar rápido, bici, elíptica, natación

🔹 HIIT (opcional)
Intensidad: 85–95 % FCmáx
Duración: 10–20 min
Frecuencia: 1–2 días/semana
Formato: intervalos cortos de alta intensidad + recuperación activa

🧠 Consejos clave
La fuerza es imprescindible para no perder músculo
No bajar de 50 % RM de forma sistemática
Mantener déficit calórico moderado
Priorizar adherencia y progresión
Dormir y recuperarse adecuadamente
Aumentar el NEAT- Non Exercice Activity Thermogenesis (pasos diarios, vida activa)
Revaluar cargas cada 4–6 semanas

⚠️ Errores comunes
Solo cardio y nada de fuerza
Usar cargas muy ligeras durante meses
Descansos excesivos
Déficits calóricos extremos"""
}

# --- FUNCIONES AUXILIARES PARA WORD ---
def create_element(name):
    return OxmlElement(name)

def create_attribute(element, name, value):
    element.set(ns.qn(name), value)

def add_page_number(run):
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

def set_row_cant_split(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    cantSplit = OxmlElement('w:cantSplit')
    trPr.append(cantSplit)

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

# --- CARGAR EXCEL (CON SOPORTE PARA ESTABILIZADORES) ---
@st.cache_data
def cargar_ejercicios():
    try:
        if os.path.exists("DB_EJERCICIOS.xlsx"):
            df = pd.read_excel("DB_EJERCICIOS.xlsx")
            df.columns = df.columns.str.strip().str.lower()
            if 'nombre' not in df.columns:
                if 'ejercicio' in df.columns: df.rename(columns={'ejercicio': 'nombre'}, inplace=True)
            
            # Columnas obligatorias básicas
            for col in ['tipo', 'imagen', 'desc']:
                if col not in df.columns: df[col] = ""
                
            # Columnas para análisis muscular (AHORA INCLUYE ESTABILIZADORES)
            for col in ['agonistas', 'sinergistas', 'estabilizadores']:
                if col not in df.columns: df[col] = ""
            
            # Normalización de texto
            df['tipo'] = df['tipo'].astype(str).str.replace('Olimpica', 'Olímpica', regex=False)
            df['tipo'] = df['tipo'].str.replace('olimpica', 'Olímpica', regex=False, case=False)
            df['tipo'] = df['tipo'].str.replace('Rehabilitacion', 'Rehabilitación', regex=False)
            df['tipo'] = df['tipo'].str.replace('Rotualiana', 'Rotuliana', regex=False)
            df['tipo'] = df['tipo'].str.strip()
            
            # Rellenar nulos
            df = df.fillna("")
            return df.to_dict('records')
        else:
            return None
    except Exception as e:
        return f"Error: {str(e)}"

DB_EJERCICIOS = cargar_ejercicios()

# --- GENERADOR WORD (CON ESTABILIZADORES) ---
def generar_word_final(rutina_df, lista_estiramientos, objetivo, alumno, titulo_material, intensidad_str, cardio_tipo, cardio_tiempo, series_str, incluir_analisis_muscular):
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)

    # Footer
    footer = section.footer
    p_foot = footer.paragraphs[0]
    p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT 
    run_autor = p_foot.add_run("Programa creado por José Carlos Tejedor Lorenzo.            Página ")
    run_autor.font.size = Pt(10)
    run_num = p_foot.add_run()
    run_num.font.size = Pt(10)
    add_page_number(run_num)

    # PÁGINA 1
    head_tbl = doc.add_table(rows=1, cols=2)
    head_tbl.autofit = False
    head_tbl.columns[0].width = Inches(9.8) 
    head_tbl.columns[1].width = Inches(1.0)
    
    c1 = head_tbl.cell(0,0)
    p = c1.paragraphs[0]
    r1 = p.add_run(f"PROGRAMA DE ENTRENAMIENTO DE: {titulo_material.upper()}\n")
    r1.font.bold = True
    r1.font.size = Pt(12) 
    r1.font.color.rgb = RGBColor(41, 128, 185)
    
    nombre_mostrar = alumno if alumno.strip() else "ALUMNO"
    
    font_size_meta = Pt(10) 
    r_obj_label = p.add_run("OBJETIVO: ")
    r_obj_label.font.bold = True
    r_obj_label.font.size = font_size_meta
    r_obj_val = p.add_run(f"{objetivo}")
    r_obj_val.font.size = font_size_meta
    
    p.add_run("   |   ").font.size = font_size_meta 
    
    r_int_label = p.add_run("INTENSIDAD DE TRABAJO: ")
    r_int_label.font.bold = True
    r_int_label.font.size = font_size_meta
    r_int_val = p.add_run(f"({intensidad_str})")
    r_int_val.font.size = font_size_meta
    
    p.add_run("   |   ").font.size = font_size_meta 
    
    r_alu_label = p.add_run("ALUMNO/A: ")
    r_alu_label.font.bold = True
    r_alu_label.font.size = font_size_meta
    r_alu_val = p.add_run(f"{nombre_mostrar.upper()}")
    r_alu_val.font.size = font_size_meta

    c2 = head_tbl.cell(0,1)
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.add_run(f"FECHA:\n{datetime.now().strftime('%d/%m/%Y')}").bold = True
    
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

    h1 = doc.add_heading(level=1)
    titulo_seccion_1 = '1. Guía Visual de Ejercicios con Análisis Muscular' if incluir_analisis_muscular else '1. Guía Visual de Ejercicios'
    run_h1 = h1.add_run(titulo_seccion_1)
    run_h1.font.size = Pt(18)
    run_h1.font.color.rgb = RGBColor(44, 62, 80)

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

    num_ej = len(rutina_df)
    cols_visual = 4
    rows_visual = (num_ej + cols_visual - 1) // cols_visual
    vis_table = doc.add_table(rows=rows_visual, cols=cols_visual)
    vis_table.style = 'Table Grid'
    
    # Altura de fila: Aumentada un poco más para que quepan los 3 grupos musculares
    TR_HEIGHT_TWIPS = 3800 if incluir_analisis_muscular else 2800 
    
    for row in vis_table.rows:
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        trHeight = OxmlElement('w:trHeight')
        trHeight.set(qn('w:val'), str(TR_HEIGHT_TWIPS))
        trHeight.set(qn('w:hRule'), "atLeast")
        trPr.append(trHeight)
        set_row_cant_split(row)

    for i, row_data in enumerate(rutina_df.to_dict('records')):
        r = i // cols_visual
        c = i % cols_visual
        cell = vis_table.cell(r, c)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 1. Imagen
        ruta_img, msg = encontrar_imagen_recursiva(row_data['Imagen'])
        if ruta_img:
            try:
                run = p.add_run()
                run.add_picture(ruta_img, width=Inches(2.4), height=Inches(1.55))
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(2)
            except:
                p.add_run(f"[Error]\n")
        else:
            p.add_run(f"\n[FOTO NO DISPONIBLE]\n")
        
        # 2. Nombre Ejercicio
        run_nom = p.add_run("\n" + row_data['Ejercicio'])
        run_nom.font.bold = True
        run_nom.font.size = Pt(10)
        
        # 3. Análisis Muscular (Si se solicita)
        if incluir_analisis_muscular:
            p.add_run("\n" + "_"*25 + "\n").font.size = Pt(6)
            
            # Agonistas
            run_ago_label = p.add_run("Músculos Agonistas:\n")
            run_ago_label.font.bold = True
            run_ago_label.font.size = Pt(8)
            ago_text = str(row_data.get('agonistas', ''))
            p.add_run(f"{ago_text}\n").font.size = Pt(8)
            
            # Sinergistas
            run_sin_label = p.add_run("Músculos Sinergistas:\n")
            run_sin_label.font.bold = True
            run_sin_label.font.size = Pt(8)
            sin_text = str(row_data.get('sinergistas', ''))
            p.add_run(f"{sin_text}\n").font.size = Pt(8)

            # Estabilizadores (NUEVO)
            run_est_label = p.add_run("Músculos Estabilizadores:\n")
            run_est_label.font.bold = True
            run_est_label.font.size = Pt(8)
            est_text = str(row_data.get('estabilizadores', ''))
            p.add_run(f"{est_text}").font.size = Pt(8)

    doc.add_page_break()

    # PÁGINA 2
    h2 = doc.add_heading(level=1)
    run_h2 = h2.add_run('2. Rutina Detallada')
    run_h2.font.size = Pt(18)
    run_h2.font.color.rgb = RGBColor(44, 62, 80)

    tech_table = doc.add_table(rows=1, cols=6)
    tech_table.style = 'Table Grid'
    tech_table.autofit = False 
    widths = [0.7, 3.5, 1.5, 1.0, 1.5, 2.4] 
    headers = ["Orden", "Ejercicio", "Series x Reps", "Carga", "Descanso", "Notas"]
    row_hdr = tech_table.rows[0]
    set_row_cant_split(row_hdr)
    for i, h in enumerate(headers):
        style_header_cell(row_hdr.cells[i], h, widths[i])
        
    for idx, row_data in rutina_df.iterrows():
        row_cells = tech_table.add_row().cells
        set_row_cant_split(tech_table.rows[-1])
        for i in range(6):
            row_cells[i].width = Inches(widths[i])
        row_cells[0].text = str(idx + 1)
        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row_cells[1].text = row_data['Ejercicio']
        row_cells[2].text = f"{series_str} x {row_data['Reps']}"
        row_cells[3].text = f"{row_data['Peso']} kg"
        row_cells[4].text = row_data['Descanso']
        row_cells[5].text = f"Int: {row_data['Intensidad_Real']}" 

    doc.add_paragraph("\n")

    if lista_estiramientos:
        h3 = doc.add_heading(level=1)
        run_h3 = h3.add_run('3. Ejercicios de Estiramientos')
        run_h3.font.size = Pt(18)
        run_h3.font.color.rgb = RGBColor(44, 62, 80)
        h3.paragraph_format.keep_with_next = True

        num_est = len(lista_estiramientos)
        cols_est = 4
        rows_est = (num_est + cols_est - 1) // cols_est
        est_table = doc.add_table(rows=rows_est, cols=cols_est)
        est_table.style = 'Table Grid'
        for row in est_table.rows:
            tr = row._tr
            trPr = tr.get_or_add_trPr()
            trHeight = OxmlElement('w:trHeight')
            trHeight.set(qn('w:val'), str(2600))
            trHeight.set(qn('w:hRule'), "atLeast")
            trPr.append(trHeight)
            set_row_cant_split(row)

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
                    run.add_picture(ruta_img, width=Inches(2.2), height=Inches(1.4))
                    p.paragraph_format.space_before = Pt(3)
                    p.paragraph_format.space_after = Pt(3)
                except:
                    p.add_run(f"[Error]\n")
            else:
                p.add_run(f"\n[FOTO NO DISPONIBLE]\n")
            run_nom = p.add_run("\n" + item_est['nombre'])
            run_nom.font.bold = True
            run_nom.font.size = Pt(9)
        doc.add_paragraph("\n")

    h4 = doc.add_heading(level=1)
    run_h4 = h4.add_run('4. Percepción del Esfuerzo (RPE) - Escala de Borg')
    run_h4.font.size = Pt(18)
    run_h4.font.color.rgb = RGBColor(44, 62, 80)
    h4.paragraph_format.keep_with_next = True

    borg_table = doc.add_table(rows=3, cols=5)
    borg_table.style = 'Table Grid'
    borg_table.autofit = True
    for row in borg_table.rows:
        set_row_cant_split(row)

    borg_data = [
        {"val": "6-8", "txt": "Muy Ligero", "icon": "🙂", "color": "A9DFBF"},
        {"val": "9-11", "txt": "Ligero", "icon": "😌", "color": "D4EFDF"},
        {"val": "12-14", "txt": "Algo Duro", "icon": "😐", "color": "F9E79F"},
        {"val": "15-17", "txt": "Duro", "icon": "😓", "color": "F5CBA7"},
        {"val": "18-20", "txt": "Máximo", "icon": "🥵", "color": "E6B0AA"}
    ]
    
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

    row_text = borg_table.rows[1]
    for i, data in enumerate(borg_data):
        c = row_text.cells[i]
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(data['txt']).font.bold = True
        set_cell_bg_color(c, data['color'])

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

    doc.add_paragraph("\n")

    # --- SECCIÓN 5: MARCO TEÓRICO ---
    h5 = doc.add_heading(level=1)
    run_h5 = h5.add_run(f"5. {objetivo.upper()}")
    run_h5.font.size = Pt(18)
    run_h5.font.color.rgb = RGBColor(44, 62, 80)
    
    raw_text = INFO_OBJETIVOS.get(objetivo, "Información no disponible.")
    clean_lines = raw_text.split('\n')[1:] 
    
    emojis_clave = ['🎯', '🏋️‍♂️', '❤️', '🔁', '🟢', '🟡', '🔵', '🔥', '🔹', '🧠', '⚠️']
    
    for line in clean_lines:
        if not line.strip(): 
            continue 
            
        p_teoria = doc.add_paragraph()
        
        if any(line.strip().startswith(e) for e in emojis_clave):
            parts = line.strip().split(' ', 1)
            emoji_part = parts[0]
            text_part = parts[1] if len(parts) > 1 else ""
            
            r_emo = p_teoria.add_run(emoji_part + " ")
            r_emo.font.size = Pt(18) 
            
            r_txt = p_teoria.add_run(text_part)
            r_txt.font.size = Pt(11) 
        else:
            r_normal = p_teoria.add_run(line)
            r_normal.font.size = Pt(11)

    doc.add_paragraph("\n")

    # --- SECCIÓN 6: RESUMEN (IMAGEN) ---
    h6 = doc.add_heading(level=1)
    run_h6 = h6.add_run('6. RESUMEN DE FORMAS DE TRABAJO')
    run_h6.font.size = Pt(18)
    run_h6.font.color.rgb = RGBColor(44, 62, 80)
    h6.paragraph_format.keep_with_next = True 
    
    ruta_resumen, msg = encontrar_imagen_recursiva("tabla_resumen") 
    if ruta_resumen:
        try:
            doc.add_picture(ruta_resumen, width=Inches(9.0))
        except:
            doc.add_paragraph("[Error al insertar la imagen de resumen]")
    else:
        doc.add_paragraph("[Imagen 'tabla_resumen' no encontrada]")

    doc.add_paragraph("\n")

    # --- SECCIÓN 7: REFLEXIÓN ALUMNO ---
    h7 = doc.add_heading(level=1)
    run_h7 = h7.add_run('7. MI CIRCUITO DE TRABAJO SE BASA EN LOS SIGUIENTES PRINCIPIOS DE ENTRENAMIENTO Y SIGUE LA SIGUIENTE LÓGICA')
    run_h7.font.size = Pt(14)
    run_h7.font.color.rgb = RGBColor(44, 62, 80)
    h7.paragraph_format.keep_with_next = True
    
    p_inst = doc.add_paragraph("(Explica cómo y por qué estableces este circuito según tus objetivos y criterios científicos):")
    p_inst.paragraph_format.space_after = Pt(200) 

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ STREAMLIT ---

st.markdown("""
<style>
.big-font { font-size:30px !important; font-weight: bold; }
.sub-font { font-size:20px !important; font-style: italic; color: #555; }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="big-font">Generador Científico de Rutinas creado por José Carlos Tejedor Lorenzo</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-font">Situación de aprendizaje: Trabajo en Salas de Musculación 1º de Bachillerato IES Lucía de Medrano.</p>', unsafe_allow_html=True)
st.markdown("---")

imagenes_encontradas = []
for root, dirs, files in os.walk("."):
    for file in files:
        if file.lower().endswith(('.png', '.jpg', '.jpeg')):
            imagenes_encontradas.append(file)
if imagenes_encontradas:
    st.sidebar.success(f"✅ {len(imagenes_encontradas)} imágenes detectadas.")
else:
    st.sidebar.error("❌ No hay imágenes en GitHub.")

if DB_EJERCICIOS is None:
    st.error("Error: DB_EJERCICIOS.xlsx no encontrado.")
    st.stop()
elif isinstance(DB_EJERCICIOS, str):
    st.error(DB_EJERCICIOS)
    st.stop()

col1, col2 = st.columns(2)
with col1:
    alumno = st.text_input("Nombre del Alumno:", "", key=get_key("alumno"))
    
    # 1. OBTENER TIPOS
    tipos_todos = sorted(list(set([e['tipo'] for e in DB_EJERCICIOS if e['tipo']])))
    tipos_entreno = [t for t in tipos_todos if 'estiramiento' not in t.lower()]
    
    # DEFAULT = None PARA EMPEZAR VACÍO
    sel_tipos = st.multiselect(
        "Material de Entrenamiento (Elige para empezar):", 
        options=tipos_entreno, 
        default=None, 
        key=get_key("sel_material")
    )
    
    cardio_seleccion = st.selectbox(
        "Todos los días cardio:", 
        ["Bicicleta", "Cinta de Correr", "Step", "Remo de cardio"],
        key=get_key("cardio")
    )

with col2:
    # 4 OPCIONES DE OBJETIVO (AHORA INCLUYE REHABILITACIÓN)
    objetivo = st.selectbox("Objetivo:", 
                            ["Hipertrofia Muscular", "Definición Muscular", "Resistencia Muscular", "Programa de Pérdida de Peso", "Rehabilitación Muscular y Articular", "Fuerza Máxima", "Mantenimiento Muscular"], 
                            key=get_key("objetivo"))
    
    intensidad_seleccionada = 0
    reps_seleccionadas = ""
    descanso_seleccionado = ""
    cardio_duracion = "" 
    series_finales = ""
    
    if objetivo == "Fuerza Máxima":
        st.info("Rango: 1-5 Reps | Intensidad: 85-100% RM | 4-6 Series")
        cardio_duracion = "Bajo"
        series_finales = "4-6"
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [85, 90, 95, 100], key=get_key("int_fm"))
        with col_b:
            val_reps = st.selectbox("Repeticiones:", [1, 2, 3, 4, 5], key=get_key("reps_fm"))
            reps_seleccionadas = str(val_reps)
        with col_c:
            descanso_seleccionado = st.selectbox("Descanso:", ["3 min", "4 min", "5 min", "6 min"], key=get_key("desc_fm"))

    elif objetivo == "Hipertrofia Muscular":
        st.info("Rango: 6-12 Reps | Intensidad: 65-85% RM | 3-6 Series")
        cardio_duracion = "Moderado"
        series_finales = "3-6"
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [65, 70, 75, 80, 85], key=get_key("int_hyp"))
        with col_b:
            val_reps = st.selectbox("Repeticiones:", [6, 7, 8, 9, 10, 11, 12], key=get_key("reps_hyp"))
            reps_seleccionadas = str(val_reps)
        with col_c:
            descanso_seleccionado = st.selectbox("Descanso:", ["60 seg", "90 seg", "120 seg"], key=get_key("desc_hyp"))

    elif objetivo == "Definición Muscular":
        st.info("Rango: 10-15 Reps | Intensidad: 60-75% RM | 3-5 Series")
        cardio_duracion = "Alto"
        series_finales = "3-5"
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [60, 65, 70, 75], key=get_key("int_def"))
        with col_b:
            val_reps = st.selectbox("Repeticiones:", [10, 11, 12, 13, 14, 15], key=get_key("reps_def"))
            reps_seleccionadas = str(val_reps)
        with col_c:
            descanso_seleccionado = st.selectbox("Descanso:", ["30 seg", "45 seg", "60 seg"], key=get_key("desc_def"))

    elif objetivo == "Programa de Pérdida de Peso":
        st.info("Rango: 12-20 Rps | Intensidad: 50–70 % RM | 3-4 Series")
        cardio_duracion = "30-60 min + HIIT"
        series_finales = "3-4"
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [50, 55, 60, 65, 70], key=get_key("int_pp"))
        with col_b:
            val_reps = st.selectbox("Repeticiones:", [12, 13, 14, 15, 16, 17, 18, 19, 20], key=get_key("reps_pp"))
            reps_seleccionadas = str(val_reps)
        with col_c:
            descanso_seleccionado = st.selectbox("Descanso:", ["20 seg", "30 seg", "45 seg"], key=get_key("desc_pp"))

    elif objetivo == "Resistencia Muscular":
        st.info("Rango: 15-30+ Reps | Intensidad: 30-60% RM | 2-4 Series")
        cardio_duracion = "Muy Alto"
        series_finales = "2-4"
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [30, 35, 40, 45, 50, 55, 60], key=get_key("int_res"))
        with col_b:
            val_reps = st.selectbox("Repeticiones:", [15, 20, 25, 30], key=get_key("reps_res"))
            reps_seleccionadas = str(val_reps)
        with col_c:
            descanso_seleccionado = st.selectbox("Descanso:", ["15 seg", "30 seg", "45 seg"], key=get_key("desc_res"))

    elif objetivo == "Mantenimiento Muscular":
        st.info("Rango: 8-12 Reps | Intensidad: 60-75% RM | 2-3 Series")
        cardio_duracion = "Moderado"
        series_finales = "2-3"
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [60, 65, 70, 75], key=get_key("int_man"))
        with col_b:
            val_reps = st.selectbox("Repeticiones:", [8, 9, 10, 11, 12], key=get_key("reps_man"))
            reps_seleccionadas = str(val_reps)
        with col_c:
            descanso_seleccionado = st.selectbox("Descanso:", ["60 seg", "75 seg", "90 seg"], key=get_key("desc_man"))

    elif objetivo == "Rehabilitación Muscular y Articular":
        st.info("Rango: 1-30 Reps | Intensidad: 20-60% RM | 1-5 Series | Depende Fase")
        cardio_duracion = "Muy bajo"
        col_rh1, col_rh2, col_rh3 = st.columns(3)
        with col_rh1:
            intensidad_seleccionada = st.selectbox("Intensidad (% RM):", [20, 30, 40, 50, 60], key=get_key("int_rehab"))
        with col_rh2:
            val_reps = st.number_input("Nº Repeticiones:", 1, 30, 12, key=get_key("reps_rehab"))
            reps_seleccionadas = str(val_reps)
        with col_rh3:
            descanso_seleccionado = st.selectbox("Descanso:", ["30 seg", "45 seg", "1 min", "2 min"], key=get_key("desc_rehab"))
        
        series_finales = st.selectbox("Series:", ["1", "2", "3", "4", "5"], index=2, key=get_key("ser_reh"))

if sel_tipos:
    ej_filtrados = [e for e in DB_EJERCICIOS if e['tipo'] in sel_tipos]
    
    # RANGO 1-12
    default_val = 8 if objetivo == "Rehabilitación Muscular y Articular" else 6
    max_val = min(12, len(ej_filtrados)) 
    
    if default_val > max_val: default_val = max_val
    if default_val < 1: default_val = 1
    
    num_ej = st.slider("Cantidad de Ejercicios:", 1, 12, default_val, key=get_key("slider_ej"))
else:
    st.warning("👈 Selecciona primero el Material de Entrenamiento para ver los ejercicios.")
    st.stop()

st.subheader("Selección de Ejercicios")

if sel_tipos:
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
    seleccion = st.multiselect("Elige los ejercicios:", nombres_fil, max_selections=num_ej, key=get_key("sel_ej"))

    rellenar_auto = st.checkbox(f"Rellenar automáticamente hasta llegar a {num_ej} ejercicios", value=True, key=get_key("check_auto"))

    seleccionados_data = []
    nombres_finales = seleccion.copy()

    # --- ESTABILIZACIÓN DE LA SELECCIÓN ---
    config_id = f"{sel_tipos}_{num_ej}_{seleccion}_{rellenar_auto}_{st.session_state.reset_counter}"
    
    if 'last_config_id' not in st.session_state or st.session_state.last_config_id != config_id:
        if rellenar_auto and len(nombres_finales) < num_ej:
            pool = [x for x in ej_filtrados if x['nombre'] not in nombres_finales]
            needed = num_ej - len(nombres_finales)
            if needed <= len(pool):
                extras = random.sample(pool, needed)
                nombres_finales.extend([x['nombre'] for x in extras])
        st.session_state.final_names = nombres_finales
        st.session_state.last_config_id = config_id
    
    nombres_finales_estables = st.session_state.final_names
    
    seleccionados_data = []
    for nom in nombres_finales_estables:
        obj_ejercicio = next((x for x in ej_filtrados if x['nombre'] == nom), None)
        if obj_ejercicio:
            seleccionados_data.append(obj_ejercicio)

    st.markdown("---")
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
            # === MEMORIA INTELIGENTE PARA 1RM ===
            val_key = f"rm_{i}_{ej['nombre']}_{st.session_state.reset_counter}"
            rm_inputs[ej['nombre']] = st.number_input(
                f"1RM {ej['nombre']} (kg)", 
                min_value=0, 
                max_value=500, 
                value=60, 
                step=1, 
                key=val_key
            )

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

        num_est_select = st.slider("Cantidad de estiramientos:", 1, 12, 4, key=get_key("slider_est"))
        seleccion_est = st.multiselect("Elige estiramientos:", nombres_est, max_selections=num_est_select, key=get_key("sel_est"))
        
        # Estabilización de estiramientos
        config_est_id = f"EST_{num_est_select}_{seleccion_est}_{st.session_state.reset_counter}"
        
        if 'last_est_id' not in st.session_state or st.session_state.last_est_id != config_est_id:
            estiramientos_finales_nombres = seleccion_est.copy()
            if len(estiramientos_finales_nombres) < num_est_select:
                pool_est = [x['nombre'] for x in pool_estiramientos if x['nombre'] not in estiramientos_finales_nombres]
                needed_est = num_est_select - len(estiramientos_finales_nombres)
                if needed_est <= len(pool_est):
                     estiramientos_finales_nombres.extend(random.sample(pool_est, needed_est))
            st.session_state.final_est_names = estiramientos_finales_nombres
            st.session_state.last_est_id = config_est_id
            
        estiramientos_finales = []
        for nom in st.session_state.final_est_names:
             estiramientos_finales.append(next(x for x in pool_estiramientos if x['nombre'] == nom))

    else:
        st.warning("⚠️ No se han encontrado ejercicios marcados como 'Estiramientos' en el Excel.")
        estiramientos_finales = []

    # --- BOTONES FINALES ---
    st.write("---")
    st.subheader("Generar Informe")
    col_pdf, col_reset = st.columns([2, 1])

    with col_pdf:
        if st.button("📄 GENERAR DOCUMENTO ESTÁNDAR", type="primary", use_container_width=True, key=get_key("btn_std")):
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
                    "Intensidad_Real": f"{intensidad_seleccionada}%",
                    "agonistas": item.get('agonistas', ''),
                    "sinergistas": item.get('sinergistas', ''),
                    "estabilizadores": item.get('estabilizadores', '')
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
                cardio_tiempo=cardio_duracion,
                series_str=series_finales,
                incluir_analisis_muscular=False # <--- MODO ESTÁNDAR
            )
            st.success(f"Informe Estándar Generado: {objetivo}")
            st.download_button("📥 Descargar Word Estándar", docx, f"Rutina_{alumno}_Estandar.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=get_key("dl_std"))

        if st.button("🧬 GENERAR DOCUMENTO CON ANÁLISIS MUSCULAR", use_container_width=True, key=get_key("btn_ana")):
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
                    "Intensidad_Real": f"{intensidad_seleccionada}%",
                    "agonistas": item.get('agonistas', ''),
                    "sinergistas": item.get('sinergistas', ''),
                    "estabilizadores": item.get('estabilizadores', '')
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
                cardio_tiempo=cardio_duracion,
                series_str=series_finales,
                incluir_analisis_muscular=True # <--- MODO ANÁLISIS
            )
            st.success(f"Informe con Análisis Muscular Generado: {objetivo}")
            st.download_button("📥 Descargar Word con Análisis", docx, f"Rutina_{alumno}_Analisis.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=get_key("dl_ana"))

# --- LÓGICA DE REINICIO ---
def reset_app():
    st.session_state.reset_counter += 1
    st.cache_data.clear()
    if 'last_config_id' in st.session_state: del st.session_state.last_config_id
    if 'last_est_id' in st.session_state: del st.session_state.last_est_id

with col_reset:
    st.write("")
    st.button("🔄 Reiniciar", use_container_width=True, on_click=reset_app)
