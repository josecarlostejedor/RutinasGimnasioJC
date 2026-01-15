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
st.set_page_config(page_title="Generador Rutinas Final", layout="wide")

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

# --- BUSCADOR "RASTREADOR" DE IMÁGENES ---
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

# --- LÓGICA PARAMETROS ---
def obtener_parametros(objetivo):
    if objetivo == "Hipertrofia": return {"reps": "8-12", "descanso": "90 seg"}
    elif objetivo == "Fuerza Máxima": return {"reps": "3-5", "descanso": "3-5 min"}
    elif objetivo == "Resistencia": return {"reps": "15-20", "descanso": "45 seg"}

# --- GENERADOR WORD (ACTUALIZADO CON NUEVOS TÍTULOS) ---
def generar_word_final(rutina_df, objetivo, alumno, titulo_material):
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
    # LÓGICA DE TÍTULO MODIFICADA
    r1 = p.add_run(f"PROGRAMA DE ENTRAMIENTO DE: {titulo_material.upper()}\n")
    r1.font.bold = True
    r1.font.size = Pt(16)
    r1.font.color.rgb = RGBColor(41, 128, 185)
    
    # Manejo de nombre de alumno vacío
    nombre_mostrar = alumno if alumno.strip() else "ALUMNO"
    p.add_run(f"OBJETIVO: {objetivo} | ALUMNO: {nombre_mostrar.upper()}")

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
    
    headers = ["#", "Ejercicio", "Series x Reps", "Carga (Kg)", "Descanso",
