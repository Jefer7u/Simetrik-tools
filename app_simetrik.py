import streamlit as st
import json
import pandas as pd
import io
import os
import re
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

# --- CONFIGURACIÓN VISUAL INSTITUCIONAL ---
st.set_page_config(page_title="Simetrik Docs Generator | PeYa", page_icon="📊", layout="wide")

COLOR_PEYA_RED = "EA0050"
COLOR_WHITE = "FFFFFF"
COLOR_GREY_LIGHT = "F2F2F2"
COLOR_BORDER = "D9D9D9"

def style_cell(cell, is_header=False, is_summary=False):
    thin = Side(border_style="thin", color=COLOR_BORDER)
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    cell.alignment = Alignment(vertical='center', wrap_text=True, horizontal='left')
    if is_header:
        cell.fill = PatternFill(start_color=COLOR_PEYA_RED, end_color=COLOR_PEYA_RED, fill_type="solid")
        cell.font = Font(color=COLOR_WHITE, bold=True, size=11, name='Arial')
        cell.alignment = Alignment(horizontal='center', vertical='center')
    elif is_summary:
        cell.fill = PatternFill(start_color=COLOR_GREY_LIGHT, end_color=COLOR_GREY_LIGHT, fill_type="solid")
        cell.font = Font(bold=True, name='Arial')

def get_v_lookup_details(col, res_map, col_map):
    v_info = col.get('v_lookup')
    if not v_info: return ""
    v_set = v_info.get('v_lookup_set', {})
    origin_name = res_map.get(v_set.get('origin_source_id'), f"ID: {v_set.get('origin_source_id')}")
    rules = v_set.get('rules', [])
    logic_parts = [f"[{col_map.get(r.get('column_a_id'), 'Col_A')} == {col_map.get(r.get('column_b_id'), 'Col_B')}]" for r in rules]
    return f"🔍 DESDE: {origin_name}\n🔑 MATCH: {' & '.join(logic_parts)}"

def limpiar_nombre_excel(nombre, export_id):
    # Excel solo permite 31 caracteres y prohíbe ciertos símbolos
    nombre_limpio = re.sub(r'[\\/*?:[\]]', '', str(nombre))
    # Dejamos espacio para el ID al final para que sea único
    res = f"{nombre_limpio[:20]}_{export_id}"
    return res[:31]

def procesar_json_pro(json_file):
    data = json.load(json_file)
    resources = data.get('resources', [])
    nodes = data.get('nodes', [])
    
    # 1. Mapeos
    res_map = {r.get('export_id'): r.get('name') for r in resources}
    col_map = {}
    for r in resources:
        for c in r.get('columns', []):
            col_map[c.get('export_id')] = c.get('label') or c.get('name')

    # 2. Relaciones
    relaciones = {r.get('export_id'): {"parents": [], "children": []} for r in resources}
    for n in nodes:
        t_id, s_val = n.get('target'), n.get('source')
        if t_id and s_val:
            s_list = s_val if isinstance(s_val, list) else [s_val]
            for sid in s_list:
                if t_id in relaciones: relaciones[t_id]["parents"].append(res_map.get(sid, sid))
                if sid in relaciones: relaciones[sid]["children"].append(res_map.get(t_id, t_id))

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # --- ÍNDICE ---
        ws_idx = writer.book.create_sheet("📚 Índice de Recursos", 0)
        headers = ["ID RECURSO", "NOMBRE RECURSO", "TIPO", "PROVIENE DE", "ALIMENTA A", "ACCESO"]
        for i, h in enumerate(headers, 1):
            style_cell(ws_idx.cell(1, i, h), is_header=True)

        # Diccionario para guardar el mapeo ID -> Nombre de Hoja
        map_hojas = {}

        for idx, res in enumerate(resources, 2):
            eid = res.get('export_id')
            sheet_name = limpiar_nombre_excel(res.get('name'), eid)
            map_hojas[eid] = sheet_name

            ws_idx.cell(idx, 1, eid)
            ws_idx.cell(idx, 2, res.get('name'))
            ws_idx.cell(idx, 3, str(res.get('resource_type', '')).upper())
            ws_idx.cell(idx, 4, ", ".join(relaciones[eid]["parents"]) or "Fuente Primaria")
            ws_idx.cell(idx, 5, ", ".join(relaciones[eid]["children"]) or "Nodo Final")
            
            cell_link = ws_idx.cell(idx, 6, "Ir al detalle")
            cell_link.hyperlink = f"#'{sheet_name}'!A1"
            cell_link.font = Font(color="0000FF", underline="single")
            
            for i in range(1, 6): style_cell(ws_idx.cell(idx, i))

        for col_l in ['B', 'D', 'E']: ws_idx.column_dimensions[col_l].width = 35
        ws_idx.column_dimensions['A'].width = 15

        # --- DETALLES ---
        for res in resources:
            eid = res.get('export_id')
            ws = writer.book.create_sheet(map_hojas[eid])
            
            ws.merge_cells('A1:C1')
            ws['A1'] = f"DETALLE: {res.get('name')}"
            ws['A1'].font = Font(bold=True, size=14, color=COLOR_PEYA_RED)
            
            ws.cell(2, 1, "ID Export:"); ws.cell(2, 2, eid)
            ws.cell(3, 1, "Tipo:"); ws.cell(3, 2, str(res.get('resource_type', '')).upper())
            for r_row in [2, 3]:
                style_cell(ws.cell(r_row, 1), is_summary=True)
                style_cell(ws.cell(r_row, 2))

            # Conectividad
            ws.cell(5, 1, "FLUJO DE DATOS").font = Font(bold=True, color=COLOR_PEYA_RED)
            ws.cell(6, 1, "PROVIENE DE:"); ws.cell(6, 2, ", ".join(relaciones[eid]["parents"]) or "N/A")
            ws.cell(7, 1, "ALIMENTA A:"); ws.cell(7, 2, ", ".join(relaciones[eid]["children"]) or "N/A")
            for r_row in [6, 7]:
                style_cell(ws.cell(r_row, 1), is_summary=True)
                style_cell(ws.cell(r_row, 2))

            # Columnas
            ws.cell(9, 1, "ESTRUCTURA DE COLUMNAS").font = Font(bold=True, color=COLOR_PEYA_RED)
            h_cols = ["COLUMNA", "TIPO", "LÓGICA / BUSCAR V"]
            for j, h in enumerate(h_cols, 1): style_cell(ws.cell(10, j, h), is_header=True)

            curr = 11
            for col in res.get('columns', []):
                ws.cell(curr, 1, col.get('label') or col.get('name'))
                ws.cell(curr, 2, col.get('data_format', 'string'))
                
                v_detail = get_v_lookup_details(col, res_map, col_map)
                queries = [t.get('query') for t in (col.get('transformations') or []) if t.get('query')]
                logic_text = v_detail + ("\n" if v_detail and queries else "") + "\n".join([f"⚙️ {q}" for q in queries])
                
                ws.cell(curr, 3, logic_text or "-")
                for j in range(1, 4): style_cell(ws.cell(curr, j))
                curr += 1
            
            ws.column_dimensions['A'].width = 30
            ws.column_dimensions['C'].width = 70

    if "Sheet" in writer.book.sheetnames: writer.book.remove(writer.book["Sheet"])
    output.seek(0)
    return output

# --- UI ---
st.markdown(f"<div style='background-color:#{COLOR_PEYA_RED};padding:20px;border-radius:10px'><h1 style='color:white;margin:0'>Simetrik Doc Pro | PeYa</h1></div>", unsafe_allow_html=True)
up = st.file_uploader("Sube el archivo JSON", type=['json'])

if up:
    nombre_descarga = f"{os.path.splitext(up.name)[0]}_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
    if st.button("🚀 Generar Reporte"):
        try:
            with st.spinner('Procesando...'):
                data_excel = procesar_json_pro(up)
            st.success("✅ ¡Listo!")
            st.download_button("📥 Descargar Excel", data_excel, nombre_descarga)
        except Exception as e:
            st.error(f"Error: {e}")
