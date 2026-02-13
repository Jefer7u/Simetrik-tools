import streamlit as st
import json
import pandas as pd
import io
import os
import re
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

# --- CONFIGURACIÓN VISUAL INSTITUCIONAL PEYA ---
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
    return f"🔍 BUSCAR V EN: {origin_name}\n🔑 MATCH: {' & '.join(logic_parts)}"

def limpiar_nombre_excel(nombre, export_id):
    nombre_limpio = re.sub(r'[\\/*?:[\]]', '', str(nombre))
    res = f"{nombre_limpio[:20]}_{export_id}"
    return res[:31]

def procesar_json_pro(json_file):
    data = json.load(json_file)
    resources = data.get('resources', [])
    nodes = data.get('nodes', [])
    
    res_map = {r.get('export_id'): r.get('name') for r in resources}
    col_map = {}
    for r in resources:
        for c in r.get('columns', []):
            col_map[c.get('export_id')] = c.get('label') or c.get('name')

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
        ws_idx = writer.book.create_sheet("📚 Índice de Flujo", 0)
        headers = ["ID RECURSO", "NOMBRE RECURSO", "TIPO", "PROVIENE DE", "ALIMENTA A", "LINK"]
        for i, h in enumerate(headers, 1):
            style_cell(ws_idx.cell(1, i, h), is_header=True)

        map_hojas = {}
        for idx, res in enumerate(resources, 2):
            eid = res.get('export_id')
            sheet_name = limpiar_nombre_excel(res.get('name'), eid)
            map_hojas[eid] = sheet_name

            ws_idx.cell(idx, 1, eid)
            ws_idx.cell(idx, 2, res.get('name'))
            ws_idx.cell(idx, 3, str(res.get('resource_type', '')).replace('_',' ').upper())
            ws_idx.cell(idx, 4, ", ".join(relaciones[eid]["parents"]) or "Origen")
            ws_idx.cell(idx, 5, ", ".join(relaciones[eid]["children"]) or "Fin de Flujo")
            
            cell_link = ws_idx.cell(idx, 6, "Ver detalle →")
            cell_link.hyperlink = f"#'{sheet_name}'!A1"
            cell_link.font = Font(color="0000FF", underline="single")
            for i in range(1, 6): style_cell(ws_idx.cell(idx, i))

        # --- DETALLES ---
        for res in resources:
            eid = res.get('export_id')
            ws = writer.book.create_sheet(map_hojas[eid])
            
            # Header
            ws.merge_cells('A1:C1')
            ws['A1'] = f"RECURSO: {res.get('name')}"
            ws['A1'].font = Font(bold=True, size=14, color=COLOR_PEYA_RED)
            
            ws.cell(2, 1, "ID RECURSO:"); ws.cell(2, 2, eid)
            ws.cell(3, 1, "TIPO:"); ws.cell(3, 2, str(res.get('resource_type', '')).replace('_',' ').upper())
            for r_row in [2, 3]:
                style_cell(ws.cell(r_row, 1), is_summary=True)
                style_cell(ws.cell(r_row, 2))

            # Columnas
            ws.cell(5, 1, "CONFIGURACIÓN DE COLUMNAS").font = Font(bold=True, color=COLOR_PEYA_RED)
            h_cols = ["COLUMNA", "TIPO DATO", "LÓGICA / TRANSFORMACIÓN"]
            for j, h in enumerate(h_cols, 1): style_cell(ws.cell(6, j, h), is_header=True)

            curr = 7
            for col in res.get('columns', []):
                ws.cell(curr, 1, col.get('label') or col.get('name'))
                ws.cell(curr, 2, col.get('data_format', 'string'))
                
                # Filtrar "N/A" y lógicas vacías
                v_detail = get_v_lookup_details(col, res_map, col_map)
                
                # Obtenemos transformaciones ignorando las que dicen "N/A"
                raw_queries = [t.get('query') for t in (col.get('transformations') or []) if t.get('query')]
                queries = [q for q in raw_queries if q.strip().upper() != "N/A"]
                
                logic_text = v_detail
                if queries:
                    if logic_text: logic_text += "\n"
                    logic_text += "\n".join([f"⚙️ {q}" for q in queries])
                
                ws.cell(curr, 3, logic_text or "Dato Directo")
                for j in range(1, 4): style_cell(ws.cell(curr, j))
                curr += 1
            
            ws.column_dimensions['A'].width = 35
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 75

    if "Sheet" in writer.book.sheetnames: writer.book.remove(writer.book["Sheet"])
    output.seek(0)
    return output

# --- STREAMLIT UI ---
st.markdown(f"""
    <div style='background-color:#{COLOR_PEYA_RED};padding:20px;border-radius:12px;text-align:center'>
        <h1 style='color:white;margin:0;font-family:Arial'>Simetrik Pro Documentation</h1>
        <p style='color:white;opacity:0.9'>PedidosYa Finance Operations</p>
    </div>""", unsafe_allow_html=True)

up = st.file_uploader("", type=['json'])

if up:
    nombre_descarga = f"{os.path.splitext(up.name)[0]}_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"
    if st.button("🚀 GENERAR EXCEL PROFESIONAL"):
        try:
            with st.spinner('Limpiando N/A y generando mapeos...'):
                data_excel = procesar_json_pro(up)
            st.success(f"¡Reporte listo! Nombre: {nombre_descarga}")
            st.download_button("📥 Descargar Reporte", data_excel, nombre_descarga)
        except Exception as e:
            st.error(f"Se produjo un error: {e}")
