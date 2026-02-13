import streamlit as st
import json
import pandas as pd
import io
import os
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

# --- CONFIGURACIÓN VISUAL INSTITUCIONAL ---
st.set_page_config(page_title="Simetrik Docs Generator | PeYa", page_icon="📊", layout="wide")

COLOR_PEYA_RED = "EA0050"
COLOR_WHITE = "FFFFFF"
COLOR_GREY_LIGHT = "F2F2F2"
COLOR_BORDER = "D9D9D9"

def style_cell(cell, is_header=False, is_summary=False):
    """Aplica el brand book de PeYa a las celdas de Excel."""
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
    """Extrae lógica de cruces v_lookup detallada del JSON."""
    v_info = col.get('v_lookup')
    if not v_info:
        return ""
    
    v_set = v_info.get('v_lookup_set', {})
    origin_id = v_set.get('origin_source_id')
    origin_name = res_map.get(origin_id, f"ID: {origin_id}")
    
    rules = v_set.get('rules', [])
    logic_parts = []
    for r in rules:
        col_a = col_map.get(r.get('column_a_id'), "Col_A")
        col_b = col_map.get(r.get('column_b_id'), "Col_B")
        logic_parts.append(f"{col_a} == {col_b}")
    
    return f"🔍 DESDE: {origin_name}\n🔑 MATCH: {' & '.join(logic_parts)}"

def procesar_json_pro(json_file):
    data = json.load(json_file)
    resources = data.get('resources', [])
    nodes = data.get('nodes', [])
    
    # 1. Diccionarios de Mapeo
    res_map = {r.get('export_id'): r.get('name') for r in resources}
    col_map = {}
    for r in resources:
        for c in r.get('columns', []):
            col_map[c.get('export_id')] = c.get('label') or c.get('name')

    # 2. Análisis de Relaciones (Nodos)
    relaciones = {r.get('export_id'): {"parents": [], "children": []} for r in resources}
    for n in nodes:
        t_id = n.get('target')
        s_val = n.get('source')
        if t_id and s_val:
            s_list = s_val if isinstance(s_val, list) else [s_val]
            for sid in s_list:
                if t_id in relaciones: relaciones[t_id]["parents"].append(res_map.get(sid, sid))
                if sid in relaciones: relaciones[sid]["children"].append(res_map.get(t_id, t_id))

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        nombres_usados = set()
        
        # --- HOJA: ÍNDICE ---
        ws_idx = writer.book.create_sheet("📚 Índice de Recursos", 0)
        headers = ["ID", "RECURSO", "TIPO", "PROVIENE DE", "ALIMENTA A", "ACCESO"]
        for i, h in enumerate(headers, 1):
            cell = ws_idx.cell(1, i, h)
            style_cell(cell, is_header=True)

        for idx, res in enumerate(resources, 2):
            eid = res.get('export_id')
            # Limpiar nombre de hoja para Excel
            sheet_name = "".join(x for x in res.get('name')[:25] if x.isalnum() or x==' ').strip()
            if not sheet_name: sheet_name = f"Recurso_{eid}"
            if sheet_name in nombres_usados: sheet_name = f"{sheet_name[:20]}_{eid}"
            nombres_usados.add(sheet_name)

            ws_idx.cell(idx, 1, eid)
            ws_idx.cell(idx, 2, res.get('name'))
            ws_idx.cell(idx, 3, res.get('resource_type', '').replace('_', ' ').upper())
            ws_idx.cell(idx, 4, ", ".join(relaciones[eid]["parents"]) or "Input Directo")
            ws_idx.cell(idx, 5, ", ".join(relaciones[eid]["children"]) or "Nodo Final")
            
            link = ws_idx.cell(idx, 6, "Ver Detalle")
            link.hyperlink = f"#'{sheet_name}'!A1"
            link.font = Font(color="0000FF", underline="single")
            
            for i in range(1, 6): style_cell(ws_idx.cell(idx, i))

        # Ajuste de columnas índice
        for col_l in ['B', 'D', 'E']: ws_idx.column_dimensions[col_l].width = 35

        # --- HOJAS: DETALLE ---
        progress_bar = st.progress(0)
        for i, res in enumerate(resources):
            eid = res.get('export_id')
            # Recuperar el nombre exacto usado en el índice
            s_name = [n for n in nombres_usados if str(eid) in n or res.get('name')[:10] in n][0]
            
            ws = writer.book.create_sheet(s_name)
            
            # Header del recurso
            ws.merge_cells('A1:C1')
            ws['A1'] = f"DETALLE TÉCNICO: {res.get('name')}"
            ws['A1'].font = Font(bold=True, size=14, color=COLOR_PEYA_RED)
            
            ws.cell(2, 1, "ID Export:"); ws.cell(2, 2, eid)
            ws.cell(3, 1, "Tipo:"); ws.cell(3, 2, res.get('resource_type', '').upper())
            for r_row in [2, 3]:
                style_cell(ws.cell(r_row, 1), is_summary=True)
                style_cell(ws.cell(r_row, 2))

            # Tabla de Columnas
            ws.cell(5, 1, "ESTRUCTURA Y TRANSFORMACIONES").font = Font(bold=True, color=COLOR_PEYA_RED)
            h_cols = ["COLUMNA (LABEL)", "FORMATO", "LÓGICA / BUSCAR V"]
            for j, h in enumerate(h_cols, 1):
                style_cell(ws.cell(6, j, h), is_header=True)

            curr = 7
            for col in res.get('columns', []):
                ws.cell(curr, 1, col.get('label') or col.get('name'))
                ws.cell(curr, 2, col.get('data_format', 'string'))
                
                v_detail = get_v_lookup_details(col, res_map, col_map)
                queries = [t.get('query') for t in (col.get('transformations') or []) if t.get('query')]
                logic_text = v_detail + ("\n" if v_detail and queries else "") + "\n".join([f"⚙️ {q}" for q in queries])
                
                ws.cell(curr, 3, logic_text or "Dato Original")
                for j in range(1, 4): style_cell(ws.cell(curr, j))
                curr += 1
            
            ws.column_dimensions['A'].width = 30
            ws.column_dimensions['C'].width = 70
            progress_bar.progress((i + 1) / len(resources))

    if "Sheet" in writer.book.sheetnames: writer.book.remove(writer.book["Sheet"])
    output.seek(0)
    return output

# --- UI STREAMLIT ---
st.markdown(f"""
    <div style='background-color: #{COLOR_PEYA_RED}; padding: 20px; border-radius: 10px; margin-bottom: 25px;'>
        <h1 style='color: white; margin: 0;'>Simetrik Documentation Pro</h1>
        <p style='color: white; opacity: 0.8;'>PedidosYa Financial Operations & Control</p>
    </div>
    """, unsafe_allow_html=True)

up = st.file_uploader("Arrastra el JSON exportado de Simetrik", type=['json'])

if up:
    # Lógica de nombre de archivo dinámico
    nombre_base = os.path.splitext(up.name)[0] # Quita el .json
    fecha_hoy = datetime.now().strftime("%Y-%m-%d_%H%M")
    nombre_descarga = f"{nombre_base}_{fecha_hoy}.xlsx"

    if st.button("🚀 Generar Reporte Profesional"):
        try:
            with st.spinner('Analizando estructura y mapeando relaciones...'):
                processed_data = procesar_json_pro(up)
            
            st.success(f"✅ ¡Estructura de '{up.name}' procesada con éxito!")
            st.download_button(
                label="📥 Descargar Reporte Excel", 
                data=processed_data, 
                file_name=nombre_descarga,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error crítico procesando el archivo: {e}")
