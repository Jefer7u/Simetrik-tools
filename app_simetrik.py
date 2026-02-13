import streamlit as st
import json
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

# --- CONFIGURACIÓN VISUAL PEYA ---
st.set_page_config(page_title="Simetrik Docs Generator", page_icon="📊", layout="wide")

COLOR_PEYA_MAIN = "EA0050"
COLOR_HEADER_TXT = "FFFFFF"
COLOR_BORDER = "CCCCCC"

def aplicar_estilos_peya(ws, min_row, max_row, min_col, max_col, es_header=False):
    thin = Side(border_style="thin", color=COLOR_BORDER)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            if es_header:
                cell.fill = PatternFill(start_color=COLOR_PEYA_MAIN, end_color=COLOR_PEYA_MAIN, fill_type="solid")
                cell.font = Font(color=COLOR_HEADER_TXT, bold=True, size=11)

def formatear_reglas_segmento(segmento, col_map):
    texto_logica = []
    try:
        for fs in (segmento.get('segment_filter_sets') or []):
            set_logic = []
            condicion_set = fs.get('condition', 'AND')
            for rule in (fs.get('segment_filter_rules') or []):
                col_id = rule.get('column_id')
                col_name = col_map.get(col_id, f"ID_{col_id}")
                op = str(rule.get('operator', '=')).lower()
                val = str(rule.get('value', ''))
                set_logic.append(f"[{col_name} {op} {val}]" if 'null' not in op else f"[{col_name} {op}]")
            if set_logic: texto_logica.append(f"({f' {condicion_set} '.join(set_logic)})")
    except: return "Error en filtros"
    return " OR ".join(texto_logica) if texto_logica else "Sin filtros"

def limpiar_nombre_hoja(nombre, nombres_usados):
    invalido = [':', '\\', '/', '?', '*', '[', ']']
    nombre_limpio = nombre
    for char in invalido: nombre_limpio = nombre_limpio.replace(char, '')
    nombre_base = nombre_limpio[:28]
    nombre_final = nombre_base
    contador = 1
    while nombre_final in nombres_usados:
        nombre_final = f"{nombre_base[:26]}({contador})"
        contador += 1
    nombres_usados.add(nombre_final)
    return nombre_final

def procesar_json(json_file):
    data = json.load(json_file)
    res_map = {}
    col_map = {}
    resources = data.get('resources') or []
    nodes = data.get('nodes') or []

    # Mapeo inicial para IDs de recursos y nombres de columnas
    for res in resources:
        r_id = res.get('export_id') or res.get('_id')
        res_map[r_id] = res.get('name', 'Sin Nombre')
        for col in (res.get('columns') or []):
            col_map[col.get('export_id')] = col.get('label') or col.get('name')

    # Lógica de Relaciones
    relaciones = {r.get('export_id'): {"parents": [], "children": []} for r in resources}
    for node in nodes:
        target, sources = node.get('target'), node.get('source')
        if target and sources:
            source_list = sources if isinstance(sources, list) else [sources]
            for s_id in source_list:
                if target in relaciones: relaciones[target]["parents"].append(res_map.get(s_id, f"ID_{s_id}"))
                if s_id in relaciones: relaciones[s_id]["children"].append(res_map.get(target, f"ID_{target}"))

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        nombres_usados = set()
        nombres_hojas_map = [limpiar_nombre_hoja(r.get('name', 'Resource'), nombres_usados) for r in resources]

        # --- HOJA INICIO ---
        ws_inicio = writer.book.create_sheet("Inicio")
        ws_inicio['A1'] = "DOCUMENTACIÓN TÉCNICA SIMETRIK"
        aplicar_estilos_peya(ws_inicio, 1, 1, 1, 5, True)
        
        headers_inicio = ["ID RECURSO", "RECURSO", "TIPO", "ENTRADAS (PARENTS)", "LINK"]
        for ci, h in enumerate(headers_inicio, 1): ws_inicio.cell(5, ci, h)
        aplicar_estilos_peya(ws_inicio, 5, 5, 1, 5, True)
        
        for i, r in enumerate(resources):
            row, e_id = 6 + i, r.get('export_id')
            ws_inicio.cell(row, 1, e_id)
            ws_inicio.cell(row, 2, r.get('name'))
            ws_inicio.cell(row, 3, r.get('resource_type', 'N/A').upper())
            ws_inicio.cell(row, 4, ", ".join(relaciones.get(e_id, {}).get('parents', [])) or "Fuente Primaria")
            
            link = ws_inicio.cell(row, 5, "Ir a Hoja")
            link.hyperlink, link.style = f"#'{nombres_hojas_map[i]}'!A1", "Hyperlink"
            aplicar_estilos_peya(ws_inicio, row, row, 1, 4, False)

        for col_letter in ['A', 'B', 'C', 'D']: ws_inicio.column_dimensions[col_letter].width = 25

        # --- HOJAS DE DETALLE ---
        for i, res in enumerate(resources):
            e_id = res.get('export_id')
            ws = writer.book.create_sheet(nombres_hojas_map[i])
            
            # Info Completa del Recurso
            ws.cell(1, 1, f"RECURSO: {res.get('name')}").font = Font(size=14, bold=True, color=COLOR_PEYA_MAIN)
            ws.cell(2, 1, f"ID: {e_id} | Tipo: {res.get('resource_type', 'N/A').upper()}")
            ws.cell(3, 1, f"Descripción: {res.get('description') or 'Sin descripción'}")
            
            curr = 5
            # Relaciones detalladas
            ws.cell(curr, 1, "CONECTIVIDAD EN EL FLUJO").font = Font(bold=True, color=COLOR_PEYA_MAIN); curr += 1
            ws.cell(curr, 1, "Inputs:"); ws.cell(curr, 2, ", ".join(relaciones[e_id]["parents"]) or "Ninguno"); curr += 1
            ws.cell(curr, 1, "Outputs:"); ws.cell(curr, 2, ", ".join(relaciones[e_id]["children"]) or "Ninguno"); curr += 2

            # Columnas con detección de Buscar V
            ws.cell(curr, 1, "ESTRUCTURA DE COLUMNAS").font = Font(bold=True, color=COLOR_PEYA_MAIN); curr += 1
            headers = ["COLUMNA", "TIPO DATO", "TRANSFORMACIONES / BUSCAR V"]
            for ci, h in enumerate(headers, 1): ws.cell(curr, ci, h)
            aplicar_estilos_peya(ws, curr, curr, 1, 3, True); curr += 1
            
            for col in (res.get('columns') or []):
                trans_list = []
                # Lógica para extraer info de v_lookup
                v_info = col.get('v_lookup')
                if v_info:
                    v_set = v_info.get('v_lookup_set', {})
                    origen_nombre = res_map.get(v_set.get('origin_source_id'), "Desconocido")
                    reglas = v_set.get('rules', [])
                    llaves = [f"[{col_map.get(rule.get('column_a_id'), 'A')} == {col_map.get(rule.get('column_b_id'), 'B')}]" for rule in reglas]
                    trans_list.append(f"🔍 BUSCAR V desde: '{origen_nombre}' | Llaves: {' & '.join(llaves)}")
                
                # Transformaciones normales
                for t in (col.get('transformations') or []):
                    if t.get('query'): trans_list.append(f"⚙️ {t.get('query')}")
                
                vals = [col.get('label'), col.get('data_format'), " \n".join(trans_list)]
                for ci, v in enumerate(vals, 1): ws.cell(curr, ci, v)
                aplicar_estilos_peya(ws, curr, curr, 1, 3, False); curr += 1
            
            ws.column_dimensions['A'].width = 30; ws.column_dimensions['C'].width = 80

        if "Sheet" in writer.book.sheetnames: writer.book.remove(writer.book["Sheet"])
    output.seek(0)
    return output

# --- INTERFAZ STREAMLIT ---
st.markdown(f"<h1 style='color: #{COLOR_PEYA_MAIN};'>Simetrik Doc Generator Pro</h1>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("Sube el JSON de Simetrik", type=['json'])

if uploaded_file and st.button("🚀 Generar Documentación"):
    excel_data = procesar_json(uploaded_file)
    st.download_button("📥 Descargar Excel", excel_data, f"Simetrik_Report_{datetime.now().strftime('%H%M')}.xlsx")
