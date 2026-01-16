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

# --- FUNCIONES DE ESTILO Y MAPEO (Lógica V5) ---
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
                if 'null' in op: set_logic.append(f"[{col_name} {op}]")
                else: set_logic.append(f"[{col_name} {op} {val}]")
            if set_logic: texto_logica.append(f"({f' {condicion_set} '.join(set_logic)})")
    except: return "Error en filtros"
    return " OR ".join(texto_logica) if texto_logica else "Sin filtros (Trae todo)"

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
    # Leer JSON desde el archivo subido
    data = json.load(json_file)
    
    # Mapeos
    res_map = {}
    col_map = {}
    seg_map = {}
    for res in (data.get('resources') or []):
        res_id = res.get('export_id') or res.get('_id')
        res_name = res.get('name', 'Sin Nombre')
        if res_id: res_map[res_id] = res_name
        for col in (res.get('columns') or []):
            c_id = col.get('export_id')
            label = col.get('label') or col.get('name')
            if c_id: col_map[c_id] = label
        if res.get('resource_type') == 'source_union':
            union_data = res.get('source_union') or {}
            for u_col in (union_data.get('union_columns') or []):
                uc_id = u_col.get('union_column_id')
                if uc_id: col_map[uc_id] = f"UnionCol_{uc_id}"
        for seg in (res.get('segments') or []):
            s_id = seg.get('export_id')
            s_name = seg.get('name')
            if s_id: seg_map[s_id] = s_name

    resources = data.get('resources') or []
    nodes = data.get('nodes') or []

    # Buffer en memoria para el Excel
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        nombres_usados = set()
        nombres_hojas_map = [limpiar_nombre_hoja(r.get('name', 'Resource'), nombres_usados) for r in resources]

        # HOJA INICIO
        ws_inicio = writer.book.create_sheet("Inicio")
        ws_inicio['A1'] = "RESUMEN EJECUTIVO"
        aplicar_estilos_peya(ws_inicio, 1, 1, 1, 2, True)
        
        resumen = [("Nodos del Flujo", len(nodes)), ("Total Recursos", len(resources))]
        for i, (k, v) in enumerate(resumen, 2):
            ws_inicio.cell(i, 1, k); ws_inicio.cell(i, 2, v)
            aplicar_estilos_peya(ws_inicio, i, i, 1, 2, False)
        
        start_idx = 5
        ws_inicio.cell(start_idx, 1, "ÍNDICE DE RECURSOS"); ws_inicio.cell(start_idx, 2, "LINK")
        aplicar_estilos_peya(ws_inicio, start_idx, start_idx, 1, 2, True)
        
        for i, r in enumerate(resources):
            row = start_idx + 1 + i
            ws_inicio.cell(row, 1, r.get('name'))
            link = ws_inicio.cell(row, 2, "Ir a Hoja")
            link.hyperlink = f"#'{nombres_hojas_map[i]}'!A1"
            link.style = "Hyperlink"
            aplicar_estilos_peya(ws_inicio, row, row, 1, 1, False)
        ws_inicio.column_dimensions['A'].width = 50

        # HOJAS DE RECURSOS
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, res in enumerate(resources):
            # Barra de progreso visual
            progreso = int((i + 1) / len(resources) * 100)
            progress_bar.progress(progreso)
            status_text.text(f"Procesando: {res.get('name')}...")

            try:
                sheet_name = nombres_hojas_map[i]
                ws = writer.book.create_sheet(sheet_name)
                curr = 1
                
                ws.merge_cells(start_row=curr, start_column=1, end_row=curr, end_column=5)
                title = ws.cell(curr, 1, f"RECURSO: {res.get('name')}")
                title.font = Font(size=14, bold=True, color=COLOR_PEYA_MAIN)
                curr += 2

                # Conciliación
                if res.get('resource_type') == 'reconciliation':
                    rec = res.get('reconciliation') or {}
                    headers = ["LADO", "FUENTE", "GRUPO"]
                    for ci, h in enumerate(headers, 1): ws.cell(curr, ci, h)
                    aplicar_estilos_peya(ws, curr, curr, 1, 3, True); curr += 1
                    
                    rows = [
                        ["LADO A", res_map.get(rec.get('a_source_settings', {}).get('resource_id'), "N/A"), seg_map.get(rec.get('segment_a_id'), "Todos")],
                        ["LADO B", res_map.get(rec.get('b_source_settings', {}).get('resource_id'), "N/A"), seg_map.get(rec.get('segment_b_id'), "Todos")]
                    ]
                    for r in rows:
                        for ci, v in enumerate(r, 1): ws.cell(curr, ci, v)
                        aplicar_estilos_peya(ws, curr, curr, 1, 3, False); curr += 1
                    curr += 2

                    rule_sets = rec.get('reconciliation_rule_sets') or []
                    if rule_sets:
                        headers = ["GRUPO", "A", "OP", "B", "TOLERANCIA"]
                        for ci, h in enumerate(headers, 1): ws.cell(curr, ci, h)
                        aplicar_estilos_peya(ws, curr, curr, 1, 5, True); curr += 1
                        for rs in rule_sets:
                            for rule in (rs.get('reconciliation_rules') or []):
                                ca = col_map.get(rule.get('column_a_id'), str(rule.get('column_a_id'))).split('->')[-1].strip()
                                cb = col_map.get(rule.get('column_b_id'), str(rule.get('column_b_id'))).split('->')[-1].strip()
                                tol = f"{rule.get('tolerance')} {rule.get('tolerance_unit')}" if rule.get('tolerance') else "Exacto"
                                for ci, v in enumerate([rs.get('name'), ca, rule.get('operator'), cb, tol], 1): ws.cell(curr, ci, v)
                                aplicar_estilos_peya(ws, curr, curr, 1, 5, False); curr += 1
                        curr += 2

                # Segmentos
                segments = res.get('segments') or []
                if segments:
                    ws.cell(curr, 1, "GRUPOS").font = Font(bold=True, color=COLOR_PEYA_MAIN); curr += 1
                    headers = ["NOMBRE", "LÓGICA"]
                    for ci, h in enumerate(headers, 1): ws.cell(curr, ci, h)
                    aplicar_estilos_peya(ws, curr, curr, 1, 2, True); curr += 1
                    for seg in segments:
                        ws.cell(curr, 1, seg.get('name'))
                        ws.cell(curr, 2, formatear_reglas_segmento(seg, col_map))
                        aplicar_estilos_peya(ws, curr, curr, 1, 2, False); curr += 1
                    curr += 2

                # Columnas
                headers = ["COLUMNA", "TIPO", "TRANSFORMACIÓN"]
                for ci, h in enumerate(headers, 1): ws.cell(curr, ci, h)
                aplicar_estilos_peya(ws, curr, curr, 1, 3, True); curr += 1
                for col in (res.get('columns') or []):
                    tr = " | ".join([t.get('query','') for t in (col.get('transformations') or []) if t.get('query')])
                    vals = [col.get('label'), col.get('data_format'), tr]
                    for ci, v in enumerate(vals, 1): ws.cell(curr, ci, v)
                    aplicar_estilos_peya(ws, curr, curr, 1, 3, False); curr += 1
                
                ws.column_dimensions['A'].width = 35; ws.column_dimensions['C'].width = 50

            except Exception as e:
                st.error(f"Error en recurso {res.get('name')}: {e}")

        status_text.text("¡Procesamiento completado!")
        
    output.seek(0)
    return output

# --- INTERFAZ DE USUARIO ---
st.markdown(f"""
    <h1 style='color: #{COLOR_PEYA_MAIN};'>Generador de Documentación Simetrik</h1>
    <p>Sube el archivo JSON descargado de Simetrik para generar el reporte técnico.</p>
    """, unsafe_allow_html=True)

uploaded_file = st.file_uploader("Arrastra tu archivo JSON aquí", type=['json'])

if uploaded_file is not None:
    if st.button("🚀 Generar Documentación Excel"):
        try:
            excel_data = procesar_json(uploaded_file)
            
            # Nombre dinámico
            timestamp = datetime.now().strftime("%H%M%S")
            file_name = f"Reporte_Simetrik_PeYa_{timestamp}.xlsx"
            
            st.success("¡Archivo generado con éxito! Descárgalo abajo:")
            
            st.download_button(
                label="📥 Descargar Reporte Excel",
                data=excel_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Ocurrió un error crítico: {e}")

else:
    st.info("👆 Esperando archivo...")