import streamlit as st
import json
import pandas as pd
import io
import os
import re
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Simetrik Docs Pro | PeYa", page_icon="📊", layout="wide")

C = {
    "red":    "EA0050",
    "white":  "FFFFFF",
    "grey":   "F5F5F5",
    "grey2":  "EBEBEB",
    "dark":   "1C1C1C",
    "border": "D8D8D8",
    "blue":   "1565C0",
    "teal":   "00695C",
    "amber":  "E65100",
    "purple": "4A148C",
    "green":  "1B5E20",
    "slate":  "37474F",
}

RT_LABEL = {
    "native":                  "📥 Fuente",
    "source_union":            "🔗 Unión de Fuentes",
    "source_group":            "📊 Agrupación (Group By)",
    "reconciliation":          "⚖️ Conciliación Estándar",
    "advanced_reconciliation": "🔬 Conciliación Avanzada",
    "consolidation":           "🗂️ Consolidación",
    "resource_join":           "🔀 Join de Recursos",
    "cumulative_balance":      "📈 Balance Acumulado",
}

RT_COLOR = {
    "native":                  C["blue"],
    "source_union":            C["teal"],
    "source_group":            C["amber"],
    "reconciliation":          C["red"],
    "advanced_reconciliation": C["purple"],
    "consolidation":           C["slate"],
    "resource_join":           C["green"],
    "cumulative_balance":      C["green"],
}


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS OPENPYXL
# ══════════════════════════════════════════════════════════════════════════════
def mk_border():
    t = Side(border_style="thin", color=C["border"])
    return Border(left=t, right=t, top=t, bottom=t)

def sc(cell, bg=None, bold=False, color=C["dark"], size=10,
       ha='left', va='top', wrap=True):
    cell.border = mk_border()
    cell.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=wrap)
    cell.font = Font(name='Arial', bold=bold, size=size, color=color)
    if bg:
        cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")

def hdr(cell, text, bg=C["dark"]):
    cell.value = text
    sc(cell, bg=bg, bold=True, color=C["white"], size=10,
       ha='center', va='center', wrap=False)

def section_title(ws, row, text, bg=C["red"], cols=5):
    ws.merge_cells(f'A{row}:{chr(64+cols)}{row}')
    c = ws.cell(row, 1, text)
    sc(c, bg=bg, bold=True, color=C["white"], size=10,
       ha='left', va='center', wrap=False)
    ws.row_dimensions[row].height = 20
    return row + 1

def meta_row(ws, row, label, value, cols=5, bg_val=None):
    bg_val = bg_val or C["grey"]
    c_l = ws.cell(row, 1, label)
    sc(c_l, bg=C["slate"], bold=True, color=C["white"], size=9,
       ha='left', va='center', wrap=False)
    ws.merge_cells(f'B{row}:{chr(64+cols)}{row}')
    c_v = ws.cell(row, 2, str(value) if value is not None else "—")
    sc(c_v, bg=bg_val, size=9, va='center', wrap=False)
    ws.row_dimensions[row].height = 14
    return row + 1


# ══════════════════════════════════════════════════════════════════════════════
# PARSERS
# ══════════════════════════════════════════════════════════════════════════════
def build_maps(data):
    res_map = {}
    col_map = {}
    for r in data.get('resources', []):
        eid = r.get('export_id')
        res_map[eid] = r.get('name', str(eid))
        for c in (r.get('columns') or []):
            cid = c.get('export_id')
            col_map[cid] = c.get('label') or c.get('name') or str(cid)
        sg = r.get('source_group') or {}
        for c in sg.get('columns', []):
            cid = c.get('column_id')
            if cid and cid not in col_map:
                col_map[cid] = f"col_{cid}"
        for v in sg.get('values', []):
            cid = v.get('column_id')
            if cid and cid not in col_map:
                col_map[cid] = f"col_{cid}"
        adv = r.get('advanced_reconciliation') or {}
        for rg in adv.get('reconcilable_groups', []):
            for cs in rg.get('columns_selection', []):
                cid = cs.get('column_id')
                if cid and cid not in col_map:
                    col_map[cid] = f"col_{cid}"
            sc2 = rg.get('segmentation_config') or {}
            ccid = sc2.get('criteria_column_id')
            if ccid and ccid not in col_map:
                col_map[ccid] = f"col_{ccid}"
    return res_map, col_map


def parse_transformation_logic(col, res_map, col_map):
    lines = []
    v = col.get('v_lookup')
    if v:
        vs = v.get('v_lookup_set') or {}
        origin_id = vs.get('origin_source_id')
        origin = res_map.get(origin_id, f"ID:{origin_id}")
        rules = vs.get('rules', [])
        keys = " & ".join(
            f"[A.{col_map.get(r.get('column_a_id'), '?')} = B.{col_map.get(r.get('column_b_id'), '?')}]"
            for r in rules
        )
        lines.append(f"BUSCAR V EN: {origin}")
        if keys:
            lines.append(f"CLAVE MATCH: {keys}")
    parents = [t for t in (col.get('transformations') or []) if t.get('is_parent')]
    for t in parents:
        q = (t.get('query') or '').strip()
        if q and q.upper() != 'N/A':
            lines.append(f"FÓRMULA: {q}")
    return "\n".join(lines) if lines else "Dato directo / heredado"


def parse_segment_filters(segs, col_map):
    result = []
    for seg in (segs or []):
        rules = []
        for fset in (seg.get('segment_filter_sets') or []):
            for r in (fset.get('segment_filter_rules') or []):
                col_name = col_map.get(r.get('column_id'), f"ID:{r.get('column_id')}")
                rules.append(
                    f"{r.get('condition','')} [{col_name}] {r.get('operator','')} {r.get('value','')}".strip()
                )
        if rules:
            result.append({"name": seg.get('name', ''), "rules": rules})
    return result


def parse_reconciliation_rule_sets(rule_sets, col_map, is_advanced=False):
    parsed = []
    for rs in sorted(rule_sets, key=lambda x: x.get('position', 99)):
        rules_desc = []
        for rule in (rs.get('reconciliation_rules') or []):
            col_a = col_map.get(rule.get('column_a_id'), f"ID:{rule.get('column_a_id')}")
            col_b = col_map.get(rule.get('column_b_id'), f"ID:{rule.get('column_b_id')}")
            op    = rule.get('operator', '=')
            tol   = rule.get('tolerance', 0)
            tol_u = rule.get('tolerance_unit') or ''
            tol_s = f"  [tol ±{tol} {tol_u}]" if tol else ""
            rules_desc.append(f"A.{col_a}  {op}  B.{col_b}{tol_s}")
        sweep_info = []
        if is_advanced:
            for sw in (rs.get('sweep_sides') or []):
                prefix = sw.get('prefix_side', '?')
                isr = sw.get('input_sweep_resource') or {}
                seg_meta_id = isr.get('segmentation_metadata_id')
                elem = isr.get('element_to_visualize', '')
                sweep_info.append(
                    f"Lado {prefix} → {'seg_meta:' + str(seg_meta_id) if seg_meta_id else elem}"
                )
        parsed.append({
            "pos":        rs.get('position', 0),
            "name":       rs.get('name', ''),
            "cross_type": rs.get('cross_type', ''),
            "new_ver":    rs.get('is_new_version', False),
            "rules":      rules_desc,
            "sweep":      sweep_info,
        })
    return parsed


def parse_reconcilable_groups(adv, col_map):
    groups = []
    for rg in (adv.get('reconcilable_groups') or []):
        prefix  = rg.get('prefix_side', '?')
        trigger = rg.get('is_trigger', False)
        sc2     = rg.get('segmentation_config') or {}
        crit_id = sc2.get('criteria_column_id')
        crit_col = col_map.get(crit_id, f"ID:{crit_id}") if crit_id else "—"
        segments = [m.get('value', '') for m in sc2.get('segmentation_metadata', []) if m.get('value')]
        col_ids  = [col_map.get(c.get('column_id'), f"ID:{c.get('column_id')}")
                    for c in rg.get('columns_selection', [])]
        groups.append({
            "prefix":   prefix,
            "trigger":  trigger,
            "crit_col": crit_col,
            "segments": segments,
            "cols":     col_ids,
        })
    return groups


def parse_source_group(sg, col_map):
    if not sg:
        return [], []
    group_cols = [col_map.get(c.get('column_id'), f"ID:{c.get('column_id')}")
                  for c in sorted(sg.get('columns', []), key=lambda x: x.get('position', 0))]
    agg_vals   = [(v.get('function', '?'),
                   col_map.get(v.get('column_id'), f"ID:{v.get('column_id')}"))
                  for v in sorted(sg.get('values', []), key=lambda x: x.get('position', 0))]
    return group_cols, agg_vals


def limpiar_hoja(nombre, eid):
    return re.sub(r'[\\/*?:\[\]]', '', str(nombre))[:18] + f"_{eid}"[:13]


def build_relations(resources, nodes, res_map):
    all_ids = {r.get('export_id') for r in resources}
    rels = {r.get('export_id'): {"parents": [], "children": []} for r in resources}
    for n in nodes:
        t_id  = n.get('target')
        s_val = n.get('source')
        if not (t_id and s_val):
            continue
        s_list = s_val if isinstance(s_val, list) else [s_val]
        for sid in s_list:
            label = res_map.get(sid, str(sid))
            ext   = "" if sid in all_ids else " ↗"
            if t_id in rels:
                rels[t_id]["parents"].append(label + ext)
            if sid in rels:
                rels[sid]["children"].append(res_map.get(t_id, str(t_id)) +
                                              ("" if t_id in all_ids else " ↗"))
    return rels


# ══════════════════════════════════════════════════════════════════════════════
# GENERADOR EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def generar_excel(data, selected_ids):
    all_resources = data.get('resources', [])
    nodes         = data.get('nodes', [])
    res_map, col_map = build_maps(data)

    seen, resources = set(), []
    for r in all_resources:
        eid = r.get('export_id')
        if eid in selected_ids and eid not in seen:
            seen.add(eid)
            resources.append(r)

    rels      = build_relations(resources, nodes, res_map)
    map_hojas = {r.get('export_id'): limpiar_hoja(r.get('name', ''), r.get('export_id'))
                 for r in resources}

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        wb = writer.book

        # ── ÍNDICE ─────────────────────────────────────────────────────────────
        ws = wb.create_sheet("📚 Índice", 0)
        ws.sheet_view.showGridLines = False

        ws.merge_cells('A1:G1')
        c = ws.cell(1, 1, "SIMETRIK DOCUMENTATION PRO  ·  PeYa Finance Operations & Control")
        sc(c, bg=C["red"], bold=True, color=C["white"], size=13,
           ha='center', va='center', wrap=False)
        ws.row_dimensions[1].height = 32

        ws.merge_cells('A2:G2')
        c = ws.cell(2, 1,
            f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}   |   "
            f"Recursos documentados: {len(resources)}")
        sc(c, bg=C["dark"], color=C["white"], size=9, ha='center', va='center', wrap=False)
        ws.row_dimensions[2].height = 15

        for i, h in enumerate(["ID", "NOMBRE RECURSO", "TIPO", "PROVIENE DE",
                                "ALIMENTA A", "ENCADENADA", "LINK"], 1):
            hdr(ws.cell(4, i, h), h, bg=C["dark"])
        ws.row_dimensions[4].height = 20

        for row_n, res in enumerate(resources, 5):
            eid = res.get('export_id')
            rt  = res.get('resource_type', '')
            recon = res.get('reconciliation') or {}
            adv   = res.get('advanced_reconciliation') or {}
            chained = recon.get('is_chained', False) or adv.get('is_chained', False)
            bg = C["grey"] if row_n % 2 == 0 else C["white"]

            for col_n, val in enumerate([
                eid,
                res.get('name', ''),
                RT_LABEL.get(rt, rt),
                ", ".join(rels[eid]["parents"]) or "— origen",
                ", ".join(rels[eid]["children"]) or "— fin de flujo",
                "Sí" if chained else "No",
            ], 1):
                c = ws.cell(row_n, col_n, val)
                sc(c, bg=bg, size=9, va='center', wrap=False)
                if col_n == 3:
                    c.font = Font(name='Arial', bold=True, size=9,
                                  color=RT_COLOR.get(rt, C["dark"]))
                c.border = mk_border()

            lnk = ws.cell(row_n, 7, "Ver →")
            lnk.hyperlink = f"#'{map_hojas[eid]}'!A1"
            lnk.font = Font(name='Arial', color="0D47A1", underline="single", size=9)
            lnk.border = mk_border()
            ws.row_dimensions[row_n].height = 15

        ws.column_dimensions['A'].width = 11
        ws.column_dimensions['B'].width = 46
        ws.column_dimensions['C'].width = 26
        ws.column_dimensions['D'].width = 38
        ws.column_dimensions['E'].width = 38
        ws.column_dimensions['F'].width = 12
        ws.column_dimensions['G'].width = 8

        # ── HOJAS DE DETALLE ───────────────────────────────────────────────────
        for res in resources:
            eid  = res.get('export_id')
            rt   = res.get('resource_type', '')
            name = res.get('name', '')
            tc   = RT_COLOR.get(rt, C["dark"])
            COLS = 5

            ws = wb.create_sheet(map_hojas[eid])
            ws.sheet_view.showGridLines = False

            row = 1
            ws.merge_cells(f'A{row}:E{row}')
            c = ws.cell(row, 1, f"{RT_LABEL.get(rt, '')}  ·  {name}")
            sc(c, bg=tc, bold=True, color=C["white"], size=12,
               ha='left', va='center', wrap=False)
            ws.row_dimensions[row].height = 30
            row += 1

            row = meta_row(ws, row, "ID Recurso",  eid,  cols=COLS)
            row = meta_row(ws, row, "Tipo",        RT_LABEL.get(rt, rt), cols=COLS)
            row = meta_row(ws, row, "Proviene de",
                           ", ".join(rels[eid]["parents"]) or "Origen", cols=COLS)
            row = meta_row(ws, row, "Alimenta a",
                           ", ".join(rels[eid]["children"]) or "Fin de flujo", cols=COLS)
            row += 1

            # ── CONCILIACIÓN ESTÁNDAR ──────────────────────────────────────────
            recon = res.get('reconciliation')
            if recon:
                row = section_title(ws, row, "⚖️  REGLAS DE CONCILIACIÓN ESTÁNDAR",
                                    bg=C["red"], cols=COLS)
                a_cfg = recon.get('a_source_settings') or {}
                b_cfg = recon.get('b_source_settings') or {}
                src_a = (f"[{recon.get('segment_a_prefix','A')}] "
                         f"{res_map.get(a_cfg.get('resource_id'), '—')}")
                src_b = (f"[{recon.get('segment_b_prefix','B')}] "
                         f"{res_map.get(b_cfg.get('resource_id'), '—')}")
                row = meta_row(ws, row, "Fuente A (no trigger)", src_a,
                               cols=COLS, bg_val="FFF8F8")
                row = meta_row(ws, row, "Fuente B (trigger)",    src_b,
                               cols=COLS, bg_val="FFF8F8")
                row = meta_row(ws, row, "Conciliación encadenada",
                               "Sí" if recon.get('is_chained') else "No", cols=COLS)
                row += 1

                for col_n, h in enumerate(
                        ["POS.", "NOMBRE DEL RULE SET", "REGLAS DE MATCHING  (A vs B)"], 1):
                    hdr(ws.cell(row, col_n, h), h, bg=C["red"])
                ws.merge_cells(f'C{row}:E{row}')
                ws.row_dimensions[row].height = 18
                row += 1

                for i, rs in enumerate(parse_reconciliation_rule_sets(
                        recon.get('reconciliation_rule_sets', []), col_map)):
                    bg = C["grey"] if i % 2 == 0 else C["white"]
                    c1 = ws.cell(row, 1, rs["pos"])
                    c2 = ws.cell(row, 2, rs["name"])
                    ws.merge_cells(f'C{row}:E{row}')
                    c3 = ws.cell(row, 3, "\n".join(rs["rules"]))
                    for c, al in [(c1,'center'),(c2,'left'),(c3,'left')]:
                        sc(c, bg=bg, size=9, va='top', wrap=True, ha=al)
                    ws.row_dimensions[row].height = max(14, len(rs["rules"]) * 13)
                    row += 1
                row += 1

            # ── CONCILIACIÓN AVANZADA ──────────────────────────────────────────
            adv = res.get('advanced_reconciliation')
            if adv:
                row = section_title(ws, row, "🔬  REGLAS DE CONCILIACIÓN AVANZADA",
                                    bg=C["purple"], cols=COLS)

                # Grupos / lados A y B
                groups = parse_reconcilable_groups(adv, col_map)
                for g in groups:
                    trig_lbl = "  [TRIGGER]" if g["trigger"] else ""
                    row = section_title(ws, row,
                        f"  Grupo Lado {g['prefix']}{trig_lbl}  —  "
                        f"columna criterio: {g['crit_col']}",
                        bg="6A1B9A", cols=COLS)
                    if g["segments"]:
                        row = meta_row(ws, row, "Segmentos internos",
                                       " | ".join(g["segments"]),
                                       cols=COLS, bg_val="F3E5F5")
                    if g["cols"]:
                        row = meta_row(ws, row, "Columnas clave seleccionadas",
                                       ", ".join(g["cols"]),
                                       cols=COLS, bg_val="F3E5F5")
                row += 1

                # Rule sets avanzados
                for col_n, h in enumerate(
                        ["POS.", "NOMBRE / TIPO CRUCE", "REGLAS (A vs B)", "SWEEP SIDES"], 1):
                    hdr(ws.cell(row, col_n, h), h, bg=C["purple"])
                ws.merge_cells(f'C{row}:D{row}')
                ws.row_dimensions[row].height = 18
                row += 1

                for i, rs in enumerate(parse_reconciliation_rule_sets(
                        adv.get('reconciliation_rule_sets', []), col_map, is_advanced=True)):
                    bg = C["grey"] if i % 2 == 0 else C["white"]
                    name_text  = rs["name"] + (f"\n[{rs['cross_type']}]" if rs["cross_type"] else "")
                    name_text += "  ✦ new version" if rs["new_ver"] else ""
                    c1 = ws.cell(row, 1, rs["pos"])
                    c2 = ws.cell(row, 2, name_text)
                    ws.merge_cells(f'C{row}:D{row}')
                    c3 = ws.cell(row, 3, "\n".join(rs["rules"]))
                    c4 = ws.cell(row, 5, "\n".join(rs["sweep"]))
                    for c, al in [(c1,'center'),(c2,'left'),(c3,'left'),(c4,'left')]:
                        sc(c, bg=bg, size=9, va='top', wrap=True, ha=al)
                    n_lines = max(len(rs["rules"]), len(rs["sweep"]), 1)
                    ws.row_dimensions[row].height = max(14, n_lines * 13)
                    row += 1
                row += 1

            # ── SOURCE GROUP ──────────────────────────────────────────────────
            sg = res.get('source_group')
            if sg:
                row = section_title(ws, row,
                    "📊  CONFIGURACIÓN DE AGRUPACIÓN (GROUP BY / TABLA DINÁMICA)",
                    bg=C["amber"], cols=COLS)
                group_cols, agg_vals = parse_source_group(sg, col_map)
                row = meta_row(ws, row, "GROUP BY (dimensiones)",
                               " | ".join(group_cols) or "—",
                               cols=COLS, bg_val="FFF3E0")
                agg_str = "  |  ".join(f"{fn}( {col} )" for fn, col in agg_vals)
                row = meta_row(ws, row, "Agregaciones (métricas)",
                               agg_str or "—", cols=COLS, bg_val="FFF3E0")
                row = meta_row(ws, row, "Acumulativo",
                               "Sí" if sg.get('is_accumulative') else "No", cols=COLS)
                row += 1

            # ── SOURCE UNION ──────────────────────────────────────────────────
            su = res.get('source_union')
            if su:
                row = section_title(ws, row,
                    "🔗  CONFIGURACIÓN DE UNIÓN DE FUENTES",
                    bg=C["teal"], cols=COLS)
                for us in (su.get('union_segments') or []):
                    label = "TRIGGER" if us.get('is_trigger') else "No trigger"
                    ttype = us.get('trigger_type') or ''
                    row = meta_row(ws, row,
                                   f"Segmento ID {us.get('segment_id', '')}",
                                   f"{label}  {'· ' + ttype if ttype else ''}",
                                   cols=COLS, bg_val="E0F2F1")
                row += 1

            # ── SEGMENTOS / FILTROS ───────────────────────────────────────────
            segs = parse_segment_filters(res.get('segments', []), col_map)
            if segs:
                row = section_title(ws, row, "🔍  SEGMENTOS Y FILTROS CONFIGURADOS",
                                    bg=C["slate"], cols=COLS)
                for col_n, h in enumerate(["SEGMENTO", "FILTROS APLICADOS"], 1):
                    hdr(ws.cell(row, col_n, h), h, bg=C["slate"])
                ws.merge_cells(f'B{row}:E{row}')
                ws.row_dimensions[row].height = 18
                row += 1
                for i, seg in enumerate(segs):
                    bg = C["grey"] if i % 2 == 0 else C["white"]
                    c1 = ws.cell(row, 1, seg["name"])
                    ws.merge_cells(f'B{row}:E{row}')
                    c2 = ws.cell(row, 2, "\n".join(seg["rules"]))
                    for c in [c1, c2]:
                        sc(c, bg=bg, size=9, va='top', wrap=True)
                    ws.row_dimensions[row].height = max(14, len(seg["rules"]) * 12)
                    row += 1
                row += 1

            # ── COLUMNAS ──────────────────────────────────────────────────────
            columns = sorted(res.get('columns') or [],
                             key=lambda x: x.get('position', 0))
            if columns:
                row = section_title(ws, row, "📋  CONFIGURACIÓN DE COLUMNAS",
                                    bg=C["blue"], cols=COLS)
                for col_n, h in enumerate(
                        ["LABEL / NOMBRE", "TIPO DATO", "TIPO COL.", "LÓGICA · FÓRMULA · BUSCAR V"], 1):
                    hdr(ws.cell(row, col_n, h), h, bg=C["blue"])
                ws.merge_cells(f'D{row}:E{row}')
                ws.row_dimensions[row].height = 18
                row += 1

                for i, col in enumerate(columns):
                    label    = col.get('label') or col.get('name', '')
                    dtype    = col.get('data_format', '')
                    col_type = (col.get('column_type') or '').replace('_', ' ').upper()
                    logic    = parse_transformation_logic(col, res_map, col_map)
                    bg       = C["grey"] if i % 2 == 0 else C["white"]
                    c1 = ws.cell(row, 1, label)
                    c2 = ws.cell(row, 2, dtype)
                    c3 = ws.cell(row, 3, col_type)
                    ws.merge_cells(f'D{row}:E{row}')
                    c4 = ws.cell(row, 4, logic)
                    for c, al in [(c1,'left'),(c2,'center'),(c3,'center'),(c4,'left')]:
                        sc(c, bg=bg, size=9, va='top', wrap=True, ha=al)
                    ws.row_dimensions[row].height = max(14, (logic.count('\n') + 1) * 12)
                    row += 1

            ws.column_dimensions['A'].width = 32
            ws.column_dimensions['B'].width = 14
            ws.column_dimensions['C'].width = 16
            ws.column_dimensions['D'].width = 55
            ws.column_dimensions['E'].width = 18

        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

    output.seek(0)
    return output


# ══════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div style='background:linear-gradient(135deg,#EA0050 0%,#B0003A 100%);
    padding:28px 36px;border-radius:16px;
    box-shadow:0 6px 24px rgba(234,0,80,0.25);margin-bottom:4px'>
    <h1 style='color:white;margin:0;font-family:Arial,sans-serif;
        font-size:2rem;letter-spacing:-0.5px;font-weight:700'>
        📊 Simetrik Documentation Pro
    </h1>
    <p style='color:rgba(255,255,255,0.88);margin:8px 0 0;
        font-size:1rem;font-family:Arial'>
        PeYa Finance Operations &amp; Control &nbsp;·&nbsp;
        Generador de documentación técnica de flujos de conciliación
    </p>
</div>""", unsafe_allow_html=True)

st.write("")
up = st.file_uploader(
    "Subí el JSON exportado desde Simetrik",
    type=['json'],
    help="En Simetrik: Flujo → ⚙️ Configuración → Exportar JSON"
)

if not up:
    st.info("👆 Subí un JSON para comenzar.")
    st.stop()

try:
    data = json.load(up)
except Exception as e:
    st.error(f"Error al leer el JSON: {e}")
    st.stop()

all_resources = data.get('resources', [])
nodes         = data.get('nodes', [])

# Deduplicar
seen_load, resources_unique = set(), []
for r in all_resources:
    eid = r.get('export_id')
    if eid not in seen_load:
        seen_load.add(eid)
        resources_unique.append(r)

res_map, col_map = build_maps(data)
rels_all = build_relations(resources_unique, nodes, res_map)

# ── PASO 1: SELECCIÓN ─────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("### 1️⃣  Seleccioná los recursos a documentar")

c1, c2 = st.columns([3, 1])
c1.caption(f"El JSON contiene **{len(resources_unique)}** recursos únicos.")

all_types = sorted({r.get('resource_type', '') for r in resources_unique})
filtro_tipo = c2.multiselect(
    "Filtrar tipo",
    options=all_types,
    format_func=lambda x: RT_LABEL.get(x, x),
    default=all_types,
    label_visibility="collapsed"
)

resources_visible = [r for r in resources_unique
                     if r.get('resource_type', '') in filtro_tipo]

# Botones globales
bc1, bc2, _ = st.columns([1, 1, 8])
select_all   = bc1.button("✅ Todos")
deselect_all = bc2.button("⬜ Ninguno")

if 'sel' not in st.session_state:
    st.session_state.sel = {r.get('export_id'): True for r in resources_unique}
if select_all:
    for r in resources_visible:
        st.session_state.sel[r.get('export_id')] = True
if deselect_all:
    for r in resources_visible:
        st.session_state.sel[r.get('export_id')] = False

# Agrupar por tipo
tipo_groups: dict = {}
for r in resources_visible:
    tipo_groups.setdefault(r.get('resource_type', ''), []).append(r)

selected_ids = set()
for rt, group in tipo_groups.items():
    color = RT_COLOR.get(rt, C["dark"])
    st.markdown(
        f"<div style='background:#{color};color:white;padding:5px 14px;"
        f"border-radius:8px;font-size:0.85rem;font-weight:700;"
        f"margin:10px 0 4px'>{RT_LABEL.get(rt, rt)}  ({len(group)})</div>",
        unsafe_allow_html=True
    )
    for r in group:
        eid   = r.get('export_id')
        name  = r.get('name', '')
        pars  = ", ".join(rels_all[eid]["parents"]) or "—"
        chils = ", ".join(rels_all[eid]["children"]) or "—"

        ca, cb, cc, cd = st.columns([0.35, 3, 2.5, 2.5])
        checked = ca.checkbox(
            "", value=st.session_state.sel.get(eid, True), key=f"chk_{eid}"
        )
        st.session_state.sel[eid] = checked
        cb.markdown(f"**{name}**  `{eid}`")
        cc.caption(f"⬅️ {pars}")
        cd.caption(f"➡️ {chils}")
        if checked:
            selected_ids.add(eid)

# ── PASO 2: GENERAR ───────────────────────────────────────────────────────────
st.markdown("---")
n_sel = len(selected_ids)
st.markdown(f"### 2️⃣  Generar Excel  ·  **{n_sel}** recurso{'s' if n_sel != 1 else ''} seleccionado{'s' if n_sel != 1 else ''}")

if not selected_ids:
    st.warning("Seleccioná al menos un recurso para continuar.")
    st.stop()

nombre_dl = f"{os.path.splitext(up.name)[0]}_DOC_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"

if st.button("🚀  GENERAR EXCEL PROFESIONAL", type="primary", use_container_width=True):
    with st.spinner("Procesando… construyendo reglas de conciliación, segmentos y columnas"):
        try:
            excel_bytes = generar_excel(data, selected_ids)
            st.success(f"✅ ¡Listo! {n_sel} recursos documentados.")
            st.balloons()
            st.download_button(
                label="📥  Descargar Excel",
                data=excel_bytes,
                file_name=nombre_dl,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
        except Exception as e:
            st.error(f"Error generando el Excel: {e}")
            import traceback
            st.code(traceback.format_exc())

st.markdown("---")
st.caption("Simetrik Documentation Pro · PeYa Finance · v2.0")
