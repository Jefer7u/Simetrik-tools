import streamlit as st
import json
import pandas as pd
import io
import os
import re
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

st.set_page_config(page_title="Simetrik Docs  | PeYa", page_icon="🛵📄", layout="wide")

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════════════════════
C = {
    "red":    "EA0050", "white": "FFFFFF", "grey":  "F5F5F5",
    "dark":   "1C1C1C", "border":"D8D8D8", "blue":  "1565C0",
    "teal":   "00695C", "amber": "E65100", "purple":"4A148C",
    "green":  "1B5E20", "slate": "37474F", "rose":  "880E4F",
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

# Orden de tipos para sorting
RT_ORDER = {
    "native": 1, "source_union": 2, "source_group": 3,
    "reconciliation": 4, "advanced_reconciliation": 5,
    "consolidation": 6, "resource_join": 7, "cumulative_balance": 8,
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
    cell.font = Font(name='Calibri', bold=bold, size=size, color=color)
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
    sc(c_v, bg=bg_val, size=9, va='center', wrap=True)
    ws.row_dimensions[row].height = 14
    return row + 1

def row_height(n_lines, base=13):
    return max(14, n_lines * base)

# ══════════════════════════════════════════════════════════════════════════════
# PARSERS
# ══════════════════════════════════════════════════════════════════════════════
def build_maps(data):
    """Construye mapas globales: res_map, col_map, seg_map, meta_map, seg_usage."""
    res_map   = {}   # export_id → nombre del recurso
    col_map   = {}   # export_id → label de columna
    seg_map   = {}   # export_id → {name, resource, resource_id, rules}
    meta_map  = {}   # export_id → valor del segmento (avanzadas)
    seg_usage = {}   # seg export_id → [(resource_name, tipo_uso)]

    for r in data.get('resources', []):
        eid  = r.get('export_id')
        res_map[eid] = r.get('name', str(eid))

        # Columnas del recurso
        for c in (r.get('columns') or []):
            cid = c.get('export_id')
            col_map[cid] = c.get('label') or c.get('name') or str(cid)

        # Columnas de source_group
        sg = r.get('source_group') or {}
        for c in sg.get('columns', []):
            cid = c.get('column_id')
            if cid and cid not in col_map:
                col_map[cid] = f"col_{cid}"
        for v in sg.get('values', []):
            cid = v.get('column_id')
            if cid and cid not in col_map:
                col_map[cid] = f"col_{cid}"

        # Columnas de reconcilable_groups (avanzadas)
        adv = r.get('advanced_reconciliation') or {}
        for rg in adv.get('reconcilable_groups', []):
            for cs in rg.get('columns_selection', []):
                cid = cs.get('column_id')
                if cid and cid not in col_map:
                    col_map[cid] = f"col_{cid}"
            sc2 = rg.get('segmentation_config') or {}
            for m in sc2.get('segmentation_metadata', []):
                meta_map[m.get('export_id')] = m.get('value', '?')
            ccid = sc2.get('criteria_column_id')
            if ccid and ccid not in col_map:
                col_map[ccid] = f"col_{ccid}"

        # Grupos conciliables (segments)
        for seg in (r.get('segments') or []):
            rules = []
            for fset in (seg.get('segment_filter_sets') or []):
                for rule in (fset.get('segment_filter_rules') or []):
                    rules.append(rule)
            seg_map[seg.get('export_id')] = {
                'name':        seg.get('name', ''),
                'resource':    r.get('name', ''),
                'resource_id': eid,
                'default':     seg.get('default_segment', False),
                'rules':       rules,
            }

    # ── Mapa de uso: qué recursos usan cada grupo conciliable ────────────────
    for r in data.get('resources', []):
        rname = r.get('name', '')

        recon = r.get('reconciliation') or {}
        sa = recon.get('segment_a_id')
        sb = recon.get('segment_b_id')
        pa = recon.get('segment_a_prefix', 'A')
        pb = recon.get('segment_b_prefix', 'B')
        if sa:
            seg_usage.setdefault(sa, []).append((rname, f"Conciliacion lado {pa}"))
        if sb:
            seg_usage.setdefault(sb, []).append((rname, f"Conciliacion lado {pb}"))

        adv = r.get('advanced_reconciliation') or {}
        for rg in adv.get('reconcilable_groups', []):
            sgid = rg.get('segment_id')
            if sgid:
                seg_usage.setdefault(sgid, []).append(
                    (rname, f"Conciliacion Avanzada lado {rg.get('prefix_side','?')}")
                )

        su = r.get('source_union') or {}
        for us in su.get('union_segments', []):
            sgid = us.get('segment_id')
            if sgid:
                seg_usage.setdefault(sgid, []).append((rname, "Union de Fuentes"))

    return res_map, col_map, seg_map, meta_map, seg_usage


def fmt_filter_rules(rules, col_map):
    """Convierte lista de reglas de filtro a texto legible."""
    lines = []
    for r in rules:
        col_name = col_map.get(r.get('column_id'), f"ID:{r.get('column_id')}")
        cond = r.get('condition', '')
        op   = r.get('operator', '')
        val  = r.get('value', '')
        lines.append(f"{cond} [{col_name}] {op} {val}".strip())
    return "\n".join(lines) if lines else "Sin filtros configurados"


def parse_transformation_logic(col, res_map, col_map):
    lines = []

    # ── Columnas de duplicados (generated_type = uniqueness) ──────────────────
    uniq = col.get('uniqueness')
    if uniq:
        dtype      = col.get('data_format', '')
        uniq_type  = uniq.get('type')
        order_keys = uniq.get('order_keys', [])
        part_keys  = uniq.get('partition_keys', [])

        # Tipo de columna de duplicado
        if dtype == 'boolean':
            lines.append("TIPO: Booleano de duplicado")
        elif dtype == 'integer':
            lines.append("TIPO: Numeracion de duplicado")

        # Registro que se conserva
        # ORDER BY (columna y dirección)
        if order_keys:
            order_parts = []
            for ok in sorted(order_keys, key=lambda x: x.get('position', 0)):
                col_name  = col_map.get(ok.get('column_id'), f"ID:{ok.get('column_id')}")
                direction = "ASC" if ok.get('order_by', 1) == 1 else "DESC"
                order_parts.append(f"{col_name} {direction}")
            lines.append("ORDER BY: " + ", ".join(order_parts))

        # PARTITION BY (clave que define el duplicado)
        if part_keys:
            part_names = [col_map.get(pk.get('column_id'), f"ID:{pk.get('column_id')}")
                          for pk in part_keys]
            lines.append("PARTITION BY (clave de duplicado):\n  " + "\n  ".join(part_names))

        return "\n".join(lines)

    # ── BUSCAR V ──────────────────────────────────────────────────────────────
    v = col.get('v_lookup')
    if v:
        vs = v.get('v_lookup_set') or {}
        origin_id = vs.get('origin_source_id')
        origin = res_map.get(origin_id, f"ID:{origin_id}")
        rules = vs.get('rules', [])
        keys = " & ".join(
            "A." + col_map.get(r.get('column_a_id'), '?') +
            " = B." + col_map.get(r.get('column_b_id'), '?')
            for r in rules
        )
        lines.append("BUSCAR V EN: " + origin)
        if keys:
            lines.append("CLAVE MATCH: " + keys)

    # ── Transformaciones / fórmulas ───────────────────────────────────────────
    parents = [t for t in (col.get('transformations') or []) if t.get('is_parent')]
    for t in parents:
        q = (t.get('query') or '').strip()
        if q and q.upper() != 'N/A':
            lines.append("FÓRMULA: " + q)

    return "\n".join(lines) if lines else "Campo directo / heredado"


def parse_std_reconciliation(recon, res_map, col_map, seg_map):
    """Parsea una conciliación estándar incluyendo grupos conciliables activos."""
    if not recon:
        return None

    sa_id = recon.get('segment_a_id')
    sb_id = recon.get('segment_b_id')
    a_cfg = recon.get('a_source_settings') or {}
    b_cfg = recon.get('b_source_settings') or {}

    def resolve_side(cfg, seg_id, prefix):
        resource_name = res_map.get(cfg.get('resource_id'), '—')
        seg = seg_map.get(seg_id) or {}
        seg_name = seg.get('name', f"ID:{seg_id}")
        seg_rules = fmt_filter_rules(seg.get('rules', []), col_map)
        return {
            'prefix':        prefix,
            'resource_name': resource_name,
            'group_name':    seg_name,
            'group_filters': seg_rules,
            'is_trigger':    cfg.get('is_trigger', False),
        }

    sides = [
        resolve_side(a_cfg, sa_id, recon.get('segment_a_prefix', 'A')),
        resolve_side(b_cfg, sb_id, recon.get('segment_b_prefix', 'B')),
    ]

    rule_sets = []
    for rs in sorted(recon.get('reconciliation_rule_sets', []),
                     key=lambda x: x.get('position', 99)):
        rules_desc = []
        for rule in (rs.get('reconciliation_rules') or []):
            col_a = col_map.get(rule.get('column_a_id'), f"ID:{rule.get('column_a_id')}")
            col_b = col_map.get(rule.get('column_b_id'), f"ID:{rule.get('column_b_id')}")
            op    = rule.get('operator', '=')
            tol   = rule.get('tolerance', 0)
            tol_u = rule.get('tolerance_unit') or ''
            tol_s = f"  [tolerancia ±{tol} {tol_u}]" if tol else ""
            rules_desc.append(f"A.{col_a}  {op}  B.{col_b}{tol_s}")
        rule_sets.append({
            'pos':   rs.get('position', 0),
            'name':  rs.get('name', ''),
            'rules': rules_desc,
        })

    return {
        'sides':      sides,
        'is_chained': recon.get('is_chained', False),
        'rule_sets':  rule_sets,
    }


def parse_adv_reconciliation(adv, res_map, col_map, seg_map, meta_map):
    """Parsea una conciliación avanzada con grupos conciliables + segmentos internos."""
    if not adv:
        return None

    groups = []
    for rg in (adv.get('reconcilable_groups') or []):
        prefix   = rg.get('prefix_side', '?')
        seg_id   = rg.get('segment_id')
        seg      = seg_map.get(seg_id) or {}
        seg_name = seg.get('name', f"ID:{seg_id}")
        resource_name = seg.get('resource', res_map.get(rg.get('resource_id'), '—'))
        seg_rules = fmt_filter_rules(seg.get('rules', []), col_map)

        sc2      = rg.get('segmentation_config') or {}
        crit_id  = sc2.get('criteria_column_id')
        crit_col = col_map.get(crit_id, f"ID:{crit_id}") if crit_id else "—"
        segments = [m.get('value', '') for m in sc2.get('segmentation_metadata', [])
                    if m.get('value')]

        groups.append({
            'prefix':        prefix,
            'resource_name': resource_name,
            'group_name':    seg_name,
            'group_filters': seg_rules,
            'crit_col':      crit_col,
            'segments':      segments,
        })

    rule_sets = []
    for rs in sorted(adv.get('reconciliation_rule_sets', []),
                     key=lambda x: x.get('position', 99)):
        rules_desc = []
        for rule in (rs.get('reconciliation_rules') or []):
            col_a = col_map.get(rule.get('column_a_id'), f"ID:{rule.get('column_a_id')}")
            col_b = col_map.get(rule.get('column_b_id'), f"ID:{rule.get('column_b_id')}")
            op    = rule.get('operator', '=')
            tol   = rule.get('tolerance', 0)
            tol_u = rule.get('tolerance_unit') or ''
            tol_s = f"  [tolerancia ±{tol} {tol_u}]" if tol else ""
            rules_desc.append(f"A.{col_a}  {op}  B.{col_b}{tol_s}")

        sweep = []
        for sw in (rs.get('sweep_sides') or []):
            p       = sw.get('prefix_side', '?')
            isr     = sw.get('input_sweep_resource') or {}
            meta_id = isr.get('segmentation_metadata_id')
            if meta_id:
                seg_val = meta_map.get(meta_id, f"ID:{meta_id}")
            else:
                seg_val = "(recurso completo sin segmentar)"
            sweep.append(f"Lado {p}: {seg_val}")

        rule_sets.append({
            'pos':        rs.get('position', 0),
            'name':       rs.get('name', ''),
            'cross_type': rs.get('cross_type', ''),
            'new_ver':    rs.get('is_new_version', False),
            'rules':      rules_desc,
            'sweep':      sweep,
        })

    return {'groups': groups, 'rule_sets': rule_sets}


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
            result.append({
                'seg_id': seg.get('export_id'),
                'name':   seg.get('name', ''),
                'rules':  rules,
            })
    return result


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
    clean = re.sub(r'[\\/*?:\[\]]', '', str(nombre))
    return (clean[:18] + "_" + str(eid))[:31]


def sort_key(r):
    return (RT_ORDER.get(r.get('resource_type', ''), 99), r.get('export_id', 0))


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
            ext_a = "" if sid in all_ids else " ↗"
            ext_b = "" if t_id in all_ids else " ↗"
            if t_id in rels:
                rels[t_id]["parents"].append(res_map.get(sid, str(sid)) + ext_a)
            if sid in rels:
                rels[sid]["children"].append(res_map.get(t_id, str(t_id)) + ext_b)
    return rels


# ══════════════════════════════════════════════════════════════════════════════
# GENERADOR EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def generar_excel(data, selected_ids):
    all_resources             = data.get('resources', [])
    nodes                     = data.get('nodes', [])
    res_map, col_map, seg_map, meta_map, seg_usage = build_maps(data)

    # Deduplicar + filtrar + ordenar (tipo → ID)
    seen, resources = set(), []
    for r in all_resources:
        eid = r.get('export_id')
        if eid in selected_ids and eid not in seen:
            seen.add(eid)
            resources.append(r)
    resources.sort(key=sort_key)

    rels      = build_relations(resources, nodes, res_map)
    map_hojas = {r.get('export_id'): limpiar_hoja(r.get('name', ''), r.get('export_id'))
                 for r in resources}

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        wb = writer.book

        # ── ÍNDICE ─────────────────────────────────────────────────────────────
        ws = wb.create_sheet("📚 Índice", 0)
        ws.sheet_view.showGridLines = False

        ws.merge_cells('A1:H1')
        c = ws.cell(1, 1, "SIMETRIK DOCUMENTATION  ·  PeYa Finance Operations & Payments")
        sc(c, bg=C["red"], bold=True, color=C["white"], size=13,
           ha='center', va='center', wrap=False)
        ws.row_dimensions[1].height = 32

        ws.merge_cells('A2:H2')
        c = ws.cell(2, 1,
            f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}   |   "
            f"Recursos documentados: {len(resources)}")
        sc(c, bg=C["dark"], color=C["white"], size=9, ha='center', va='center', wrap=False)
        ws.row_dimensions[2].height = 15

        idx_hdrs = ["#", "ID", "NOMBRE DEL RECURSO", "TIPO",
                    "PROVIENE DE", "ALIMENTA A", "LINK 🔗"]
        for i, h in enumerate(idx_hdrs, 1):
            hdr(ws.cell(4, i, h), h, bg=C["dark"])
        ws.row_dimensions[4].height = 20

        for row_n, res in enumerate(resources, 5):
            eid     = res.get('export_id')
            rt      = res.get('resource_type', '')
            bg      = C["grey"] if row_n % 2 == 0 else C["white"]

            vals = [row_n - 4, eid, res.get('name', ''), RT_LABEL.get(rt, rt),
                    ", ".join(rels[eid]["parents"]) or "— origen",
                    ", ".join(rels[eid]["children"]) or "— fin de flujo"]
            for col_n, val in enumerate(vals, 1):
                c = ws.cell(row_n, col_n, val)
                sc(c, bg=bg, size=9, va='center', wrap=False)
                if col_n == 4:
                    c.font = Font(name='Calibri', bold=True, size=9,
                                  color=RT_COLOR.get(rt, C["dark"]))
                c.border = mk_border()

            lnk = ws.cell(row_n, 7, "Ver →")
            lnk.hyperlink = f"#'{map_hojas[eid]}'!A1"
            lnk.font = Font(name='Calibri', color="0D47A1", underline="single", size=9)
            lnk.border = mk_border()
            ws.row_dimensions[row_n].height = 15

        for col_n, w in enumerate([6, 11, 46, 26, 38, 38, 8], 1):
            ws.column_dimensions[chr(64+col_n)].width = w

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
            c = ws.cell(row, 1, RT_LABEL.get(rt, '') + "  ·  " + name)
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
            std = parse_std_reconciliation(
                res.get('reconciliation'), res_map, col_map, seg_map)
            if std:
                row = section_title(ws, row, "⚖️  REGLAS DE CONCILIACIÓN ESTÁNDAR",
                                    bg=C["red"], cols=COLS)

                # Grupos conciliables activos (A y B)
                row = section_title(ws, row, "  GRUPOS CONCILIABLES ACTIVOS",
                                    bg=C["rose"], cols=COLS)
                for col_n, h in enumerate(
                        ["LADO", "RECURSO", "GRUPO CONCILIABLE (ACTIVO)",
                         "FILTROS DEL GRUPO"], 1):
                    hdr(ws.cell(row, col_n, h), h, bg=C["rose"])
                ws.merge_cells(f'D{row}:E{row}')
                ws.row_dimensions[row].height = 18
                row += 1

                for i, side in enumerate(std['sides']):
                    bg = C["grey"] if i % 2 == 0 else "FFFFFF"
                    trig = "  [TRIGGER]" if side['is_trigger'] else ""
                    c1 = ws.cell(row, 1, side['prefix'] + trig)
                    c2 = ws.cell(row, 2, side['resource_name'])
                    c3 = ws.cell(row, 3, side['group_name'])
                    ws.merge_cells(f'D{row}:E{row}')
                    c4 = ws.cell(row, 4, side['group_filters'])
                    n_lines = side['group_filters'].count('\n') + 1
                    for c, al in [(c1,'center'),(c2,'left'),(c3,'left'),(c4,'left')]:
                        sc(c, bg=bg, size=9, va='top', wrap=True, ha=al)
                    ws.row_dimensions[row].height = row_height(n_lines)
                    row += 1
                row += 1

                row = meta_row(ws, row, "Conciliación encadenada",
                               "Sí" if std['is_chained'] else "No", cols=COLS)
                row += 1

                # Rule sets
                row = section_title(ws, row, "  RULE SETS DE MATCHING",
                                    bg="C62828", cols=COLS)
                for col_n, h in enumerate(
                        ["POS.", "NOMBRE DEL RULE SET", "REGLAS  (A vs B)"], 1):
                    hdr(ws.cell(row, col_n, h), h, bg="C62828")
                ws.merge_cells(f'C{row}:E{row}')
                ws.row_dimensions[row].height = 18
                row += 1

                for i, rs in enumerate(std['rule_sets']):
                    bg = C["grey"] if i % 2 == 0 else "FFFFFF"
                    c1 = ws.cell(row, 1, rs['pos'])
                    c2 = ws.cell(row, 2, rs['name'])
                    ws.merge_cells(f'C{row}:E{row}')
                    c3 = ws.cell(row, 3, "\n".join(rs['rules']))
                    for c, al in [(c1,'center'),(c2,'left'),(c3,'left')]:
                        sc(c, bg=bg, size=9, va='top', wrap=True, ha=al)
                    ws.row_dimensions[row].height = row_height(len(rs['rules']))
                    row += 1
                row += 1

            # ── CONCILIACIÓN AVANZADA ──────────────────────────────────────────
            adv_parsed = parse_adv_reconciliation(
                res.get('advanced_reconciliation'), res_map, col_map, seg_map, meta_map)
            if adv_parsed:
                row = section_title(ws, row, "🔬  REGLAS DE CONCILIACIÓN AVANZADA",
                                    bg=C["purple"], cols=COLS)

                # Grupos conciliables + segmentos internos
                row = section_title(ws, row, "  GRUPOS CONCILIABLES Y SEGMENTOS INTERNOS",
                                    bg="6A1B9A", cols=COLS)
                for col_n, h in enumerate(
                        ["LADO", "RECURSO", "GRUPO CONCILIABLE",
                         "FILTROS DEL GRUPO", "SEGMENTOS INTERNOS"], 1):
                    hdr(ws.cell(row, col_n, h), h, bg="6A1B9A")
                ws.row_dimensions[row].height = 18
                row += 1

                for i, g in enumerate(adv_parsed['groups']):
                    bg = C["grey"] if i % 2 == 0 else "FFFFFF"
                    segs_txt = "\n".join(g['segments']) if g['segments'] else "(sin segmentación interna)"
                    n_lines = max(g['group_filters'].count('\n') + 1,
                                  len(g['segments']) if g['segments'] else 1)
                    c1 = ws.cell(row, 1, g['prefix'])
                    c2 = ws.cell(row, 2, g['resource_name'])
                    c3 = ws.cell(row, 3, g['group_name'])
                    c4 = ws.cell(row, 4, g['group_filters'])
                    c5 = ws.cell(row, 5, segs_txt)
                    for c, al in [(c1,'center'),(c2,'left'),(c3,'left'),(c4,'left'),(c5,'left')]:
                        sc(c, bg=bg, size=9, va='top', wrap=True, ha=al)
                    ws.row_dimensions[row].height = row_height(n_lines)
                    row += 1
                row += 1

                # Rule sets avanzados con segmentos resueltos
                row = section_title(ws, row, "  RULE SETS (SEGMENTO A vs SEGMENTO B)",
                                    bg="4A148C", cols=COLS)
                for col_n, h in enumerate(
                        ["POS.", "NOMBRE / TIPO", "REGLAS  (A vs B)",
                         "SEGMENTO LADO A", "SEGMENTO LADO B"], 1):
                    hdr(ws.cell(row, col_n, h), h, bg="4A148C")
                ws.row_dimensions[row].height = 18
                row += 1

                for i, rs in enumerate(adv_parsed['rule_sets']):
                    bg = C["grey"] if i % 2 == 0 else "FFFFFF"
                    name_txt = rs['name']
                    if rs['cross_type']:
                        name_txt += "\n[" + rs['cross_type'] + "]"
                    if rs['new_ver']:
                        name_txt += "  ✦ new version"

                    # Separar sweep por lado A y B
                    seg_a = next((s.replace("Lado A: ", "")
                                  for s in rs['sweep'] if s.startswith("Lado A")), "—")
                    seg_b = next((s.replace("Lado B: ", "")
                                  for s in rs['sweep'] if s.startswith("Lado B")), "—")

                    c1 = ws.cell(row, 1, rs['pos'])
                    c2 = ws.cell(row, 2, name_txt)
                    c3 = ws.cell(row, 3, "\n".join(rs['rules']))
                    c4 = ws.cell(row, 4, seg_a)
                    c5 = ws.cell(row, 5, seg_b)
                    for c, al in [(c1,'center'),(c2,'left'),(c3,'left'),(c4,'left'),(c5,'left')]:
                        sc(c, bg=bg, size=9, va='top', wrap=True, ha=al)
                    n_lines = max(len(rs['rules']), 1)
                    ws.row_dimensions[row].height = row_height(n_lines)
                    row += 1
                row += 1

            # ── SOURCE GROUP ──────────────────────────────────────────────────
            sg = res.get('source_group')
            if sg:
                row = section_title(ws, row,
                    "📊  CONFIGURACIÓN DE AGRUPACIÓN (GROUP BY)",
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
                for col_n, h in enumerate(
                        ["FUENTE", "GRUPO CONCILIABLE", "ROL", "FILTROS DEL GRUPO"], 1):
                    hdr(ws.cell(row, col_n, h), h, bg=C["teal"])
                ws.merge_cells(f'D{row}:E{row}')
                ws.row_dimensions[row].height = 18
                row += 1

                for i, us in enumerate(su.get('union_segments') or []):
                    seg_id   = us.get('segment_id')
                    seg_info = seg_map.get(seg_id) or {}
                    resource_name = seg_info.get('resource', seg_info.get('resource_name', f"ID:{seg_id}"))
                    group_name    = seg_info.get('name', f"ID:{seg_id}")
                    filters_text  = fmt_filter_rules(seg_info.get('rules', []), col_map)
                    rol = "TRIGGER · " + (us.get('trigger_type') or '') if us.get('is_trigger') else "Fuente adicional"
                    n_lines = max(filters_text.count('\n') + 1, 1)
                    bg = C["grey"] if i % 2 == 0 else "FFFFFF"
                    c1 = ws.cell(row, 1, resource_name)
                    c2 = ws.cell(row, 2, group_name)
                    c3 = ws.cell(row, 3, rol)
                    ws.merge_cells(f'D{row}:E{row}')
                    c4 = ws.cell(row, 4, filters_text)
                    for c, al in [(c1,'left'),(c2,'left'),(c3,'center'),(c4,'left')]:
                        sc(c, bg=bg, size=9, va='top', wrap=True, ha=al)
                    ws.row_dimensions[row].height = row_height(n_lines)
                    row += 1
                row += 1

            # ── GRUPOS CONCILIABLES DEL RECURSO ──────────────────────────────
            segs_all = parse_segment_filters(res.get('segments', []), col_map)
            if segs_all:
                row = section_title(ws, row,
                    "🔍  GRUPOS CONCILIABLES DEL RECURSO",
                    bg=C["slate"], cols=COLS)
                for col_n, h in enumerate(["NOMBRE DEL GRUPO", "FILTROS APLICADOS"], 1):
                    hdr(ws.cell(row, col_n, h), h, bg=C["slate"])
                ws.merge_cells(f'B{row}:D{row}')
                hdr(ws.cell(row, 5, "USADO EN"), "USADO EN", bg=C["slate"])
                ws.row_dimensions[row].height = 18
                row += 1
                for i, seg in enumerate(segs_all):
                    bg = C["grey"] if i % 2 == 0 else "FFFFFF"
                    # Resolve usage
                    usages = seg_usage.get(seg['seg_id'], [])
                    if usages:
                        usage_lines = [u[0] + " (" + u[1] + ")" for u in usages]
                        usage_text = "\n".join(usage_lines)
                    else:
                        usage_text = "Sin uso en flujo activo"
                    n_lines = max(len(seg['rules']), len(usages) if usages else 1)
                    c1 = ws.cell(row, 1, seg['name'])
                    ws.merge_cells(f'B{row}:D{row}')
                    c2 = ws.cell(row, 2, "\n".join(seg['rules']))
                    c3 = ws.cell(row, 5, usage_text)
                    sc(c1, bg=bg, size=9, va='top', wrap=True)
                    sc(c2, bg=bg, size=9, va='top', wrap=True)
                    sc(c3, bg=bg, size=9, va='top', wrap=True,
                       color="1B5E20" if usages else "888888")
                    ws.row_dimensions[row].height = row_height(n_lines)
                    row += 1
                row += 1

            # ── COLUMNAS ──────────────────────────────────────────────────────
            columns = sorted(res.get('columns') or [],
                             key=lambda x: x.get('position', 0))
            if columns:
                row = section_title(ws, row,
                    "📋  CONFIGURACIÓN DE COLUMNAS",
                    bg=C["blue"], cols=COLS)
                for col_n, h in enumerate(
                        ["LABEL / NOMBRE", "TIPO DATO", "TIPO COL.",
                         "LÓGICA · FÓRMULA · BUSCAR V"], 1):
                    hdr(ws.cell(row, col_n, h), h, bg=C["blue"])
                ws.merge_cells(f'D{row}:E{row}')
                ws.row_dimensions[row].height = 18
                row += 1
                for i, col in enumerate(columns):
                    label    = col.get('label') or col.get('name', '')
                    dtype    = col.get('data_format', '')
                    col_type = (col.get('column_type') or '').replace('_', ' ').upper()
                    logic    = parse_transformation_logic(col, res_map, col_map)
                    bg       = C["grey"] if i % 2 == 0 else "FFFFFF"
                    c1 = ws.cell(row, 1, label)
                    c2 = ws.cell(row, 2, dtype)
                    c3 = ws.cell(row, 3, col_type)
                    ws.merge_cells(f'D{row}:E{row}')
                    c4 = ws.cell(row, 4, logic)
                    for c, al in [(c1,'left'),(c2,'center'),(c3,'center'),(c4,'left')]:
                        sc(c, bg=bg, size=9, va='top', wrap=True, ha=al)
                    ws.row_dimensions[row].height = row_height(logic.count('\n') + 1)
                    row += 1

            ws.column_dimensions['A'].width = 26
            ws.column_dimensions['B'].width = 22
            ws.column_dimensions['C'].width = 22
            ws.column_dimensions['D'].width = 22
            ws.column_dimensions['E'].width = 32

        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

    output.seek(0)
    return output


# ══════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ══════════════════════════════════════════════════════════════════════════════

# Custom CSS global — Roboto desde Google Fonts
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&family=Roboto+Mono:wght@400;500&display=swap');

/* Fuente global Roboto */
html, body, [class*="css"], .stMarkdown, .stCaption,
.stMetric, .stButton, .stDownloadButton,
div[data-testid], p, span, label, h1, h2, h3, h4 {
    font-family: 'Roboto', sans-serif !important;
}

/* Monospace para IDs */
code, .font-mono { font-family: 'Roboto Mono', monospace !important; }

/* Quitar padding top excesivo */
.block-container { padding-top: 1.5rem !important; }

/* Progress bar generación */
.stProgress > div > div { background: #EA0050 !important; }

/* Botón primary */
div[data-testid="stButton"] button[kind="primary"] {
    background: #EA0050 !important;
    border: none !important;
    font-size: 1rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.3px;
    font-family: 'Roboto', sans-serif !important;
}
div[data-testid="stButton"] button[kind="primary"]:hover {
    background: #C0003A !important;
}
div[data-testid="stDownloadButton"] button {
    font-family: 'Roboto', sans-serif !important;
}
</style>
""", unsafe_allow_html=True)

# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='background:linear-gradient(135deg,#EA0050 0%,#B0003A 100%);
    padding:24px 32px 20px;border-radius:14px;
    box-shadow:0 4px 20px rgba(234,0,80,0.2);margin-bottom:20px'>
    <div style='display:flex;align-items:center;gap:14px'>
        <span style='font-size:2.2rem'>📊</span>
        <div>
            <h1 style='color:white;margin:0;font-family:Arial,sans-serif;
                font-size:1.75rem;font-weight:700;letter-spacing:-0.3px'>
                Simetrik Documentation 
            </h1>
            <p style='color:rgba(255,255,255,0.82);margin:4px 0 0;font-size:0.9rem;font-family:Arial'>
                PeYa Finance Operations &amp; Payments &nbsp;·&nbsp; v2.2 · Jef
            </p>
        </div>
    </div>
</div>""", unsafe_allow_html=True)

# ── UPLOAD ────────────────────────────────────────────────────────────────────
up = st.file_uploader(
    "**Carga el JSON exportado desde Simetrik**",
    type=['json'],
    help="En Simetrik: Flujo → ⚙️ Configuración → Exportar JSON",
    label_visibility="visible"
)

if not up:
    st.markdown("""
    <div style='background:#F8F9FA;border:2px dashed #dee2e6;border-radius:12px;
        padding:32px;text-align:center;margin-top:8px'>
        <div style='font-size:2rem;margin-bottom:8px'>📂</div>
        <p style='color:#666;font-size:0.95rem;margin:0'>
            Arrastra el JSON aquí o usá el botón de arriba para seleccionarlo
        </p>
        <p style='color:#aaa;font-size:0.8rem;margin:6px 0 0'>
            En Simetrik: Flujo → ⚙️ Configuración → Exportar JSON
        </p>
    </div>""", unsafe_allow_html=True)
    st.stop()

try:
    data = json.load(up)
except Exception as e:
    st.error(f"Error al leer el JSON: {e}")
    st.stop()

all_resources = data.get('resources', [])
nodes         = data.get('nodes', [])

seen_load, resources_unique = set(), []
for r in all_resources:
    eid = r.get('export_id')
    if eid not in seen_load:
        seen_load.add(eid)
        resources_unique.append(r)

res_map, col_map, seg_map, meta_map, seg_usage = build_maps(data)
resources_unique.sort(key=sort_key)
rels_all = build_relations(resources_unique, nodes, res_map)

# Resumen del flujo cargado
total_recons = sum(1 for r in resources_unique
                   if r.get('resource_type') in ('reconciliation','advanced_reconciliation'))
total_fuentes = sum(1 for r in resources_unique if r.get('resource_type') == 'native')

col_s1, col_s2, col_s3, col_s4 = st.columns(4)
col_s1.metric("📦 Recursos totales", len(resources_unique))
col_s2.metric("📥 Fuentes", total_fuentes)
col_s3.metric("⚖️ Conciliaciones", total_recons)
# Nombre del JSON sin truncar con metric — usar markdown para control total
with col_s4:
    nombre_display = up.name if len(up.name) <= 28 else up.name[:25] + "…"
    st.markdown(
        "<div style='font-size:0.75rem;color:#888;margin-bottom:4px'>📄 JSON cargado</div>"
        "<div style='font-size:0.95rem;font-weight:600;color:#1a1a1a;word-break:break-all'>"
        + nombre_display + "</div>",
        unsafe_allow_html=True
    )

st.markdown("<hr style='margin:16px 0;border-color:#f0f0f0'>", unsafe_allow_html=True)

# ── PASO 1: SELECCIÓN ─────────────────────────────────────────────────────────
st.markdown("### 1️⃣ &nbsp; Selecciona los recursos a documentar")

# Filtro por tipo — pills horizontales
all_types = sorted({r.get('resource_type', '') for r in resources_unique},
                   key=lambda x: RT_ORDER.get(x, 99))

col_f1, col_f2 = st.columns([4, 1])
with col_f1:
    # Mostrar los tipos con su label completo en el multiselect
    filtro_tipo = st.multiselect(
        "Filtrar por tipo de recurso",
        options=all_types,
        format_func=lambda x: RT_LABEL.get(x, x),
        default=all_types,
        label_visibility="collapsed",
        placeholder="Selecciona tipos de recurso a mostrar…"
    )

resources_visible = [r for r in resources_unique
                     if r.get('resource_type', '') in filtro_tipo]

# Botones Todos / Ninguno + contador
bc1, bc2, bc3 = st.columns([1, 1, 6])
select_all   = bc1.button("✅ Todos", use_container_width=True)
deselect_all = bc2.button("☐ Ninguno", use_container_width=True)

if 'sel' not in st.session_state:
    st.session_state.sel = {r.get('export_id'): True for r in resources_unique}
if select_all:
    for r in resources_visible:
        st.session_state.sel[r.get('export_id')] = True
if deselect_all:
    for r in resources_visible:
        st.session_state.sel[r.get('export_id')] = False

st.write("")

tipo_groups: dict = {}
for r in resources_visible:
    tipo_groups.setdefault(r.get('resource_type', ''), []).append(r)

selected_ids = set()
for rt in sorted(tipo_groups.keys(), key=lambda x: RT_ORDER.get(x, 99)):
    group = tipo_groups[rt]
    color  = RT_COLOR.get(rt, C["dark"])
    label  = RT_LABEL.get(rt, rt)

    # Encabezado de grupo con badge
    st.markdown(
        f"<div style='display:flex;align-items:center;gap:10px;margin:14px 0 6px'>"
        f"<span style='background:#{color};color:white;padding:4px 14px;"
        f"border-radius:20px;font-size:0.8rem;font-weight:700;white-space:nowrap'>"
        f"{label}</span>"
        f"<span style='color:#aaa;font-size:0.8rem'>{len(group)} recursos</span>"
        f"</div>",
        unsafe_allow_html=True
    )

    for r in group:
        eid   = r.get('export_id')
        name  = r.get('name', '')
        pars  = ", ".join(rels_all[eid]["parents"]) or "—"
        chils = ", ".join(rels_all[eid]["children"]) or "—"

        ca, cb = st.columns([0.3, 9.7])
        checked = ca.checkbox("", value=st.session_state.sel.get(eid, True), key=f"chk_{eid}")
        st.session_state.sel[eid] = checked

        opacity = "1" if checked else "0.45"
        cb.markdown(
            f"<div style='opacity:{opacity};padding:5px 0'>"
            f"<span style='font-weight:600;font-size:0.9rem;color:#1a1a1a'>{name}</span>"
            f"&nbsp;&nbsp;<span style='font-size:0.75rem;color:#999;font-family:monospace'>{eid}</span><br>"
            f"<span style='font-size:0.75rem;color:#777'>⬅️ {pars[:80]}{'…' if len(pars)>80 else ''}"
            f"&nbsp;&nbsp;➡️ {chils[:80]}{'…' if len(chils)>80 else ''}</span>"
            f"</div>",
            unsafe_allow_html=True
        )
        if checked:
            selected_ids.add(eid)

# ── PASO 2: GENERAR ───────────────────────────────────────────────────────────
st.markdown("<hr style='margin:20px 0;border-color:#f0f0f0'>", unsafe_allow_html=True)

n_sel = len(selected_ids)

# Panel de resumen de selección — construido con concatenación para evitar
# conflictos de comillas en f-strings anidados
if n_sel > 0:
    tipos_sel = {}
    for r in resources_unique:
        if r.get('export_id') in selected_ids:
            rt = r.get('resource_type', '')
            tipos_sel[rt] = tipos_sel.get(rt, 0) + 1

    badges_html = ""
    for rt, cnt in sorted(tipos_sel.items(), key=lambda x: RT_ORDER.get(x[0], 99)):
        color = RT_COLOR.get(rt, "444444")
        label = RT_LABEL.get(rt, rt)
        badges_html += (
            "<span style='background:#" + color + ";color:white;"
            "padding:3px 12px;border-radius:12px;font-size:0.78rem;"
            "font-weight:700;white-space:nowrap'>"
            + label + " (" + str(cnt) + ")"
            "</span> "
        )

    sel_label = "seleccionados" if n_sel != 1 else "seleccionado"
    resumen_html = (
        "<div style='background:#F8F9FA;border-radius:10px;padding:12px 16px;"
        "margin-bottom:12px;display:flex;align-items:center;gap:10px;flex-wrap:wrap'>"
        "<span style='font-weight:700;color:#333;white-space:nowrap'>📋 "
        + str(n_sel) + " " + sel_label + ":</span>"
        + badges_html
        + "</div>"
    )
    st.markdown(resumen_html, unsafe_allow_html=True)
else:
    st.warning("Selecciona al menos un recurso para continuar.")
    st.stop()

nombre_dl = f"{os.path.splitext(up.name)[0]}_DOC_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx"

if st.button("🚀  GENERAR EXCEL", type="primary", use_container_width=True):
    prog = st.progress(0, text="Iniciando...")
    try:
        prog.progress(15, text="Procesando recursos...")
        prog.progress(40, text="Resolviendo grupos conciliables...")
        prog.progress(65, text="Construyendo reglas de conciliacion...")
        excel_bytes = generar_excel(data, selected_ids)
        prog.progress(90, text="Aplicando estilos...")
        prog.progress(100, text="Listo.")
        st.success(f"✅ Excel generado con **" + str(n_sel) + "** recursos documentados.")
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
        prog.empty()
        st.error(f"Error al generar el Excel: {e}")
        import traceback
        st.code(traceback.format_exc())

st.markdown("<hr style='margin:24px 0;border-color:#f0f0f0'>", unsafe_allow_html=True)
st.caption("Simetrik Documentation · PeYa Finance Operations & Payments · v2.2 · Jef")
