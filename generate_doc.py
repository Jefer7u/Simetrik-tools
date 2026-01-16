import json
import pandas as pd
from datetime import datetime

# ========= CONFIG =========
JSON_PATH = "FLUJO_CR_DLOCAL.json"
OUTPUT_EXCEL = f"Resumen_Flujo_Simetrik_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# ========= LOAD JSON =========
with open(JSON_PATH, "r", encoding="utf-8") as f:
    data = json.load(f)

resources = data.get("resources", [])

writer = pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl")

# ========= 00 - SUMMARY =========
summary_data = {
    "Flow ID": data.get("_id"),
    "Version": data.get("version"),
    "Hash": data.get("hash"),
    "Cantidad de recursos": len(resources),
    "Fecha generación": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
}

summary_df = pd.DataFrame(list(summary_data.items()), columns=["Campo", "Valor"])
summary_df.to_excel(writer, sheet_name="00_Summary", index=False)

# ========= 01 - EXPLICACION =========
explicacion = [
    "Este documento describe la configuración técnica y funcional del flujo de conciliación configurado en Simetrik.",
    f"El flujo contiene {len(resources)} recursos, incluyendo conciliaciones y fuentes de datos.",
    "Cada recurso detalla sus columnas, transformaciones, segmentos y reglas de conciliación cuando aplica.",
]

exp_df = pd.DataFrame({"Descripción": explicacion})
exp_df.to_excel(writer, sheet_name="01_Explicacion_Funcional", index=False)

# ========= RECURSOS =========
for res in resources:
    res_name = res.get("name", "Recurso").replace(" ", "_")[:25]

    rows = []
    for col in res.get("columns", []):
        transformations = col.get("transformations") or []
        rows.append({
            "Resource name": res.get("name"),
            "Resource type": res.get("resource_type"),
            "Column name": col.get("name"),
            "Label": col.get("label"),
            "Column type": col.get("column_type"),
            "Data type": col.get("data_format"),
            "Hidden": col.get("is_hidden"),
            "Has transformation": bool(transformations),
            "Transformation type": transformations[0]["operation"] if transformations else None,
            "Formula / Query": transformations[-1]["query"] if transformations else None,
            "Position": col.get("position"),
        })

    df_cols = pd.DataFrame(rows)
    df_cols.to_excel(writer, sheet_name=f"Recurso_{res_name}", index=False)

    # ===== Rules =====
    if res.get("reconciliation"):
        rule_rows = []
        for rs in res["reconciliation"].get("reconciliation_rule_sets", []):
            for rule in rs.get("reconciliation_rules", []):
                rule_rows.append({
                    "Rule set": rs.get("name"),
                    "Priority": rs.get("position"),
                    "Operator": rule.get("operator"),
                    "Tolerance": rule.get("tolerance"),
                    "Tolerance unit": rule.get("tolerance_unit"),
                    "Column A ID": rule.get("column_a_id"),
                    "Column B ID": rule.get("column_b_id"),
                })

        pd.DataFrame(rule_rows).to_excel(
            writer, sheet_name=f"Rules_{res_name}", index=False
        )

# ========= SAVE =========
writer.close()
print(f"Archivo generado: {OUTPUT_EXCEL}")
