import os
import re
import json
import pandas as pd
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

class PBIDocumenter:
    def __init__(self, root_path):
        self.root_path = root_path
        self.model_path = self._find_path('.SemanticModel')
        self.report_path = self._find_path('.Report')
        self.tables_path = os.path.join(self.model_path, 'definition', 'tables') if self.model_path else None

    def _find_path(self, suffix):
        for root, dirs, _ in os.walk(self.root_path):
            if root.endswith(suffix): return root
        return None

    def parse_tmdl_files(self):
        tables, columns, measures = [], [], []
        if not self.tables_path: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        for file in os.listdir(self.tables_path):
            if not file.endswith('.tmdl'): continue
            with open(os.path.join(self.tables_path, file), 'r', encoding='utf-8') as f:
                content = f.read()
                table_name = file.replace('.tmdl', '')
                header = content.split("column")[0]

                tables.append({
                    "Nombre": table_name,
                    "Modo": re.search(r"mode:\s+(\w+)", header).group(1) if "mode:" in header else "Import",
                    "Tipo": "DAX (Calculada)" if "partition " + table_name + " = calculated" in content else "Power Query",
                    "Visible": "No" if "isHidden" in header else "Sí",
                    "Descripción": re.search(r'description:\s+"(.*?)"', header).group(1) if 'description:' in header else ""
                })

                col_blocks = re.finditer(r"column\s+([\w\s'-]+)(.*?)(?=\n\s+(column|measure|partition|annotation)|$)", content, re.DOTALL)
                for m in col_blocks:
                    body = m.group(2)
                    columns.append({
                        "Tabla": table_name, "Campo": m.group(1).strip().replace("'", ""),
                        "Tipo Dato": re.search(r"dataType:\s+(\w+)", body).group(1) if "dataType" in body else "Inferred",
                        "Visible": "No" if "isHidden" in body else "Sí",
                        "Descripción": re.search(r'description:\s+"(.*?)"', body).group(1) if 'description:' in body else ""
                    })

                meas_blocks = re.finditer(r"measure\s+'?([^'=]+)'?\s+=\s+(.+?)(?=\n\s+(measure|column|partition|annotation)|$)", content, re.DOTALL)
                for m in meas_blocks:
                    rest = m.group(2)
                    measures.append({
                        "Tabla": table_name, "Medida": m.group(1).strip(), 
                        "DAX": rest.split('\n')[0].strip(),
                        "Formato": re.search(r"formatString:\s+(.+)", rest).group(1) if "formatString" in rest else "-",
                        "Visible": "No" if "isHidden" in rest else "Sí",
                        "Descripción": re.search(r'description:\s+"(.*?)"', rest).group(1) if 'description:' in rest else ""
                    })

        return (
            pd.DataFrame(tables),
            pd.DataFrame(columns) if columns else pd.DataFrame(columns=["Tabla", "Campo", "Tipo Dato", "Visible", "Descripción"]),
            pd.DataFrame(measures) if measures else pd.DataFrame(columns=["Tabla", "Medida", "DAX", "Formato", "Visible", "Descripción"])
        )
    
    def parse_relationships(self):
        rels = []

        # -----------------------------
        # 1) relationships.tmdl (tu caso real)
        # -----------------------------
        rel_tmdl = os.path.join(self.model_path, "definition", "relationships.tmdl")
        if os.path.exists(rel_tmdl):
            with open(rel_tmdl, "r", encoding="utf-8") as f:
                content = f.read()

            blocks = re.split(r"\brelationship\b", content)
            for block in blocks[1:]:
                from_col = re.search(r"fromColumn:\s*([^\n\r]+)", block)
                to_col = re.search(r"toColumn:\s*([^\n\r]+)", block)
                from_card = re.search(r"fromCardinality:\s*([^\n\r]+)", block)

                if from_col and to_col:
                    from_full = from_col.group(1).strip()
                    to_full = to_col.group(1).strip()

                    # Separar tabla.columna
                    if "." in from_full:
                        from_table, from_column = from_full.split(".", 1)
                    else:
                        from_table, from_column = "", from_full

                    if "." in to_full:
                        to_table, to_column = to_full.split(".", 1)
                    else:
                        to_table, to_column = "", to_full

                    # Interpretar cardinalidad (con supuestos)
                    card = ""
                    if from_card:
                        raw = from_card.group(1).strip().lower()
                        if raw == "one":
                            card = "One-to-One"
                        elif raw == "many":
                            card = "One-to-Many"
                    else:
                        # Sin info explícita: asumimos Many-to-One (from = many, to = one)
                        card = "Many-to-One"

                    rels.append({
                        "Origen": from_table,
                        "Campo O": from_column,
                        "Destino": to_table,
                        "Campo D": to_column,
                        "Cardinalidad": card
                    })

        # -----------------------------
        # 2) DataFrame final
        # -----------------------------
        if not rels:
            print("⚠ No se detectaron relaciones en relationships.tmdl")
            return pd.DataFrame(columns=["Origen", "Campo O", "Destino", "Campo D", "Cardinalidad"])

        return pd.DataFrame(rels)

    def get_report_visuals(self):
        visuals = []
        mapping = {"lineChart": "Gráfico de líneas", "card": "Tarjeta", "barChart": "Gráfico de barras", "pieChart": "Gráfico circular", "table": "Tabla", "pivotTable": "Matriz"}
        ignore = ["shape", "image", "textbox", "button"]
        
        pages_path = os.path.join(self.report_path, 'definition', 'pages')
        if not os.path.exists(pages_path): return pd.DataFrame(columns=["Hoja", "Tipo de Objeto", "ID Técnico"])

        for page_folder in os.listdir(pages_path):
            page_full = os.path.join(pages_path, page_folder)
            if not os.path.isdir(page_full): continue
            
            p_name = page_folder
            if os.path.exists(os.path.join(page_full, 'page.json')):
                with open(os.path.join(page_full, 'page.json'), 'r', encoding='utf-8') as f:
                    p_name = json.load(f).get('displayName', page_folder)

            v_dir = os.path.join(page_full, 'visuals')
            if os.path.exists(v_dir):
                for v_folder in os.listdir(v_dir):
                    v_json = os.path.join(v_dir, v_folder, 'visual.json')
                    if os.path.exists(v_json):
                        with open(v_json, 'r', encoding='utf-8') as f:
                            v_type = json.load(f).get('visual', {}).get('visualType', 'Otro')
                            if v_type not in ignore:
                                visuals.append({"Hoja": p_name, "Tipo de Objeto": mapping.get(v_type, v_type), "ID Técnico": v_folder})
        df = pd.DataFrame(visuals)
        if df.empty:
            df = pd.DataFrame(columns=["Hoja", "Tipo de Objeto", "ID Técnico"])
        return df

    def write_df(self, writer, df, sheet_name):
        if df.empty:
            # Crear una fila vacía para que Excel muestre encabezados
            empty_row = {col: "" for col in df.columns}
            df = pd.DataFrame([empty_row])
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    def generate_excel(self):
        df_t, df_c, df_m = self.parse_tmdl_files()
        df_r = self.parse_relationships()
        df_p = self.get_report_visuals()
        name = os.path.basename(self.root_path.strip('/\\'))
        
        with pd.ExcelWriter("Documentacion_PBI.xlsx", engine='openpyxl') as writer:
            # 1. Documentación
            doc = [["Documentación"], [""], ["Reporte:", name], [""], ["Objetivo:", ""], [""], ["Workspace:", ""], [""], ["Actualizaciones:", ""]]
            pd.DataFrame(doc).to_excel(writer, sheet_name='Documentación', index=False, header=False)
            
            # 2. Otras hojas
            self.write_df(writer, df_t, 'Tablas')
            self.write_df(writer, df_r, 'Relaciones')
            self.write_df(writer, df_c, 'Campos')
            self.write_df(writer, df_m, 'Medidas')
            self.write_df(writer, df_p, 'Hojas')

            # --- Formato Final ---
            for sheet in writer.sheets:
                ws = writer.sheets[sheet]
                for col in ws.columns:
                    max_length = 0
                    col_name = col[0].column_letter
                    for cell in col:
                        if cell.value: max_length = max(max_length, len(str(cell.value)))
                    ws.column_dimensions[col_name].width = min(max_length + 2, 50)
                # Negrita encabezados
                for cell in ws[1]: cell.font = Font(bold=True)
            
            writer.sheets['Documentación']['A1'].font = Font(size=18, bold=True)

        print("✅ Documentación generada con éxito.")

if __name__ == "__main__":
    import sys
    PBIDocumenter(sys.argv[1] if len(sys.argv) > 1 else ".").generate_excel()
