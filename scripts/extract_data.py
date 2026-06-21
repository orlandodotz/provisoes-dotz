import json, openpyxl, glob, os

# Encontra o arquivo .xlsx na raiz do repositório
xlsx_files = glob.glob("*.xlsx")
if not xlsx_files:
    raise FileNotFoundError("Nenhum arquivo .xlsx encontrado na raiz do repositório.")

wb = openpyxl.load_workbook(xlsx_files[0], data_only=True)
sheet = wb["BaseConsol"]

records = []
headers = [cell.value for cell in sheet[1]]

for row in sheet.iter_rows(min_row=2, values_only=True):
    empresa = row[0]
    if not empresa:
        continue
    valor = row[7]
    try:
        valor = float(valor) if valor is not None else 0.0
    except (TypeError, ValueError):
        valor = 0.0

    records.append({
        "empresa":        str(row[0] or ""),
        "mes":            str(row[1] or ""),
        "ntrans":         str(row[2] or ""),
        "ref3":           str(row[3] or ""),
        "conta":          str(row[4] or ""),
        "fornecedor":     str(row[5] or ""),
        "subgrupo":       str(row[6] or ""),
        "valor":          round(valor, 2),
        "descricaoConta": str(row[8] or ""),
        "subLimpo":       str(row[9] or ""),
    })

with open("data.json", "w", encoding="utf-8") as f:
    json.dump(records, f, ensure_ascii=False, indent=2)

print(f"data.json gerado com {len(records)} registros.")
