import json, openpyxl, glob
from datetime import datetime

MESES_PT = {
    1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun',
    7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'
}

def fmt_mes(val):
    """Converte qualquer formato de data/texto para 'dez/25'."""
    if val is None:
        return ""
    if isinstance(val, (datetime,)):
        return f"{MESES_PT[val.month]}/{str(val.year)[2:]}"
    s = str(val).strip()
    # já está no formato correto "dez/25" ou "dez-25"
    if len(s) == 6 and ('/' in s or '-' in s):
        return s.replace('-', '/')
    # tenta parsear como datetime string "2025-12-01 00:00:00"
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%Y-%m-%d', '%Y/%m/%d'):
        try:
            dt = datetime.strptime(s[:len(fmt.replace('%Y','0000').replace('%m','00').replace('%d','00').replace('%H','00').replace('%M','00').replace('%S','00'))], fmt)
            return f"{MESES_PT[dt.month]}/{str(dt.year)[2:]}"
        except Exception:
            pass
    # fallback: retorna como está
    return s

xlsx_files = glob.glob("*.xlsx")
if not xlsx_files:
    raise FileNotFoundError("Nenhum arquivo .xlsx encontrado na raiz do repositório.")

wb = openpyxl.load_workbook(xlsx_files[0], data_only=True)
sheet = wb["BaseConsol"]

records = []
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
        "mes":            fmt_mes(row[1]),
        "ntrans":         str(row[2] or ""),
        "ref3":           str(row[3] or ""),
        "conta":          str(row[4] or ""),
        "fornecedor":     str(row[5] or ""),
        "subgrupo":       str(row[6] or ""),
        "valor":          round(valor, 2),
        "descricaoConta": str(row[8] or ""),
        "subLimpo":       str(row[9] or ""),
        "area":           str(row[10] or "") if len(row) > 10 else "",
    })

with open("data.json", "w", encoding="utf-8") as f:
    json.dump(records, f, ensure_ascii=False, indent=2)

print(f"data.json gerado com {len(records)} registros.")
# mostra amostra dos meses para validar
meses = sorted(set(r['mes'] for r in records))
print(f"Meses encontrados: {meses}")
