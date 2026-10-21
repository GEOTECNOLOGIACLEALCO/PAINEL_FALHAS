import openpyxl, json
from datetime import datetime

caminho_xlsx = r"W:\00-SITE QUALIDADE APLICACAO\Voo_e_processamentosparagithub_2.0\dados.xlsx"
caminho_js = r"W:\00-SITE QUALIDADE APLICACAO\Voo_e_processamentosparagithub_2.0\dados_injetados.js"

wb = openpyxl.load_workbook(caminho_xlsx, data_only=True)

def parse_num(val):
    if val is None or val == '': return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip().replace('.', '').replace(',', '.') if str(val).count(',') == 1 and str(val).count('.') > 0 else str(val).strip().replace(',', '.')
    try: return float(s)
    except: return 0.0

def format_data(val):
    if not val: return "N/D"
    if hasattr(val, 'strftime'): return val.strftime('%d/%m/%Y')
    s = str(val).strip()
    if 'T' in s: s = s.split('T')[0]
    p = s.replace('-', '/').split(' ')[0].split('/')
    if len(p) == 3:
        if len(p[0]) == 4: return f"{p[2].zfill(2)}/{p[1].zfill(2)}/{p[0]}"
        return f"{p[0].zfill(2)}/{p[1].zfill(2)}/{p[2]}"
    return s

# 1. BASE_PROCESSAMENTOS
sheet_base = wb['BASE_PROCESSAMENTOS']
rows_base = list(sheet_base.iter_rows(values_only=True))
headers_base = rows_base[0]

base_proc = []
for r in rows_base[1:]:
    fa = str(r[0] or '').strip()
    if not fa: continue
    fazenda = str(r[1] or 'NÃO INFORMADO').strip()
    base_voo = str(r[2] or 'N/D').strip()
    coord = str(r[3] or 'N/D').strip()
    dt_voo = format_data(r[4])
    tif = str(r[5] or 'PENDENTE').upper().strip()
    dt_proc = format_data(r[6])
    ha_voo = parse_num(r[7])
    ha_linhas = parse_num(r[8])
    linhas = str(r[9] or 'PENDENTE').upper().strip()
    exp = str(r[10] or 'PENDENTE').upper().strip()
    
    bd_falhas = "FEITO" if "FEITO" in linhas else "PENDENTE"
    bd_pisoteio = "FEITO"
    
    base_proc.append({
        "fa": fa,
        "fazenda": fazenda,
        "base": base_voo,
        "coord": coord,
        "data_voo": dt_voo,
        "data_proc": dt_proc,
        "tif": "FEITO" if "FEITO" in tif else "PENDENTE",
        "linhas": "FEITO" if "FEITO" in linhas else "PENDENTE",
        "data_linhas": dt_proc,
        "ha_voo": round(ha_voo, 2),
        "area_kml": round(ha_linhas, 2),
        "bd_falhas": bd_falhas,
        "bd_pisoteio": bd_pisoteio,
        "exp": "FEITO" if "FEITO" in exp else "Pendente"
    })

# Inverter para que as mais recentes apareçam primeiro
base_proc.reverse()

# 2. BACKLOG
sheet_bk = wb['BACKLOG']
rows_bk = list(sheet_bk.iter_rows(values_only=True))
backlog = []
for r in rows_bk[1:]:
    fa = str(r[0] or '').strip()
    if not fa: continue
    backlog.append({
        "fa": fa,
        "fazenda": str(r[1] or '').strip(),
        "talhao": str(r[2] or '').strip(),
        "tipo": str(r[3] or 'Terceiro').strip(),
        "ha": parse_num(r[4]),
        "data": format_data(r[5]),
        "dias": r[6] or 0,
        "status_prazo": str(r[7] or 'N/D')
    })

# 3. FALHAS_PLANTIO
sheet_f = wb['FALHAS_PLANTIO']
rows_f = list(sheet_f.iter_rows(values_only=True))
falhas = []
for r in rows_f[1:]:
    fa = str(r[0] or '').strip()
    if not fa: continue
    falhas.append({
        "FA": fa,
        "Fazenda": str(r[1] or '').strip(),
        "Tipo": str(r[2] or 'Terceiro').strip(),
        "Area_Ha": parse_num(r[3]),
        "Comp_Linhas": parse_num(r[4]),
        "Comp_Falhas": parse_num(r[5]),
        "Porc_Falha": parse_num(r[6]),
        "Indice_Falha": str(r[7] or 'Aceitável')
    })

# 4. METAS_GLOBAIS
sheet_m = wb['METAS_GLOBAIS']
rows_m = list(sheet_m.iter_rows(values_only=True))
metas = {}
for r in rows_m[1:]:
    ano = str(r[0] or '').strip()
    if not ano: continue
    metas[ano] = {
        "propria": parse_num(r[1]),
        "terceiros": parse_num(r[2]),
        "acompanhamento_propria": parse_num(r[3]),
        "acompanhamento_terceiros": parse_num(r[4])
    }

# 5. ACOMPANHAMENTO
acomp = []
if 'ACOMPANHAMENTO' in wb.sheetnames:
    sheet_a = wb['ACOMPANHAMENTO']
    rows_a = list(sheet_a.iter_rows(values_only=True))
    for r in rows_a[1:]:
        acomp.append({
            "Ano": str(r[0] or '2026'),
            "Categoria": str(r[1] or ''),
            "Segmento": str(r[2] or ''),
            "Previsto": parse_num(r[3]),
            "Realizado": parse_num(r[4]),
            "Semana": parse_num(r[5]),
            "Mes": parse_num(r[6])
        })

dados_completos = {
    "base_processamentos": base_proc,
    "backlog": backlog,
    "falhas_plantio": falhas,
    "metas_globais": metas,
    "acompanhamento": acomp,
    "creditos_inforow": 11195.23,
    "data_compilacao": datetime.now().strftime("%d/%m/%Y às %H:%M")
}

with open(caminho_js, "w", encoding="utf-8") as f:
    f.write("window.DADOS_CLEALCO_EMBARCADOS = " + json.dumps(dados_completos, ensure_ascii=False, indent=2) + ";\n")

print(f"dados_injetados.js criado com sucesso! {len(base_proc)} voos, {len(backlog)} backlog, {len(falhas)} falhas.")
