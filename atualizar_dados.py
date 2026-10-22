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
    if not fa or 'comparativo' in fa.lower() or 'valores_coluna' in fa.lower():
        continue
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

# 6. CONFIGURAÇÕES ADICIONAIS (Créditos Inforow)
creditos_dinamicos = 0.0
if 'CONFIG' in wb.sheetnames:
    sheet_c = wb['CONFIG']
    rows_c = list(sheet_c.iter_rows(values_only=True))
    if len(rows_c) > 1:
        creditos_dinamicos = parse_num(rows_c[1][0])

dados_completos = {
    "base_processamentos": base_proc,
    "backlog": backlog,
    "falhas_plantio": falhas,
    "metas_globais": metas,
    "acompanhamento": acomp,
    "creditos_inforow": creditos_dinamicos,
    "data_compilacao": datetime.now().strftime("%d/%m/%Y às %H:%M")
}

with open(caminho_js, "w", encoding="utf-8") as f:
    f.write("window.DADOS_CLEALCO_EMBARCADOS = " + json.dumps(dados_completos, ensure_ascii=False, indent=2) + ";\n")

print(f"dados_injetados.js criado com sucesso! {len(base_proc)} voos, {len(backlog)} backlog, {len(falhas)} falhas.")

# ================= 6. MOTOR GEOGRÁFICO (MAPA OPERACIONAL DE VOOS) =================
try:
    from osgeo import ogr
    import os, time
    
    t_mapa = time.time()
    pasta_shp = r"W:\11-SHAPES TEMATICOS\09-SHP ATUALIZADO"
    caminho_shp = os.path.join(pasta_shp, "SHP_CLEALCO.shp")
    caminho_mapa_js = r"W:\00-SITE QUALIDADE APLICACAO\Voo_e_processamentosparagithub_2.0\dados_mapa_voos.js"
    
    # Busca automaticamente o shapefile mais recente ou atualizado na pasta
    if os.path.exists(pasta_shp):
        shps_encontrados = [
            os.path.join(pasta_shp, f) for f in os.listdir(pasta_shp)
            if f.lower().endswith('.shp') and not f.lower().startswith('~$')
        ]
        if shps_encontrados:
            # Ordena decrescente por data de modificação
            shps_encontrados.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            # Prioriza SHP_CLEALCO se existir ou pega o mais recente
            shp_clealco_cand = [s for s in shps_encontrados if 'CLEALCO' in os.path.basename(s).upper()]
            caminho_shp = shp_clealco_cand[0] if shp_clealco_cand else shps_encontrados[0]
            dt_mod = datetime.fromtimestamp(os.path.getmtime(caminho_shp)).strftime('%d/%m/%Y %H:%M')
            print(f"Shapefile detectado automaticamente: {os.path.basename(caminho_shp)} (Modificado em: {dt_mod})")
    
    if os.path.exists(caminho_shp):
        print(f"Processando Shapefile ({os.path.basename(caminho_shp)}) para o Mapa de Voos...")
        
        # Mapa de voados
        fas_voadas = set(p['fa'].strip().lstrip('0') for p in base_proc if p.get('fa'))
        
        # Dicionário do Backlog indexado por (FA, Talhao)
        backlog_dict = {}
        for b in backlog:
            fa_norm = str(b.get('fa') or '').strip().lstrip('0')
            tal_norm = str(b.get('talhao') or '').strip().lstrip('0')
            if fa_norm:
                backlog_dict[(fa_norm, tal_norm)] = {
                    "fa": b.get('fa'),
                    "fazenda": b.get('fazenda'),
                    "talhao": b.get('talhao'),
                    "tipo": b.get('tipo', 'Terceiro'),
                    "ha": b.get('ha', 0.0),
                    "data_plantio": b.get('data', ''),
                    "dias_brota": int(b.get('dias')) if str(b.get('dias')).isdigit() else 0,
                    "status_prazo": b.get('status_prazo', ''),
                    "porte_cm": "",
                    "variedade": "",
                    "semaforo": b.get('status_prazo', '')
                }
        
        # Enriquecer com planilhas de vistoria/porte se existirem
        pastas_planilhas = [
            r"M:\10-Planilhas\Voos Pós Plantio - Analise de Falhas\2026\BASES - RECEBIDAS",
            r"W:\00-SITE QUALIDADE APLICACAO\Voo_e_processamentosparagithub_2.0\doc_planilhas"
        ]
        
        for pasta in pastas_planilhas:
            if not os.path.exists(pasta): continue
            for arq in os.listdir(pasta):
                if arq.endswith('.xlsx') and not arq.startswith('~$'):
                    p_full = os.path.join(pasta, arq)
                    try:
                        wb_p = openpyxl.load_workbook(p_full, read_only=True, data_only=True)
                        if 'VOOS' in wb_p.sheetnames:
                            ws_p = wb_p['VOOS']
                            rows_p = list(ws_p.iter_rows(values_only=True))
                            h_idx = 3
                            for r_i, r in enumerate(rows_p[:6]):
                                if any('TALH' in str(c).upper() for c in r if c):
                                    h_idx = r_i
                                    break
                            headers = [str(c).upper().strip() if c else '' for c in rows_p[h_idx]]
                            idx_fa = next((i for i, h in enumerate(headers) if 'F.A' in h or h == 'FA' or 'FAZENDA' in h), 1)
                            idx_tal = next((i for i, h in enumerate(headers) if 'TALH' in h), 4)
                            idx_porte = next((i for i, h in enumerate(headers) if 'PORTE' in h), -1)
                            idx_var = next((i for i, h in enumerate(headers) if 'VARIEDADE' in h), -1)
                            idx_sem = next((i for i, h in enumerate(headers) if 'SEMAFARO' in h or 'SEMAFORO' in h), -1)
                            
                            for r in rows_p[h_idx+1:]:
                                fa_v = str(r[idx_fa] or '').strip().lstrip('0')
                                tal_v = str(r[idx_tal] or '').strip().lstrip('0')
                                k = (fa_v, tal_v)
                                if k in backlog_dict:
                                    if idx_porte != -1 and r[idx_porte] and not backlog_dict[k]["porte_cm"]:
                                        backlog_dict[k]["porte_cm"] = str(r[idx_porte]).strip()
                                    if idx_var != -1 and r[idx_var] and not backlog_dict[k]["variedade"]:
                                        backlog_dict[k]["variedade"] = str(r[idx_var]).strip()
                                    if idx_sem != -1 and r[idx_sem]:
                                        backlog_dict[k]["semaforo"] = str(r[idx_sem]).strip()
                    except: pass
        
        # Leitura do SHP
        ds = ogr.Open(caminho_shp)
        layer = ds.GetLayer()
        features_geojson = []
        matched_keys = set()
        
        shp_itens = []
        for feat in layer:
            geom = feat.GetGeometryRef()
            if not geom: continue
            faz_code = str(feat.GetField('FAZENDA') or '').strip().lstrip('0')
            tal_code = str(feat.GetField('TALHAO') or '').strip().lstrip('0')
            nome_f = str(feat.GetField('NOME_FAZ') or '').strip().upper()
            shp_itens.append({
                "fa": faz_code,
                "tal": tal_code,
                "nome": nome_f,
                "feat": feat
            })
            
        for item in shp_itens:
            faz = item["fa"]
            tal = item["tal"]
            nome = item["nome"]
            feat = item["feat"]
            
            k = (faz, tal)
            bk = backlog_dict.get(k)
            
            # Fallback por nome da fazenda se nao achou por codigo FA
            if not bk:
                for b_k, b_val in backlog_dict.items():
                    if b_k not in matched_keys and b_k[1] == tal:
                        b_nome = b_val["fazenda"].upper()
                        if b_nome in nome or nome in b_nome:
                            bk = b_val
                            k = b_k
                            break
            
            status_voo = "OUTRO"
            if bk:
                matched_keys.add(k)
                dias = bk["dias_brota"]
                if dias > 90 or "atrasado" in bk["status_prazo"].lower():
                    status_voo = "CRITICO"
                elif dias >= 60:
                    status_voo = "ATENCAO"
                else:
                    status_voo = "NOPRAZO"
            elif faz in fas_voadas:
                status_voo = "VOADO"
            else:
                continue
                
            geom = feat.GetGeometryRef()
            geom_simpl = geom.Simplify(0.00004) # ~4m de precisao para manter arquivo leve
            geojson_geom = json.loads(geom_simpl.ExportToJson())
            
            props = {
                "fa": bk["fa"] if bk else feat.GetField('FAZENDA'),
                "fazenda": bk["fazenda"] if bk else feat.GetField('NOME_FAZ'),
                "talhao": bk["talhao"] if bk else feat.GetField('TALHAO'),
                "area_ha": bk["ha"] if bk else round(float(feat.GetField('AREA_TOTAL') or 0), 2),
                "status": status_voo,
                "tipo": bk["tipo"] if bk else ("Própria" if "PROP" in str(feat.GetField('ZONA') or '') else "Terceiro"),
                "dias_brota": bk["dias_brota"] if bk else 0,
                "porte_cm": bk["porte_cm"] if bk else "",
                "variedade": bk["variedade"] if bk else "",
                "data_plantio": bk["data_plantio"] if bk else ""
            }
            
            features_geojson.append({
                "type": "Feature",
                "properties": props,
                "geometry": geojson_geom
            })
            
        geo_col = {
            "type": "FeatureCollection",
            "features": features_geojson
        }
        
        with open(caminho_mapa_js, "w", encoding="utf-8") as f:
            f.write("window.DADOS_MAPA_VOOS = " + json.dumps(geo_col, ensure_ascii=False) + ";\n")
            
        mb = os.path.getsize(caminho_mapa_js) / (1024*1024)
        print(f"dados_mapa_voos.js gerado com sucesso! {len(features_geojson)} talhões mapeados ({len(matched_keys)} backlog cruzados) - {mb:.2f} MB em {time.time()-t_mapa:.2f}s")
    else:
        print(f"Aviso: Shapefile {caminho_shp} não encontrado.")
except Exception as e:
    print(f"Aviso: Erro ao gerar dados_mapa_voos.js: {e}")

