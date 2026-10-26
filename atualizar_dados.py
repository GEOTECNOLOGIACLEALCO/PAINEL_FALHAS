import openpyxl, json, os, base64, math, re, shutil, time, urllib.parse
from datetime import datetime
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes

CHAVE_ACESSO_CLEALCO = "Clealco@Geo2026"

def criptografar_payload(data_str, password=CHAVE_ACESSO_CLEALCO):
    salt = os.urandom(16)
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
    )
    key = kdf.derive(password.encode('utf-8'))
    aesgcm = AESGCM(key)
    iv = os.urandom(12)
    ciphertext = aesgcm.encrypt(iv, data_str.encode('utf-8'), None)
    return {
        "c": True,
        "s": base64.b64encode(salt).decode('utf-8'),
        "i": base64.b64encode(iv).decode('utf-8'),
        "d": base64.b64encode(ciphertext).decode('utf-8')
    }

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

# 6. CONFIGURAÇÕES ADICIONAIS (Créditos Inforow e Senha de Acesso)
creditos_dinamicos = 0.0
senha_config = CHAVE_ACESSO_CLEALCO

if 'CONFIG' in wb.sheetnames:
    sheet_c = wb['CONFIG']
    rows_c = list(sheet_c.iter_rows(values_only=True))
    if len(rows_c) > 0:
        headers_c = [str(h or '').lower().strip() for h in rows_c[0]]
        # Lê creditos_inforow
        idx_cred = next((i for i, h in enumerate(headers_c) if 'credito' in h or 'inforow' in h), 0)
        if len(rows_c) > 1 and len(rows_c[1]) > idx_cred:
            creditos_dinamicos = parse_num(rows_c[1][idx_cred])
        
        # Lê senha_acesso se existir na coluna
        idx_senha = next((i for i, h in enumerate(headers_c) if 'senha' in h or 'chave' in h or 'password' in h), -1)
        if idx_senha != -1 and len(rows_c) > 1 and len(rows_c[1]) > idx_senha and rows_c[1][idx_senha]:
            senha_config = str(rows_c[1][idx_senha]).strip()
            print(f"Senha de acesso lida diretamente da aba CONFIG do Excel.")

CHAVE_ACESSO_CLEALCO = senha_config

dados_completos = {
    "base_processamentos": base_proc,
    "backlog": backlog,
    "falhas_plantio": falhas,
    "metas_globais": metas,
    "acompanhamento": acomp,
    "creditos_inforow": creditos_dinamicos,
    "data_compilacao": datetime.now().strftime("%d/%m/%Y às %H:%M")
}

dados_str = json.dumps(dados_completos, ensure_ascii=False)
payload_cifrado = criptografar_payload(dados_str, CHAVE_ACESSO_CLEALCO)

with open(caminho_js, "w", encoding="utf-8") as f:
    f.write("window.DADOS_CLEALCO_EMBARCADOS = " + json.dumps(payload_cifrado, ensure_ascii=False) + ";\n")

print(f"dados_injetados.js criptografado com sucesso! {len(base_proc)} voos, {len(backlog)} backlog, {len(falhas)} falhas.")

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
        
        # 1. Varredura automática das pastas de imagens de voo para localizar os KMLs de talhões
        pastas_imagens = [
            r"W:\01-DRONE\01-IMAGENS VOO\2026\1 - Não Processados",
            r"W:\01-DRONE\01-IMAGENS VOO\2026\2 - Falhas e Linhas"
        ]
        
        kmls_por_fa = {}
        pastas_voo_todas = {} # Todas as pastas de voo detectadas (com ou sem KML)
        for base_img in pastas_imagens:
            if not os.path.exists(base_img): continue
            for pasta_faz in os.listdir(base_img):
                caminho_faz = os.path.join(base_img, pasta_faz)
                if not os.path.isdir(caminho_faz): continue
                m = re.match(r'^(\d+)', pasta_faz.strip())
                fa_code = m.group(1).lstrip('0') if m else None
                if not fa_code: continue
                
                pastas_voo_todas[fa_code] = {
                    'pasta': pasta_faz,
                    'revoo': 'terminar' in pasta_faz.lower()
                }
                
                achados = []
                for arq in os.listdir(caminho_faz):
                    if arq.lower().endswith('.kml') and ('talhao' in arq.lower() or 'talho' in arq.lower()):
                        achados.append(os.path.join(caminho_faz, arq))
                for sub in ['4 - Nuvens de Pontos', '3 - Linhas_Falhas', '3 - Linhas e Falhas']:
                    sub_p = os.path.join(caminho_faz, sub)
                    if os.path.exists(sub_p) and os.path.isdir(sub_p):
                        for arq in os.listdir(sub_p):
                            p_arq = os.path.join(sub_p, arq)
                            if os.path.isfile(p_arq) and arq.lower().endswith('.kml') and ('talhao' in arq.lower() or 'talho' in arq.lower()):
                                achados.append(p_arq)
                            elif os.path.isdir(p_arq) and ('parte' in arq.lower() or 'talhao' in arq.lower()):
                                for sub_sub_arq in os.listdir(p_arq):
                                    if sub_sub_arq.lower().endswith('.kml') and ('talhao' in sub_sub_arq.lower() or 'talho' in sub_sub_arq.lower()):
                                        achados.append(os.path.join(p_arq, sub_sub_arq))
                if achados:
                    kmls_por_fa[fa_code] = {
                        'pasta': pasta_faz,
                        'kmls': achados,
                        'revoo': 'terminar' in pasta_faz.lower()
                    }
                    
        print(f"Varredura concluída: {len(pastas_voo_todas)} pastas de voo ({len(kmls_por_fa)} com KML gerado).")
        
        # 2. Carrega as geometrias dos KMLs
        geoms_por_fa = {}
        for fa_c, info in kmls_por_fa.items():
            geoms_por_fa[fa_c] = []
            for kml_p in info['kmls']:
                try:
                    ds_k = ogr.Open(kml_p)
                    if ds_k:
                        lyr_k = ds_k.GetLayer()
                        for f_k in lyr_k:
                            g_k = f_k.GetGeometryRef()
                            if g_k: geoms_por_fa[fa_c].append(g_k.Clone())
                except: pass
        
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
        
        # Enriquecer com planilhas de vistoria/porte e extrair talhões voados oficiais
        pastas_planilhas = [
            r"M:\10-Planilhas\Voos Pós Plantio - Analise de Falhas\2026\BASES - RECEBIDAS",
            r"W:\00-SITE QUALIDADE APLICACAO\Voo_e_processamentosparagithub_2.0\doc_planilhas"
        ]
        
        talhoes_voados_bases = {} # (fa, tal) voados segundo as bases oficiais de voo (Próprias + Terceiros)
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
                                if fa_v and tal_v:
                                    talhoes_voados_bases.setdefault(fa_v, set()).add(tal_v)
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
        fas_no_shp = set()
        
        shp_itens = []
        for feat in layer:
            geom = feat.GetGeometryRef()
            if not geom: continue
            faz_code = str(feat.GetField('FAZENDA') or '').strip().lstrip('0')
            tal_code = str(feat.GetField('TALHAO') or '').strip().lstrip('0')
            nome_f = str(feat.GetField('NOME_FAZ') or '').strip().upper()
            fas_no_shp.add(faz_code)
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
            
            status_voo = "OUTRO"
            obs_voo = ""
            
            # Caso 1: Fazenda possui KML de voo (cruza espacialmente com o SHP)
            if faz in geoms_por_fa and geoms_por_fa[faz]:
                geom_item = feat.GetGeometryRef()
                tg_area = geom_item.GetArea() if geom_item else 0.0
                max_frac = 0.0
                if tg_area > 0:
                    for kg in geoms_por_fa[faz]:
                        if geom_item.Intersects(kg):
                            inter = geom_item.Intersection(kg)
                            if inter and not inter.IsEmpty():
                                f_area = inter.GetArea() / tg_area
                                if f_area > max_frac:
                                    max_frac = f_area
                if max_frac >= 0.15:
                    status_voo = "VOADO"
                    obs_voo = f"Voo Realizado ({int(round(max_frac*100))}% de cobertura KML)"
                elif kmls_por_fa[faz]['revoo']:
                    status_voo = "ATENCAO"
                    obs_voo = "Pendente de Voo / Terminar Outras Partes"
                elif bk:
                    matched_keys.add(k)
                    dias = bk["dias_brota"]
                    status_voo = "CRITICO" if dias > 90 or "atrasado" in bk["status_prazo"].lower() else ("ATENCAO" if dias >= 60 else "NOPRAZO")
                else:
                    continue
                if bk:
                    matched_keys.add(k)
            # Caso 1B: Fazenda possui pasta de Voo mas ainda sem KML (cruza com as Bases Oficiais de Voo)
            elif faz in pastas_voo_todas:
                info_p = pastas_voo_todas[faz]
                tals_voados_faz = talhoes_voados_bases.get(faz, set())
                if tal in tals_voados_faz:
                    status_voo = "VOADO"
                    obs_voo = "Voo Realizado (Confirmado na Base de Voo)"
                elif info_p['revoo']:
                    status_voo = "ATENCAO"
                    obs_voo = "Pendente de Voo / Terminar Outras Partes"
                elif bk:
                    matched_keys.add(k)
                    dias = bk["dias_brota"]
                    status_voo = "CRITICO" if dias > 90 or "atrasado" in bk["status_prazo"].lower() else ("ATENCAO" if dias >= 60 else "NOPRAZO")
                else:
                    continue
                if bk:
                    matched_keys.add(k)
            # Caso 2: Talhão está no Backlog oficial
            elif bk:
                matched_keys.add(k)
                dias = bk["dias_brota"]
                if dias > 90 or "atrasado" in bk["status_prazo"].lower():
                    status_voo = "CRITICO"
                elif dias >= 60:
                    status_voo = "ATENCAO"
                else:
                    status_voo = "NOPRAZO"
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
                "data_plantio": bk["data_plantio"] if bk else "",
                "obs": obs_voo
            }
            
            features_geojson.append({
                "type": "Feature",
                "properties": props,
                "geometry": geojson_geom
            })
            
        # Inserção direta de KMLs para fazendas que NÃO constam no SHP cadastral (ex: FA 200 Terra Nova II, 106, 240, 285)
        for fa_fora, info_fora in kmls_por_fa.items():
            if fa_fora not in fas_no_shp:
                # Se houver arquivo consolidado com o prefixo da FA (ex: '200 - talhao.kml'), prioriza ele para não duplicar subpartes
                kmls_validos = info_fora['kmls']
                consolidado = [k for k in kmls_validos if os.path.basename(k).startswith(fa_fora) and 'parte' not in os.path.basename(os.path.dirname(k)).lower()]
                if consolidado:
                    kmls_validos = consolidado

                for kml_f in kmls_validos:
                    try:
                        ds_direct = ogr.Open(kml_f)
                        if ds_direct:
                            lyr_direct = ds_direct.GetLayer()
                            faz_nome_limpo = info_fora['pasta'].split('-')[1].strip() if '-' in info_fora['pasta'] else info_fora['pasta']
                            print(f"Inserindo {lyr_direct.GetFeatureCount()} talhões diretos do KML: {faz_nome_limpo} (FA {fa_fora})")
                            for idx_d, feat_d in enumerate(lyr_direct):
                                g_d = feat_d.GetGeometryRef()
                                if not g_d: continue
                                env_d = g_d.GetEnvelope()
                                lat_mid_d = (env_d[2] + env_d[3]) / 2.0
                                area_ha_d = round(g_d.GetArea() * (111319.0 ** 2) * math.cos(math.radians(lat_mid_d)) / 10000.0, 2)
                                nome_tal_d = feat_d.GetField("Name") or f"{idx_d + 1}"
                                g_simpl_d = g_d.Simplify(0.00004)
                                features_geojson.append({
                                    "type": "Feature",
                                    "properties": {
                                        "fa": fa_fora,
                                        "fazenda": faz_nome_limpo,
                                        "talhao": str(nome_tal_d),
                                        "area_ha": area_ha_d,
                                        "status": "VOADO",
                                        "tipo": "Própria",
                                        "dias_brota": 0,
                                        "porte_cm": "",
                                        "variedade": "",
                                        "data_plantio": "",
                                        "obs": f"Voo Realizado ({os.path.basename(kml_f)})"
                                    },
                                    "geometry": json.loads(g_simpl_d.ExportToJson())
                                })
                    except Exception as ex_direct:
                        print(f"Erro inserindo KML direto {info_fora['pasta']}: {ex_direct}")
            
        geo_col = {
            "type": "FeatureCollection",
            "features": features_geojson
        }
        
        geo_str = json.dumps(geo_col, ensure_ascii=False)
        payload_mapa_cifrado = criptografar_payload(geo_str, CHAVE_ACESSO_CLEALCO)
        
        with open(caminho_mapa_js, "w", encoding="utf-8") as f:
            f.write("window.DADOS_MAPA_VOOS = " + json.dumps(payload_mapa_cifrado, ensure_ascii=False) + ";\n")
            
        mb = os.path.getsize(caminho_mapa_js) / (1024*1024)
        print(f"dados_mapa_voos.js criptografado com sucesso! {len(features_geojson)} talhões mapeados ({len(matched_keys)} backlog cruzados) - {mb:.2f} MB em {time.time()-t_mapa:.2f}s")
    else:
        print(f"Aviso: Shapefile {caminho_shp} não encontrado.")
except Exception as e:
    print(f"Aviso: Erro ao gerar dados_mapa_voos.js: {e}")

# -------------------------------------------------------------
# 5. SINCRONIZAÇÃO E CATALOGAÇÃO AUTOMÁTICA DE MAPAS DE SAFRA (PDF)
# -------------------------------------------------------------
def sincronizar_e_catalogar_mapas():
    origem = r"F:\Mapas\01 - Mapas Safra"
    destino = r"W:\00-SITE QUALIDADE APLICACAO\Voo_e_processamentosparagithub_2.0\mapas_safra"
    
    novos_copiados = 0
    atualizados_copiados = 0
    
    # Sincronização 100% segura (Apenas LEITURA de F:)
    if os.path.exists(origem):
        try:
            for root, dirs, files in os.walk(origem):
                rel_path = os.path.relpath(root, origem)
                target_dir = os.path.join(destino, rel_path)
                if not os.path.exists(target_dir):
                    os.makedirs(target_dir, exist_ok=True)
                for f in files:
                    if not f.lower().endswith('.pdf'): continue
                    src_f = os.path.join(root, f)
                    dst_f = os.path.join(target_dir, f)
                    
                    if not os.path.exists(dst_f):
                        shutil.copy2(src_f, dst_f)
                        novos_copiados += 1
                    else:
                        src_stat = os.stat(src_f)
                        dst_stat = os.stat(dst_f)
                        # Copia se data for mais recente ou tamanho diferente
                        if src_stat.st_mtime > dst_stat.st_mtime + 2 or src_stat.st_size != dst_stat.st_size:
                            shutil.copy2(src_f, dst_f)
                            atualizados_copiados += 1
            if novos_copiados > 0 or atualizados_copiados > 0:
                print(f"[MAPAS] Sincronização: {novos_copiados} novos mapas, {atualizados_copiados} mapas atualizados de F:.")
            else:
                print("[MAPAS] Todos os mapas já estão sincronizados com a pasta oficial F:.")
        except Exception as ex_sync:
            print(f"[MAPAS] Aviso durante sincronização de F:: {ex_sync}")
    else:
        print("[MAPAS] Disco F: não acessível no momento. Utilizando mapas locais da pasta mapas_safra.")
        
    mapas_catalogo = []
    agora = time.time()
    
    if os.path.exists(destino):
        for root, dirs, files in os.walk(destino):
            rel_dir = os.path.relpath(root, destino).replace('\\', '/')
            partes = rel_dir.split('/')
            categoria = partes[0] if partes and partes[0] != '.' else 'Geral'
            subpasta = partes[1] if len(partes) > 1 else ''
            
            for f in files:
                if not f.lower().endswith('.pdf'): continue
                f_path = os.path.join(root, f)
                stat = os.stat(f_path)
                mtime = stat.st_mtime
                dias = (agora - mtime) / 86400.0
                dt_str = datetime.fromtimestamp(mtime).strftime('%d/%m/%Y %H:%M')
                
                nome_sem_ext = f[:-4].strip()
                fa = ''
                fazenda = nome_sem_ext
                prop = ''
                
                m = re.match(r'^([0-9A-Za-z\s]+?)\s*-\s*(.+)$', nome_sem_ext)
                if m:
                    fa = m.group(1).strip()
                    resto = m.group(2).strip()
                    m_prop = re.search(r'\((.+?)\)$', resto)
                    if m_prop:
                        prop = m_prop.group(1).strip()
                        fazenda = resto[:m_prop.start()].strip()
                    else:
                        fazenda = resto
                
                if rel_dir == '.':
                    rel_url = f'mapas_safra/{f}'
                    sp_path = f
                else:
                    rel_url = f'mapas_safra/{rel_dir}/{f}'
                    sp_path = f'{rel_dir}/{f}'
                
                # Gera URL com o link de acesso livre/convidado (sem necessidade de login corporativo)
                token_guest_base = "https://clealco-my.sharepoint.com/:f:/g/personal/fabiano_castro_clealco_com_br/IgCSHZ5yRz0xQ4PYlEqfB59jATIXumbfQOUSQbrvpRYwANk"
                encoded_id_cat = urllib.parse.quote(f"/personal/fabiano_castro_clealco_com_br/Documents/Mapas/{categoria}")
                url_pasta_sharepoint = f"{token_guest_base}?id={encoded_id_cat}&e=adZ905"
                
                encoded_id_file = urllib.parse.quote(f"/personal/fabiano_castro_clealco_com_br/Documents/Mapas/{sp_path}")
                url_sharepoint = f"{token_guest_base}?id={encoded_id_file}&e=adZ905"

                mapas_catalogo.append({
                    'arquivo': f,
                    'url': rel_url,
                    'url_sharepoint': url_sharepoint,
                    'url_pasta_sharepoint': url_pasta_sharepoint,
                    'categoria': categoria,
                    'subpasta': subpasta,
                    'fa': fa,
                    'fazenda': fazenda,
                    'proprietario': prop,
                    'data_modificacao': dt_str,
                    'timestamp': int(mtime),
                    'tamanho_kb': round(stat.st_size / 1024, 1),
                    'dias_atras': round(dias, 1),
                    'is_novo': dias <= 7,
                    'is_hoje': dias <= 1
                })
                
    mapas_catalogo.sort(key=lambda x: x['timestamp'], reverse=True)
    
    caminho_js_mapas = os.path.join(os.path.dirname(destino), "dados_mapas_safra.js")
    with open(caminho_js_mapas, "w", encoding="utf-8") as f_out:
        agora_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        f_out.write("// Catalogo Automatico de Mapas de Safra em PDF\n")
        f_out.write(f"// Gerado em: {agora_str}\n")
        f_out.write(f"// Total de Mapas: {len(mapas_catalogo)}\n")
        f_out.write("window.DADOS_MAPAS_SAFRA = ")
        json.dump(mapas_catalogo, f_out, ensure_ascii=False, indent=2)
        f_out.write(";\n")
        
    novos = [m for m in mapas_catalogo if m['is_novo']]
    print(f"[MAPAS] Catálogo atualizado em dados_mapas_safra.js: {len(mapas_catalogo)} mapas ({len(novos)} novos/atualizados nos últimos 7 dias).")

sincronizar_e_catalogar_mapas()

