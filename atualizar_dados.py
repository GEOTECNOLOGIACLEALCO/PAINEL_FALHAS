import openpyxl, json, os, base64, math, re, shutil, time, urllib.parse, glob, csv
from datetime import datetime, timedelta
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
fas_em_falhas = set()
for r in rows_f[1:]:
    fa = str(r[0] or '').strip()
    if not fa or 'comparativo' in fa.lower() or 'valores_coluna' in fa.lower():
        continue
    fas_em_falhas.add(fa)
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

# Complementa em tempo real com arquivos do BD Falhas e consolida partes adicionais (_P2, etc.)
dir_falhas_live = r"W:\01-DRONE\08-BANCO DE DADOS\2026\Falhas"
if os.path.exists(dir_falhas_live):
    mapa_faz_nome = {item["fa"]: item["fazenda"] for item in base_proc}
    mapa_falhas_por_fa = {f["FA"]: f for f in falhas}
    for f_csv in glob.glob(os.path.join(dir_falhas_live, "*.csv")):
        b_name = os.path.basename(f_csv).replace('.csv', '').replace('.CSV', '').strip()
        if not re.match(r'^\d+', b_name) or 'comparativo' in b_name.lower() or 'valores' in b_name.lower():
            continue
        
        m_base = re.match(r'^(\d+)', b_name)
        b_fa = m_base.group(1) if m_base else b_name

        # Se a fazenda já está na base compilada de falhas, não adiciona novamente
        if b_fa in mapa_falhas_por_fa:
            continue

        num_clean = re.sub(r'\D', '', b_fa)
        tipo_calc = "Própria" if (num_clean.startswith('20') and len(num_clean) == 6) else "Terceiro"
        nome_f = mapa_faz_nome.get(b_fa, "N/D")
        
        ha_val, comp_l, comp_f, p_falha = 0.0, 0.0, 0.0, 0.0
        try:
            with open(f_csv, 'r', encoding='latin1', errors='ignore') as fp:
                r_all = list(csv.reader(fp))
                if len(r_all) > 1:
                    r1 = r_all[1]
                    if len(r1) > 2:
                        ha_val = float(r1[2].strip().replace(',', '.'))
                    c_hdr = [str(c).strip().upper() for c in r_all[0]]
                    i_l = next((i for i, c in enumerate(c_hdr) if 'COMP_LINHAS' in c or 'COMP LINHAS' in c), None)
                    i_f = next((i for i, c in enumerate(c_hdr) if 'COMP_FALHAS' in c or 'COMP FALHAS' in c), None)
                    i_p = next((i for i, c in enumerate(c_hdr) if 'PORC_FALHA' in c or 'PORC FALHA' in c), 5)
                    if i_l is not None and len(r1) > i_l:
                        comp_l = float(str(r1[i_l]).replace(',', '.'))
                    if i_f is not None and len(r1) > i_f:
                        comp_f = float(str(r1[i_f]).replace(',', '.'))
                    if i_p is not None and len(r1) > i_p:
                        p_falha = float(str(r1[i_p]).replace(',', '.'))

            lim_verde = 8.0 if tipo_calc == "Própria" else 5.0
            ind = 'Aceitável' if p_falha <= lim_verde else ('Atenção' if p_falha <= 10.0 else 'Crítico')
            novo_item = {
                "FA": b_fa,
                "Fazenda": nome_f,
                "Tipo": tipo_calc,
                "Area_Ha": round(ha_val, 2),
                "Comp_Linhas": round(comp_l, 2),
                "Comp_Falhas": round(comp_f, 2),
                "Porc_Falha": p_falha,
                "Indice_Falha": ind
            }
            falhas.append(novo_item)
            mapa_falhas_por_fa[b_fa] = novo_item
        except Exception:
            pass

# Normaliza os Tipos de todas as fazendas na lista de falhas para evitar problemas de acentuação
for f_item in falhas:
    fa_str = str(f_item.get("FA", "")).strip()
    is_p = (fa_str.startswith('20') and len(fa_str) == 6) or str(f_item.get("Tipo", "")).lower().startswith("pr")
    f_item["Tipo"] = "Própria" if is_p else "Terceiro"

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

# 🛡️ RECALCULO DINÂMICO EM TEMPO REAL: Falhas no plantio (2026)
# Garante que qualquer novo CSV em W:\01-DRONE\08-BANCO DE DADOS\2026\Falhas entre imediatamente na conta!
dir_bd_falhas_2026 = r"W:\01-DRONE\08-BANCO DE DADOS\2026\Falhas"
if os.path.exists(dir_bd_falhas_2026):
    try:
        hoje_now = datetime.now()
        mes_atual = hoje_now.month
        ano_atual = hoje_now.year
        inicio_sem = hoje_now - timedelta(days=hoje_now.weekday())
        inicio_sem = inicio_sem.replace(hour=0, minute=0, second=0, microsecond=0)
        fim_sem = inicio_sem + timedelta(days=6, hours=23, minutes=59, seconds=59)

        # Mapa de datas a partir das pastas de imagens de voo
        mapa_datas_voos = {}
        for d_base in [r"W:\01-DRONE\01-IMAGENS VOO\2026\2 - Falhas e Linhas", r"W:\01-DRONE\01-IMAGENS VOO\2026\1 - Não Processados"]:
            if os.path.exists(d_base):
                for pasta in os.listdir(d_base):
                    c_pasta = os.path.join(d_base, pasta)
                    if os.path.isdir(c_pasta):
                        m_fa = re.search(r'(\d+)', pasta)
                        if m_fa:
                            fa_k = m_fa.group(1).lstrip('0')
                            t_max = os.path.getmtime(c_pasta)
                            try:
                                for arq_sub in os.listdir(c_pasta):
                                    c_sub = os.path.join(c_pasta, arq_sub)
                                    if os.path.isfile(c_sub):
                                        u_sub = arq_sub.upper()
                                        if 'LINHA' in u_sub or u_sub.endswith('.CSV') or u_sub.endswith('.KML'):
                                            t_arq = os.path.getmtime(c_sub)
                                            if t_arq > t_max: t_max = t_arq
                            except: pass
                            if fa_k not in mapa_datas_voos or t_max > mapa_datas_voos[fa_k]:
                                mapa_datas_voos[fa_k] = t_max

        falhas_calc = {
            'Própria': {'r': 0.0, 's': 0.0, 'm': 0.0},
            'Fornecedores': {'r': 0.0, 's': 0.0, 'm': 0.0}
        }

        for arq_csv in glob.glob(os.path.join(dir_bd_falhas_2026, "*.csv")):
            b_name = os.path.basename(arq_csv).replace('.csv', '').replace('.CSV', '').strip()
            if not re.match(r'^\d+', b_name) or 'comparativo' in b_name.lower() or 'valores' in b_name.lower():
                continue
            m_fa = re.match(r'^(\d+)', b_name)
            fa_base = m_fa.group(1) if m_fa else b_name
            num_fa = re.sub(r'\D', '', fa_base)
            k_fa = num_fa.lstrip('0')
            is_prop = (num_fa.startswith('20') and len(num_fa) == 6)
            seg_chave = 'Própria' if is_prop else 'Fornecedores'

            area_ha = 0.0
            try:
                with open(arq_csv, 'r', encoding='latin1', errors='ignore') as f_csv:
                    rdr = csv.reader(f_csv)
                    _ = next(rdr, None)
                    r2 = next(rdr, None)
                    if r2 and len(r2) > 2:
                        area_ha = float(r2[2].strip().replace(',', '.'))
            except:
                continue

            # Data de realização
            dt_r = None
            if k_fa in mapa_datas_voos:
                dt_r = datetime.fromtimestamp(mapa_datas_voos[k_fa])
            else:
                try: dt_r = datetime.fromtimestamp(os.path.getmtime(arq_csv))
                except: pass

            falhas_calc[seg_chave]['r'] += area_ha
            if dt_r:
                if dt_r.year == ano_atual and dt_r.month == mes_atual:
                    falhas_calc[seg_chave]['m'] += area_ha
                if inicio_sem <= dt_r <= fim_sem:
                    falhas_calc[seg_chave]['s'] += area_ha

        # Atualiza ou insere as linhas de Falhas no plantio no acomp
        norm_txt = lambda s: str(s or '').lower().replace('ó', 'o').replace('ô', 'o').replace('í', 'i').replace('ã', 'a').strip()
        for item in acomp:
            if str(item.get("Ano", "")).strip() == '2026' and norm_txt(item.get("Categoria", "")) == 'falhas no plantio':
                seg_it = item.get("Segmento", "")
                if 'prop' in norm_txt(seg_it):
                    item["Realizado"] = round(falhas_calc['Própria']['r'], 2)
                    item["Semana"] = round(falhas_calc['Própria']['s'], 2)
                    item["Mes"] = round(falhas_calc['Própria']['m'], 2)
                elif 'fornec' in norm_txt(seg_it):
                    item["Realizado"] = round(falhas_calc['Fornecedores']['r'], 2)
                    item["Semana"] = round(falhas_calc['Fornecedores']['s'], 2)
                    item["Mes"] = round(falhas_calc['Fornecedores']['m'], 2)
                elif 'consol' in norm_txt(seg_it):
                    item["Realizado"] = round(falhas_calc['Própria']['r'] + falhas_calc['Fornecedores']['r'], 2)
                    item["Semana"] = round(falhas_calc['Própria']['s'] + falhas_calc['Fornecedores']['s'], 2)
                    item["Mes"] = round(falhas_calc['Própria']['m'] + falhas_calc['Fornecedores']['m'], 2)
    except Exception as e:
        print(f"⚠️ Aviso no recálculo dinâmico de falhas: {e}")

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
        pastas_nao_processados = set()
        pastas_falhas_linhas = set()
        
        for base_img in pastas_imagens:
            if not os.path.exists(base_img): continue
            is_np = ('1 - Não Processados' in base_img or '1 - N' in base_img)
            is_fl = ('2 - Falhas e Linhas' in base_img or '2 - F' in base_img)
            for pasta_faz in os.listdir(base_img):
                caminho_faz = os.path.join(base_img, pasta_faz)
                if not os.path.isdir(caminho_faz): continue
                
                fa_matches = re.findall(r'\b\d{3,6}\b', pasta_faz.strip())
                fas_encontradas = [f.lstrip('0') for f in fa_matches if f not in ['2026', '2025', '2024', '2023']]
                
                m = re.match(r'^(\d+)', pasta_faz.strip())
                fa_code = m.group(1).lstrip('0') if m else (fas_encontradas[0] if fas_encontradas else None)
                if not fa_code and not fas_encontradas: continue
                if not fas_encontradas and fa_code:
                    fas_encontradas = [fa_code]
                
                for f_c in fas_encontradas:
                    if is_np:
                        pastas_nao_processados.add(f_c)
                    elif is_fl:
                        pastas_falhas_linhas.add(f_c)
                    
                    if f_c not in pastas_voo_todas:
                        pastas_voo_todas[f_c] = {
                            'pasta': pasta_faz,
                            'revoo': 'terminar' in pasta_faz.lower()
                        }
                    else:
                        if 'terminar' not in pasta_faz.lower():
                            pastas_voo_todas[f_c]['revoo'] = False
                
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
                    for f_c in fas_encontradas:
                        if f_c not in kmls_por_fa:
                            kmls_por_fa[f_c] = {
                                'pasta': pasta_faz,
                                'kmls': list(set(achados)),
                                'revoo': 'terminar' in pasta_faz.lower()
                            }
                        else:
                            for kml_item in achados:
                                if kml_item not in kmls_por_fa[f_c]['kmls']:
                                    kmls_por_fa[f_c]['kmls'].append(kml_item)
                            if 'terminar' not in pasta_faz.lower():
                                kmls_por_fa[f_c]['revoo'] = False
                    
        # Complementa com base_proc oficial
        for item in base_proc:
            fa_p = str(item.get('fa') or '').strip().lstrip('0')
            if fa_p:
                if item.get('tif') == 'PENDENTE':
                    pastas_nao_processados.add(fa_p)
                elif item.get('tif') == 'FEITO' and item.get('linhas') == 'FEITO':
                    pastas_falhas_linhas.add(fa_p)
                    
        print(f"Varredura concluída: {len(pastas_voo_todas)} pastas de voo ({len(kmls_por_fa)} com KML gerado). Fila Não Processados: {len(pastas_nao_processados)} FAs.")
        
        # 2. Carrega as geometrias dos KMLs com Bounding Box pré-calculado
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
                            if g_k: geoms_por_fa[fa_c].append((g_k.Clone(), g_k.GetEnvelope()))
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
        
        # 🔄 Dicionário de REVOO (Talhões ou Fazendas marcados para voar novamente)
        custom_revoo_dict = {}
        if 'REVOO' in wb.sheetnames:
            try:
                ws_rev = wb['REVOO']
                for r_rev in list(ws_rev.iter_rows(values_only=True))[1:]:
                    fa_r = str(r_rev[0] or '').strip().lstrip('0')
                    tal_r = str(r_rev[1] or '').strip().lstrip('0').upper()
                    motivo_r = str(r_rev[2] or 'Voar novamente').strip()
                    if fa_r:
                        custom_revoo_dict[(fa_r, tal_r)] = motivo_r
            except Exception as e_rev:
                print(f"Aviso ao ler aba REVOO: {e_rev}")
        
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
                            idx_realizada = next((i for i, h in enumerate(headers) if 'REALIZADA' in h or 'VOADA' in h), -1)
                            
                            for r in rows_p[h_idx+1:]:
                                fa_v = str(r[idx_fa] or '').strip().lstrip('0')
                                tal_v = str(r[idx_tal] or '').strip().lstrip('0')
                                k = (fa_v, tal_v)
                                if fa_v and tal_v and idx_realizada != -1:
                                    val_real = parse_num(r[idx_realizada])
                                    if val_real > 0 and k not in backlog_dict:
                                        talhoes_voados_bases.setdefault(fa_v, set()).add(tal_v)
                                if k in backlog_dict:
                                    if idx_porte != -1 and r[idx_porte] and not backlog_dict[k]["porte_cm"]:
                                        backlog_dict[k]["porte_cm"] = str(r[idx_porte]).strip()
                                    if idx_var != -1 and r[idx_var] and not backlog_dict[k]["variedade"]:
                                        backlog_dict[k]["variedade"] = str(r[idx_var]).strip()
                                    if idx_sem != -1 and r[idx_sem]:
                                        backlog_dict[k]["semaforo"] = str(r[idx_sem]).strip()
                    except: pass
        
        def round_coords(coords, p=5):
            if isinstance(coords, (float, int)):
                return round(coords, p)
            return [round_coords(x, p) for x in coords]

        # Leitura do SHP
        ds = ogr.Open(caminho_shp)
        layer = ds.GetLayer()
        features_geojson = []
        matched_keys = set()
        fas_no_shp = set()
        fas_com_talhao_voado = set()
        
        for feat in layer:
            geom = feat.GetGeometryRef()
            if not geom: continue
            faz = str(feat.GetField('FAZENDA') or '').strip().lstrip('0')
            fas_no_shp.add(faz)
            tal = str(feat.GetField('TALHAO') or '').strip().lstrip('0')
            
            k = (faz, tal)
            bk = backlog_dict.get(k)

            status_voo = "BASE_CLEALCO"
            obs_voo = "Base Clealco (Área Cadastrada)"
            
            # Caso 1: Fazenda possui KML de voo (cruza espacialmente com o SHP)
            teve_kml_match = False
            if faz in geoms_por_fa and geoms_por_fa[faz]:
                tg_area = geom.GetArea()
                max_frac = 0.0
                if tg_area > 0:
                    env_i = geom.GetEnvelope() # (minX, maxX, minY, maxY)
                    for kg, env_k in geoms_por_fa[faz]:
                        # Bounding box filter ultrarrápido (evita cálculo pesado GEOS)
                        if env_i[0] > env_k[1] or env_i[1] < env_k[0] or env_i[2] > env_k[3] or env_i[3] < env_k[2]:
                            continue
                        if geom.Intersects(kg):
                            inter = geom.Intersection(kg)
                            if inter and not inter.IsEmpty():
                                f_area = inter.GetArea() / tg_area
                                if f_area > max_frac:
                                    max_frac = f_area
                if max_frac >= 0.05:
                    teve_kml_match = True
                    status_voo = "VOADO"
                    fas_com_talhao_voado.add(faz)
                    obs_voo = f"Voo Realizado ({int(round(max_frac*100))}% de cobertura KML)"

            # Caso 2: Confirmação por Base de Voo oficial (se não cruzou por KML)
            if not teve_kml_match:
                tals_voados_faz = talhoes_voados_bases.get(faz, set())
                if tal in tals_voados_faz:
                    status_voo = "VOADO"
                    fas_com_talhao_voado.add(faz)
                    obs_voo = "Voo Realizado (Confirmado na Base de Voo)"

            # Caso 3: Voo recente em Não Processados (mesmo sem talhão/KML individual gerado)
            if status_voo not in ["VOADO"]:
                if faz in pastas_nao_processados:
                    status_voo = "VOO_RECENTE"
                    fas_com_talhao_voado.add(faz)
                    obs_voo = "Voo Recente (Aguardando Processamento de Linhas/Falhas)"

            # Caso 4: Se não foi voado e nem é voo recente, checa se está no Backlog oficial (A Voar)
            if status_voo not in ["VOADO", "VOO_RECENTE"]:
                if bk:
                    dias = bk["dias_brota"]
                    if dias > 90 or "atrasado" in bk["status_prazo"].lower():
                        status_voo = "CRITICO"
                        obs_voo = f"Crítico (> 90 dias: {dias}d)"
                    elif dias >= 60:
                        status_voo = "ATENCAO"
                        obs_voo = f"Atenção (60 a 90 dias: {dias}d)"
                    else:
                        status_voo = "NOPRAZO"
                        obs_voo = f"No Prazo (< 60 dias: {dias}d)"

            # Se estava no backlog, registra a chave como vinculada ao mapa
            if bk:
                matched_keys.add(k)

            # 🔄 REGRA DE RE-VOO / SEM PORTE (Prioridade Máxima):
            motivo_revoo = custom_revoo_dict.get((faz, tal)) or custom_revoo_dict.get((faz, 'TODOS'))
            if not motivo_revoo and (pastas_voo_todas.get(faz, {}).get('revoo') or kmls_por_fa.get(faz, {}).get('revoo')):
                if status_voo not in ["VOADO"]:
                    motivo_revoo = "Pendente de Voo / Terminar Outras Partes"

            if motivo_revoo:
                status_voo = "REVOO"
                obs_voo = f"🔁 Voar Novamente ({motivo_revoo})"
                if faz in fas_com_talhao_voado and (faz, 'TODOS') in custom_revoo_dict:
                    fas_com_talhao_voado.discard(faz)
                
            geom_simpl = geom.Simplify(0.00008) # ~8m de precisao para visual ultraleve e responsivo
            geojson_geom = json.loads(geom_simpl.ExportToJson())
            geojson_geom['coordinates'] = round_coords(geojson_geom['coordinates'], 5)
            
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
            
        # Inserção direta de KMLs para fazendas que NÃO constam no SHP cadastral ou cujo KML não intersectou o SHP
        for fa_fora, info_fora in kmls_por_fa.items():
            if fa_fora not in fas_no_shp or fa_fora not in fas_com_talhao_voado:
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
                            is_prop = fa_fora.startswith('20') and len(fa_fora) == 6 or fa_fora in ['200', '106', '240', '285']
                            tipo_f = "Própria" if is_prop else "Terceiro"
                            print(f"Inserindo {lyr_direct.GetFeatureCount()} talhões diretos do KML: {faz_nome_limpo} (FA {fa_fora})")
                            for idx_d, feat_d in enumerate(lyr_direct):
                                g_d = feat_d.GetGeometryRef()
                                if not g_d: continue
                                env_d = g_d.GetEnvelope()
                                lat_mid_d = (env_d[2] + env_d[3]) / 2.0
                                area_ha_d = round(g_d.GetArea() * (111319.0 ** 2) * math.cos(math.radians(lat_mid_d)) / 10000.0, 2)
                                nome_tal_d = feat_d.GetField("Name") or f"{idx_d + 1}"
                                g_simpl_d = g_d.Simplify(0.00008)
                                gd_d = json.loads(g_simpl_d.ExportToJson())
                                gd_d['coordinates'] = round_coords(gd_d['coordinates'], 5)
                                
                                st_direct = "VOO_RECENTE" if (fa_fora in pastas_nao_processados and fa_fora not in pastas_falhas_linhas) else "VOADO"
                                obs_direct = f"Voo Recente ({os.path.basename(kml_f)})" if st_direct == "VOO_RECENTE" else f"Voo Realizado ({os.path.basename(kml_f)})"
                                m_rev_d = custom_revoo_dict.get((fa_fora, str(nome_tal_d).lstrip('0'))) or custom_revoo_dict.get((fa_fora, 'TODOS'))
                                if m_rev_d:
                                    st_direct = "REVOO"
                                    obs_direct = f"🔁 Voar Novamente ({m_rev_d})"
                                features_geojson.append({
                                    "type": "Feature",
                                    "properties": {
                                        "fa": fa_fora,
                                        "fazenda": faz_nome_limpo,
                                        "talhao": str(nome_tal_d),
                                        "area_ha": area_ha_d,
                                        "status": st_direct,
                                        "tipo": tipo_f,
                                        "dias_brota": 0,
                                        "porte_cm": "",
                                        "variedade": "",
                                        "data_plantio": "",
                                        "obs": obs_direct
                                    },
                                    "geometry": gd_d
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
    
    # Identifica a pasta local do OneDrive da Clealco (Documents/Mapas)
    onedrive_base = os.environ.get('OneDriveCommercial') or os.environ.get('OneDrive')
    destino_onedrive = os.path.join(onedrive_base, 'Mapas') if onedrive_base else None
    
    novos_copiados = 0
    atualizados_copiados = 0
    onedrive_copiados = 0
    
    # Sincronização 100% segura (Apenas LEITURA de F:)
    if os.path.exists(origem):
        try:
            for root, dirs, files in os.walk(origem):
                rel_path = os.path.relpath(root, origem)
                target_dir = os.path.join(destino, rel_path)
                if not os.path.exists(target_dir):
                    os.makedirs(target_dir, exist_ok=True)
                    
                target_onedrive_dir = os.path.join(destino_onedrive, rel_path) if destino_onedrive else None
                if target_onedrive_dir and not os.path.exists(target_onedrive_dir):
                    os.makedirs(target_onedrive_dir, exist_ok=True)
                    
                for f in files:
                    if not f.lower().endswith('.pdf'): continue
                    src_f = os.path.join(root, f)
                    dst_f = os.path.join(target_dir, f)
                    src_stat = os.stat(src_f)
                    
                    # 1. Copia para mapas_safra local (cache do site)
                    if not os.path.exists(dst_f):
                        shutil.copy2(src_f, dst_f)
                        novos_copiados += 1
                    else:
                        dst_stat = os.stat(dst_f)
                        if src_stat.st_mtime > dst_stat.st_mtime + 2 or src_stat.st_size != dst_stat.st_size:
                            shutil.copy2(src_f, dst_f)
                            atualizados_copiados += 1
                            
                    # 2. Copia para o OneDrive local (para envio automático à nuvem da Clealco)
                    if target_onedrive_dir:
                        dst_od_f = os.path.join(target_onedrive_dir, f)
                        if not os.path.exists(dst_od_f):
                            shutil.copy2(src_f, dst_od_f)
                            onedrive_copiados += 1
                        else:
                            od_stat = os.stat(dst_od_f)
                            if src_stat.st_mtime > od_stat.st_mtime + 2 or src_stat.st_size != od_stat.st_size:
                                shutil.copy2(src_f, dst_od_f)
                                onedrive_copiados += 1
                                
            if novos_copiados > 0 or atualizados_copiados > 0 or onedrive_copiados > 0:
                print(f"[MAPAS] Sincronização: {novos_copiados} novos mapas, {atualizados_copiados} mapas atualizados de F: ({onedrive_copiados} sincronizados no OneDrive).")
            else:
                print("[MAPAS] Todos os mapas já estão 100% sincronizados com F: e OneDrive.")
        except Exception as ex_sync:
            print(f"[MAPAS] Aviso durante sincronização de F:: {ex_sync}")
    else:
        print("[MAPAS] Disco F: não acessível no momento. Utilizando mapas locais da pasta mapas_safra.")
        # Se F: não estiver acessível, garante que novos mapas em mapas_safra também vão para o OneDrive
        if destino_onedrive and os.path.exists(destino):
            try:
                for root, dirs, files in os.walk(destino):
                    rel_path = os.path.relpath(root, destino)
                    target_od = os.path.join(destino_onedrive, rel_path)
                    if not os.path.exists(target_od):
                        os.makedirs(target_od, exist_ok=True)
                    for f in files:
                        if not f.lower().endswith('.pdf'): continue
                        src_f = os.path.join(root, f)
                        dst_f = os.path.join(target_od, f)
                        if not os.path.exists(dst_f):
                            shutil.copy2(src_f, dst_f)
                            onedrive_copiados += 1
                        else:
                            src_st = os.stat(src_f)
                            dst_st = os.stat(dst_f)
                            if src_st.st_mtime > dst_st.st_mtime + 2 or src_st.st_size != dst_st.st_size:
                                shutil.copy2(src_f, dst_f)
                                onedrive_copiados += 1
                if onedrive_copiados > 0:
                    print(f"[MAPAS] {onedrive_copiados} mapas locais sincronizados para a pasta do OneDrive.")
            except Exception as ex_od:
                print(f"[MAPAS] Aviso durante sincronização local com OneDrive: {ex_od}")
        
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
                
                # URL oficial direta do arquivo PDF no visualizador do SharePoint/OneDrive (:b:/r/ com web=1 abre o PDF)
                encoded_sp_path = urllib.parse.quote(sp_path)
                url_sharepoint = f"https://clealco-my.sharepoint.com/:b:/r/personal/fabiano_castro_clealco_com_br/Documents/Mapas/{encoded_sp_path}?csf=1&web=1"
                
                # URL oficial da pasta da categoria no SharePoint/OneDrive (:f:/r/ com web=1)
                encoded_cat_path = urllib.parse.quote(categoria)
                url_pasta_sharepoint = f"https://clealco-my.sharepoint.com/:f:/r/personal/fabiano_castro_clealco_com_br/Documents/Mapas/{encoded_cat_path}?csf=1&web=1"

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

def compilar_historico_falhas():
    print("[HISTORICO FALHAS] Iniciando compilação 2021, 2022, 2023, 2024, 2025, 2026...")
    import csv, glob
    
    # 1. Mapeamento de nomes e zonas de fazendas do SHP e Excel
    info_por_fa = {}
    shp_path = r"W:\11-SHAPES TEMATICOS\09-SHP ATUALIZADO\SHP_CLEALCO.shp"
    if os.path.exists(shp_path):
        try:
            from osgeo import ogr
            ds = ogr.Open(shp_path)
            if ds:
                layer = ds.GetLayer()
                for feat in layer:
                    faz = str(feat.GetField('FAZENDA') or '').strip().lstrip('0')
                    nome = str(feat.GetField('NOME_FAZ') or '').strip().upper()
                    zona = str(feat.GetField('ZONA') or '').strip()
                    if faz and faz not in info_por_fa:
                        info_por_fa[faz] = {'nome': nome, 'zona': zona}
        except Exception as e:
            print(f"[HISTORICO FALHAS] Aviso ao ler SHP (osgeo/ogr opcional): {e}")

    # Complementar com dados.xlsx
    excel_path = r"W:\00-SITE QUALIDADE APLICACAO\Voo_e_processamentosparagithub_2.0\dados.xlsx"
    if os.path.exists(excel_path):
        try:
            wb_temp = openpyxl.load_workbook(excel_path, data_only=True)
            for sname in ['BASE_PROCESSAMENTOS', 'FALHAS_PLANTIO']:
                if sname in wb_temp.sheetnames:
                    s = wb_temp[sname]
                    for row in s.iter_rows(values_only=True):
                        fa = str(row[0] or '').strip().lstrip('0')
                        nome = str(row[1] or '').strip()
                        if fa and nome and nome != 'NÃO INFORMADO':
                            if fa not in info_por_fa:
                                info_por_fa[fa] = {'nome': nome, 'zona': ''}
                            elif not info_por_fa[fa]['nome']:
                                info_por_fa[fa]['nome'] = nome
        except Exception as e:
            print(f"[HISTORICO FALHAS] Aviso ao ler dados.xlsx: {e}")

    base_dirs = {
        '2021': r"W:\01-DRONE\08-BANCO DE DADOS\2021\Falhas",
        '2022': r"W:\01-DRONE\08-BANCO DE DADOS\2022\Falhas",
        '2023': r"W:\01-DRONE\08-BANCO DE DADOS\2023\Falhas",
        '2024': r"W:\01-DRONE\08-BANCO DE DADOS\2024\Falhas",
        '2025': r"W:\01-DRONE\08-BANCO DE DADOS\2025\Falhas",
        '2026': r"W:\01-DRONE\08-BANCO DE DADOS\2026\Falhas"
    }

    def classificar_status(pct, tipo):
        if tipo == 'Própria':
            if pct <= 8.0: return 'Aceitável'
            elif pct <= 10.0: return 'Atenção'
            else: return 'Crítico'
        else:
            if pct <= 5.0: return 'Aceitável'
            elif pct <= 10.0: return 'Atenção'
            else: return 'Crítico'

    def pnum(s):
        try: return float(str(s).strip().replace(',', '.'))
        except: return 0.0

    fazendas_todas = []
    resumo_anos = {}

    for ano, pasta in base_dirs.items():
        if not os.path.exists(pasta):
            print(f"[HISTORICO FALHAS] Pasta não encontrada: {pasta}")
            continue
        
        csvs = glob.glob(os.path.join(pasta, "*.csv"))
        ha_ano = 0.0
        km_l_ano = 0.0
        km_f_ano = 0.0
        num_f_ano = 0
        status_contagem = {'Aceitável': 0, 'Atenção': 0, 'Crítico': 0}

        for c in csvs:
            basename = os.path.splitext(os.path.basename(c))[0]
            m_fa_h = re.match(r'^(\d+)', basename)
            fa = m_fa_h.group(1) if m_fa_h else basename.split()[0].strip()
            fa_clean = fa.lstrip('0')
            
            # Tipo de parceria
            is_propria = fa.startswith('20') and len(fa) == 6
            tipo = 'Própria' if is_propria else 'Terceiro'

            info = info_por_fa.get(fa, info_por_fa.get(fa_clean, {}))
            nome_faz = info.get('nome', '')
            zona = info.get('zona', '')
            if not nome_faz:
                partes = basename.split(maxsplit=1)
                nome_faz = partes[1].strip() if len(partes) > 1 else f"FAZENDA FA {fa}"

            # Ler CSV com delimitador vírgula
            rows = []
            try:
                with open(c, 'r', encoding='latin1') as f:
                    rows = list(csv.reader(f, delimiter=','))
            except Exception:
                try:
                    with open(c, 'r', encoding='utf-8') as f:
                        rows = list(csv.reader(f, delimiter=','))
                except Exception as e:
                    print(f"Erro lendo {c}: {e}")
                    continue

            if len(rows) < 2:
                continue

            header = [h.strip() for h in rows[0]]
            if 'Porc_Falha' not in header:
                continue

            col_ha = header.index('A_Liq') if 'A_Liq' in header else 4
            col_pct = header.index('Porc_Falha') if 'Porc_Falha' in header else 5
            col_nf = header.index('Num_Falhas') if 'Num_Falhas' in header else 6
            col_cf = header.index('Comp_Falhas') if 'Comp_Falhas' in header else 7
            col_cl = header.index('Comp_Linhas') if 'Comp_Linhas' in header else 8
            col_nl = header.index('Num_Linhas') if 'Num_Linhas' in header else 9

            summary_row = None
            talhoes_rows = []
            for r in rows[1:]:
                if not r or len(r) <= max(col_ha, col_pct, col_nf, col_cf, col_cl, col_nl): continue
                prog_id = r[0].strip()
                if prog_id == '0':
                    summary_row = r
                else:
                    try:
                        if float(prog_id) > 0:
                            talhoes_rows.append(r)
                    except:
                        pass

            if not summary_row and talhoes_rows:
                summary_row = talhoes_rows[0]

            if not summary_row:
                continue

            ha = pnum(summary_row[col_ha])
            pct = pnum(summary_row[col_pct])
            nf = int(pnum(summary_row[col_nf]))
            cf = pnum(summary_row[col_cf]) / 1000.0
            cl = pnum(summary_row[col_cl]) / 1000.0
            nl = int(pnum(summary_row[col_nl]))

            st = classificar_status(pct, tipo)
            status_contagem[st] += 1

            ha_ano += ha
            km_l_ano += cl
            km_f_ano += cf
            num_f_ano += nf

            # Talhões detalhados
            lista_t = []
            for tr in talhoes_rows:
                user_id = tr[1].strip() or tr[0].strip()
                t_ha = pnum(tr[col_ha])
                t_pct = pnum(tr[col_pct])
                t_nf = int(pnum(tr[col_nf]))
                t_cf = pnum(tr[col_cf]) / 1000.0
                t_cl = pnum(tr[col_cl]) / 1000.0
                t_nl = int(pnum(tr[col_nl]))
                t_st = classificar_status(t_pct, tipo)
                lista_t.append({
                    't': user_id,
                    'ha': round(t_ha, 2),
                    'cl': round(t_cl, 2),
                    'cf': round(t_cf, 2),
                    'pct': round(t_pct, 2),
                    'nf': t_nf,
                    'nl': t_nl,
                    'st': t_st
                })

            fazendas_todas.append({
                'ano': ano,
                'fa': fa,
                'faz': nome_faz,
                'tipo': tipo,
                'zona': zona,
                'ha': round(ha, 2),
                'pct': round(pct, 2),
                'cf': round(cf, 2),
                'cl': round(cl, 2),
                'nf': nf,
                'nl': nl,
                'st': st,
                'qtd_t': len(lista_t),
                'talhoes': lista_t
            })

        media_pond = round((km_f_ano / km_l_ano * 100), 2) if km_l_ano > 0 else 0.0
        resumo_anos[ano] = {
            'ha_total': round(ha_ano, 2),
            'km_linhas': round(km_l_ano, 2),
            'km_falhas': round(km_f_ano, 2),
            'num_falhas': num_f_ano,
            'media_ponderada_pct': media_pond,
            'qtd_fazendas': len([f for f in fazendas_todas if f['ano'] == ano]),
            'status_contagem': status_contagem
        }
        print(f"[HISTORICO FALHAS] Safra {ano}: {resumo_anos[ano]['qtd_fazendas']} fazendas, {ha_ano:.1f} ha, Média Ponderada: {media_pond}%")

    # Mapear histórico multi-safra para fazendas repetidas
    fa_map = {}
    for item in fazendas_todas:
        fa = item['fa']
        if fa not in fa_map:
            fa_map[fa] = []
        fa_map[fa].append({
            'ano': item['ano'],
            'pct': item['pct'],
            'ha': item['ha'],
            'st': item['st']
        })

    for item in fazendas_todas:
        item['hist_anos'] = sorted(fa_map.get(item['fa'], []), key=lambda x: x['ano'])

    # Ordenar por ano desc, depois por pct desc
    fazendas_todas.sort(key=lambda x: (x['ano'], x['pct']), reverse=True)

    payload = {
        'resumo_anos': resumo_anos,
        'fazendas': fazendas_todas,
        'data_atualizacao': datetime.now().strftime('%d/%m/%Y %H:%M')
    }

    payload_str = json.dumps(payload, ensure_ascii=False)
    encrypted = criptografar_payload(payload_str)

    output_path = r"W:\00-SITE QUALIDADE APLICACAO\Voo_e_processamentosparagithub_2.0\dados_historico_falhas.js"
    with open(output_path, "w", encoding="utf-8") as f_out:
        f_out.write("// Histórico Consolidado de Falhas de Plantio (2021 - 2026)\n")
        f_out.write(f"// Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f_out.write("// Payload criptografado AES-GCM (Clealco)\n")
        f_out.write("window.DADOS_HISTORICO_FALHAS = ")
        json.dump(encrypted, f_out, ensure_ascii=False)
        f_out.write(";\n")

    print(f"[HISTORICO FALHAS] Arquivo gravado com sucesso: {output_path} ({len(fazendas_todas)} registros compilados).")

compilar_historico_falhas()

def incrementar_versao_service_worker():
    sw_path = r"W:\00-SITE QUALIDADE APLICACAO\Voo_e_processamentosparagithub_2.0\service-worker.js"
    if not os.path.exists(sw_path):
        return
    try:
        with open(sw_path, 'r', encoding='utf-8') as f:
            content = f.read()
        import re
        match = re.search(r"const CACHE_NAME = 'clealco-voos-v(\d+)';", content)
        if match:
            v_atual = int(match.group(1))
            novo_v = v_atual + 1
            novo_content = re.sub(r"const CACHE_NAME = 'clealco-voos-v\d+';", f"const CACHE_NAME = 'clealco-voos-v{novo_v}';", content)
            with open(sw_path, 'w', encoding='utf-8') as f:
                f.write(novo_content)
            print(f"[SERVICE WORKER] Cache PWA atualizado com sucesso para: clealco-voos-v{novo_v}")
        else:
            print("[SERVICE WORKER] Aviso: Padrão de CACHE_NAME não encontrado no service-worker.js.")
    except Exception as e:
        print(f"[SERVICE WORKER] Erro ao incrementar versão do cache: {e}")

incrementar_versao_service_worker()



