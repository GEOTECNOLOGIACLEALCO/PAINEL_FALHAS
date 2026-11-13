import json, base64, os, sys
from datetime import datetime
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes

caminho_js = os.path.join(os.path.dirname(__file__), 'dados_injetados.js')
if not os.path.exists(caminho_js):
    sys.exit(0)

try:
    with open(caminho_js, 'r', encoding='utf-8') as f:
        content = f.read()
    payload = json.loads(content.split('window.DADOS_CLEALCO_EMBARCADOS = ')[1].strip().rstrip(';'))
    salt = base64.b64decode(payload['s'])
    iv = base64.b64decode(payload['i'])
    ct = base64.b64decode(payload['d'])
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=100000)
    key = kdf.derive(b'Clealco@Geo2026')
    aesgcm = AESGCM(key)
    dec = json.loads(aesgcm.decrypt(iv, ct, None).decode('utf-8'))

    falhas = dec.get('falhas_plantio', [])
    prop = sum(float(f.get('Area_Ha') or 0) for f in falhas if str(f.get('Tipo')).lower().startswith('pr') or (str(f.get('FA')).startswith('20') and len(str(f.get('FA')))==6))
    terc = sum(float(f.get('Area_Ha') or 0) for f in falhas if not (str(f.get('Tipo')).lower().startswith('pr') or (str(f.get('FA')).startswith('20') and len(str(f.get('FA')))==6)))
    total = prop + terc

    meta_prop = 3374.33
    meta_terc = 11000.00
    meta_total = meta_prop + meta_terc

    pct_prop = (prop / meta_prop) * 100 if meta_prop else 0
    pct_terc = (terc / meta_terc) * 100 if meta_terc else 0
    pct_total = (total / meta_total) * 100 if meta_total else 0

    falta_prop = max(0.0, meta_prop - prop)
    falta_terc = max(0.0, meta_terc - terc)
    falta_total = max(0.0, meta_total - total)

    agora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    print('\n' + '=' * 82)
    print('   CLEALCO GEOTECNOLOGIA - RESUMO EXECUTIVO (PAINEL PWA)')
    print('=' * 82)
    print(f'   [+] REALIZADO PROPRIAS:   {prop:10,.2f} ha  ({pct_prop:5.1f}% | Meta: {meta_prop:,.2f} ha | Falta: {falta_prop:,.2f} ha)'.replace(',', 'X').replace('.', ',').replace('X', '.'))
    print(f'   [+] REALIZADO TERCEIROS:  {terc:10,.2f} ha  ({pct_terc:5.1f}% | Meta: {meta_terc:,.2f} ha | Falta: {falta_terc:,.2f} ha)'.replace(',', 'X').replace('.', ',').replace('X', '.'))
    print(f'   [+] CONSOLIDADO GERAL:    {total:10,.2f} ha  ({pct_total:5.1f}% | Meta: {meta_total:,.2f} ha | Falta: {falta_total:,.2f} ha)'.replace(',', 'X').replace('.', ',').replace('X', '.'))
    print('-' * 82)
    print(f'   [+] Atualizacao: {agora}  | Total Fazendas: {len(falhas)}')
    print('   [+] Diario Auditoria: W:\\15-SETOR\\FABIANO\\DIARIO_DE_BORDO_AUDITORIA.txt')
    print('=' * 82 + '\n')
except Exception as e:
    pass

