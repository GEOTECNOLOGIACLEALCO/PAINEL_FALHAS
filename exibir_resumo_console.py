import json, base64, os, sys
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

    print('\n' + '=' * 76)
    print('  CLEALCO GEOTECNOLOGIA - RESUMO EXECUTIVO ATUALIZADO (PAINEL PWA)')
    print('=' * 76)
    print(f'  [+] REALIZADO PROPRIAS:   {prop:10,.2f} ha  (Meta: 3.374,33 ha)')
    print(f'  [+] REALIZADO TERCEIROS:  {terc:10,.2f} ha  (Meta: 11.000,00 ha)')
    print(f'  [+] CONSOLIDADO GERAL:    {total:10,.2f} ha  (Meta: 14.374,33 ha)')
    print('-' * 76)
    print('  [+] Diario de Auditoria:  W:\\15-SETOR\\FABIANO\\DIARIO_DE_BORDO_AUDITORIA.txt')
    print('=' * 76 + '\n')
except Exception as e:
    pass
