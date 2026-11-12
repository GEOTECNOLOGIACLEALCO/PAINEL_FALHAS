#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Clealco Geotecnologia - Pipeline & Script Orchestrator
------------------------------------------------------
Utilitário centralizado para o Analista e a IA interagirem
com segurança com os scripts principais do ecossistema:
  - W:\\15-SETOR\\FABIANO\\Voo_e_processamentos_2.0\\2.0importos.py
  - W:\\15-SETOR\\FABIANO\\Voo_e_processamentosparagithub_2.0\\atualizar_dados.py
  - W:\\15-SETOR\\FABIANO\\Voo_e_processamentosparagithub_2.0\\atualizar_dados.bat

Uso:
  python pipeline_clealco.py status
  python pipeline_clealco.py set-creditos <novo_valor> [--compile]
  python pipeline_clealco.py decrypt [injetados|mapa]
  python pipeline_clealco.py compile
  python pipeline_clealco.py sync
  python pipeline_clealco.py diagnose
"""

import os
import sys

try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except:
    pass

import json
import base64
import argparse
import subprocess
from datetime import datetime

# Caminhos Mestres
PASTA_GITHUB = r"W:\15-SETOR\FABIANO\Voo_e_processamentosparagithub_2.0"
PASTA_BACKEND = r"W:\15-SETOR\FABIANO\Voo_e_processamentos_2.0"
PASTA_NAO_PROCESSADOS = r"W:\01-DRONE\01-IMAGENS VOO\2026\1 - Não Processados"
PASTA_AUDITORIA_DOCS = os.path.join(PASTA_NAO_PROCESSADOS, "AUDITORIA - DOC - SISTEMA")

ARQ_CREDITOS_TXT = (
    os.path.join(PASTA_AUDITORIA_DOCS, "creditos_inforow.txt")
    if os.path.exists(os.path.join(PASTA_AUDITORIA_DOCS, "creditos_inforow.txt"))
    else os.path.join(PASTA_NAO_PROCESSADOS, "creditos_inforow.txt")
)
ARQ_BASE_EXCEL = (
    os.path.join(PASTA_AUDITORIA_DOCS, "BASE DE PROCESSAMENTOS.xlsx")
    if os.path.exists(os.path.join(PASTA_AUDITORIA_DOCS, "BASE DE PROCESSAMENTOS.xlsx"))
    else os.path.join(PASTA_NAO_PROCESSADOS, "BASE DE PROCESSAMENTOS.xlsx")
)
ARQ_DADOS_XLSX = os.path.join(PASTA_GITHUB, "dados.xlsx")
ARQ_DADOS_INJETADOS = os.path.join(PASTA_GITHUB, "dados_injetados.js")
ARQ_DADOS_MAPA = os.path.join(PASTA_GITHUB, "dados_mapa_voos.js")
ARQ_INDEX_HTML = os.path.join(PASTA_GITHUB, "index.html")
ARQ_SW = os.path.join(PASTA_GITHUB, "service-worker.js")
ARQ_ATUALIZAR_PY = os.path.join(PASTA_GITHUB, "atualizar_dados.py")
ARQ_ATUALIZAR_BAT = os.path.join(PASTA_GITHUB, "atualizar_dados.bat")

QGIS_PY = r"C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe"
CHAVE_ACESSO = "Clealco@Geo2026"

def get_python_exe():
    if os.path.exists(QGIS_PY):
        return QGIS_PY
    return sys.executable

def decrypt_payload(c_dict, password=CHAVE_ACESSO):
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
        from cryptography.hazmat.primitives import hashes
    except ImportError:
        print("[ERRO] Biblioteca cryptography não encontrada no Python atual. Use QGIS Python.")
        return None

    salt = base64.b64decode(c_dict["s"])
    iv = base64.b64decode(c_dict["i"])
    ciphertext = base64.b64decode(c_dict["d"])

    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
    )
    key = kdf.derive(password.encode("utf-8"))
    aesgcm = AESGCM(key)
    decrypted_bytes = aesgcm.decrypt(iv, ciphertext, None)
    return json.loads(decrypted_bytes.decode("utf-8"))

def obter_creditos_txt():
    if os.path.exists(ARQ_CREDITOS_TXT):
        try:
            with open(ARQ_CREDITOS_TXT, "r", encoding="utf-8") as f:
                return float(f.read().strip())
        except Exception as e:
            return f"Erro: {e}"
    return "NÃO ENCONTRADO"

def obter_creditos_xlsx():
    if not os.path.exists(ARQ_DADOS_XLSX):
        return "NÃO ENCONTRADO"
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ARQ_DADOS_XLSX, data_only=True)
        if "CONFIG" in wb.sheetnames:
            sheet = wb["CONFIG"]
            return float(sheet.cell(row=2, column=1).value or 0.0)
        return "ABA CONFIG AUSENTE"
    except Exception as e:
        return f"Erro: {e}"

def extrair_json_de_js(caminho_arquivo):
    with open(caminho_arquivo, "r", encoding="utf-8") as f:
        txt = f.read().strip()
    idx = txt.find("=")
    if idx == -1:
        return None
    json_str = txt[idx+1:].strip()
    if json_str.endswith(";"):
        json_str = json_str[:-1].strip()
    return json.loads(json_str)

def obter_creditos_injetados():
    if not os.path.exists(ARQ_DADOS_INJETADOS):
        return "NÃO ENCONTRADO", "N/D"
    try:
        payload = extrair_json_de_js(ARQ_DADOS_INJETADOS)
        if not payload:
            return "PAYLOAD NÃO ENCONTRADO", "N/D"
        dec = decrypt_payload(payload)
        if dec and "creditos_inforow" in dec:
            return float(dec["creditos_inforow"]), dec.get("data_compilacao", "N/D")
        return "ERRO DE DESCRIPTOGRAFIA", "N/D"
    except Exception as e:
        return f"Erro: {e}", "N/D"

def obter_versao_sw():
    if not os.path.exists(ARQ_SW):
        return "N/D"
    try:
        with open(ARQ_SW, "r", encoding="utf-8") as f:
            txt = f.read()
        import re
        m = re.search(r"CACHE_NAME\s*=\s*['\"]([^'\"]+)['\"]", txt)
        return m.group(1) if m else "N/D"
    except:
        return "N/D"

def cmd_status(args):
    print("=" * 70)
    print("         CLEALCO GEOTECNOLOGIA - STATUS DO ECOSSISTEMA")
    print("=" * 70)
    print(f"Data/Hora Local: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"Interpretador QGIS Python: {'OK' if os.path.exists(QGIS_PY) else 'AUSENTE'} ({QGIS_PY})")
    print(f"Versão Cache PWA (Service Worker): {obter_versao_sw()}")
    print("-" * 70)

    # 1. Checagem de Créditos Inforow (Tríplice Checagem)
    c_txt = obter_creditos_txt()
    c_xlsx = obter_creditos_xlsx()
    c_inj, dt_comp = obter_creditos_injetados()

    print("SALDO DE CRÉDITOS INFOROW (Tríplice Fonte):")
    print(f"  [1] creditos_inforow.txt (Fonte Raiz Backend):    {c_txt}")
    print(f"  [2] dados.xlsx (CONFIG A2 - Entrada Web):         {c_xlsx}")
    print(f"  [3] dados_injetados.js (Criptografado Publicado): {c_inj} (Compilado: {dt_comp})")

    # Alerta de divergência
    vals = [v for v in [c_txt, c_xlsx, c_inj] if isinstance(v, (int, float))]
    if len(vals) == 3 and not (vals[0] == vals[1] == vals[2]):
        print("  [!] ALERTA: HÁ DIVERGÊNCIA ENTRE AS FONTES DE CRÉDITOS!")
        print("      Execute: python pipeline_clealco.py set-creditos <valor> --compile")
    elif len(vals) == 3:
        print("  [OK] SINCRONIA PERFEITA: Todas as fontes de créditos estão idênticas!")

    print("-" * 70)
    print("VERIFICAÇÃO DE ARQUIVOS E CAMINHOS:")
    paths = [
        ("Backend 2.0importos.py", os.path.join(PASTA_BACKEND, "2.0importos.py")),
        ("BASE DE PROCESSAMENTOS.xlsx", ARQ_BASE_EXCEL),
        ("Publicador atualizar_dados.py", ARQ_ATUALIZAR_PY),
        ("Publicador atualizar_dados.bat", ARQ_ATUALIZAR_BAT),
        ("Frontend index.html", ARQ_INDEX_HTML),
        ("Mapa Vetorial dados_mapa_voos.js", ARQ_DADOS_MAPA),
    ]
    for label, p in paths:
        status_arq = "OK" if os.path.exists(p) else "FALTANDO"
        tam = f"({os.path.getsize(p) / 1024:.1f} KB)" if os.path.exists(p) else ""
        print(f"  {label:<32}: [{status_arq}] {tam}")

    print("-" * 70)
    # Git Status
    try:
        git_branch = subprocess.check_output(["git", "branch", "--show-current"], cwd=PASTA_GITHUB, text=True).strip()
        git_log = subprocess.check_output(["git", "log", "-n", "1", "--oneline"], cwd=PASTA_GITHUB, text=True).strip()
        print(f"GIT: Branch '{git_branch}' | Último Commit: {git_log}")
    except Exception as e:
        print(f"GIT: Não foi possível obter status ({e})")
    print("=" * 70)

def cmd_set_creditos(args):
    novo_valor = float(args.valor)
    print(f"[*] Atualizando Créditos Inforow para: {novo_valor:.2f}")

    # 1. Salvar no arquivo raiz da rede
    try:
        with open(ARQ_CREDITOS_TXT, "w", encoding="utf-8") as f:
            f.write(f"{novo_valor}\n")
        print(f"  [1/3] [OK] Gravado com sucesso em: {ARQ_CREDITOS_TXT}")
    except Exception as e:
        print(f"  [1/3] [ERRO] Erro ao gravar no TXT raiz: {e}")
        return

    # 2. Atualizar dados.xlsx
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ARQ_DADOS_XLSX)
        if "CONFIG" not in wb.sheetnames:
            wb.create_sheet("CONFIG")
        sheet = wb["CONFIG"]
        sheet.cell(row=1, column=1).value = "creditos_inforow"
        sheet.cell(row=2, column=1).value = novo_valor
        wb.save(ARQ_DADOS_XLSX)
        print(f"  [2/3] [OK] Gravado com sucesso em: {ARQ_DADOS_XLSX} (Aba CONFIG)")
    except Exception as e:
        print(f"  [2/3] [ERRO] Erro ao atualizar dados.xlsx: {e}")

    # 3. Atualizar fallbacks em index.html
    try:
        with open(ARQ_INDEX_HTML, "r", encoding="utf-8") as f:
            html = f.read()
        import re
        fmt_val = f"{novo_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        html = re.sub(r'id="kpi-creditos">[^<]+<', f'id="kpi-creditos">{fmt_val} cr.<', html)
        html = re.sub(r'let creditosInforow\s*=\s*[0-9\.]+;', f'let creditosInforow = {novo_valor};', html)
        html = re.sub(r'creditos_inforow\s*\|\|\s*[0-9\.]+;', f'creditos_inforow || {novo_valor};', html)
        with open(ARQ_INDEX_HTML, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  [3/3] [OK] Fallbacks atualizados em index.html ({fmt_val} cr.)")
    except Exception as e:
        print(f"  [3/3] [AVISO] Aviso ao atualizar index.html: {e}")

    if args.compile:
        print("\n[*] Recompilando dados_injetados.js via QGIS Python...")
        cmd_compile(args)

def cmd_decrypt(args):
    tipo = args.tipo.lower()
    alvo = ARQ_DADOS_INJETADOS if tipo == "injetados" else ARQ_DADOS_MAPA

    print(f"[*] Descriptografando {alvo}...")
    if not os.path.exists(alvo):
        print(f"[ERRO] Arquivo não encontrado: {alvo}")
        return

    payload = extrair_json_de_js(alvo)
    if not payload:
        print(f"[ERRO] Não foi possível extrair o payload JSON de {alvo}")
        return

    dados = decrypt_payload(payload)
    if not dados:
        return

    print("[OK] Descriptografia realizada com sucesso!")
    if tipo == "injetados":
        print(f"  - Data Compilação: {dados.get('data_compilacao')}")
        print(f"  - Créditos Inforow: {dados.get('creditos_inforow')}")
        print(f"  - Total de Voos: {len(dados.get('base_processamentos', []))}")
        print(f"  - Backlog FAs: {len(dados.get('backlog', []))}")
        print(f"  - Registros Falhas: {len(dados.get('falhas_plantio', []))}")
    else:
        features = dados.get("features", [])
        print(f"  - Total de Polígonos de Talhões (KML/SHP): {len(features)}")
        if features:
            props = features[0].get("properties", {})
            print(f"  - Exemplo Propriedades do 1º Talhão: {list(props.keys())}")
            print(f"  - Exemplo FA: {props.get('fa')} | Fazenda: {props.get('fazenda')} | Talhão: {props.get('talhao')}")

    if args.export:
        out_file = f"debug_{tipo}.json"
        with open(out_file, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
        print(f"  [DUMP] JSON exportado para: {os.path.abspath(out_file)}")

def cmd_compile(args):
    py_exe = get_python_exe()
    print(f"[*] Executando atualizar_dados.py com: {py_exe}")
    res = subprocess.run([py_exe, "atualizar_dados.py"], cwd=PASTA_GITHUB)
    if res.returncode == 0:
        print("[OK] Compilação concluída com sucesso!")
    else:
        print(f"[ERRO] Falha durante compilação (Código {res.returncode})")

def cmd_sync(args):
    print(f"[*] Disparando atualizar_dados.bat...")
    res = subprocess.run(["cmd.exe", "/c", ARQ_ATUALIZAR_BAT], cwd=PASTA_GITHUB)
    if res.returncode == 0:
        print("[OK] Sincronização e envio para GitHub concluídos com sucesso!")
    else:
        print(f"[ERRO] Falha na sincronização (Código {res.returncode})")

def cmd_diagnose(args):
    print("=" * 70)
    print("         CLEALCO GEOTECNOLOGIA - DIAGNÓSTICO DO AMBIENTE")
    print("=" * 70)
    print(f"Python Executável Atual: {sys.executable} (v{sys.version.split()[0]})")
    print(f"Python QGIS 3.12: {QGIS_PY} -> {'PRESENTE' if os.path.exists(QGIS_PY) else 'AUSENTE'}")

    libs = ["openpyxl", "cryptography", "osgeo", "geopandas", "PIL"]
    print("\nVerificação de Bibliotecas no Interpretador Atual:")
    for lib in libs:
        try:
            __import__(lib)
            print(f"  - {lib:<15}: [OK] Disponível")
        except ImportError:
            print(f"  - {lib:<15}: [FALTA] Ausente")

    print("\nTeste de Acesso a Pastas de Rede:")
    pastas = [
        ("W: 1 - Não Processados", PASTA_NAO_PROCESSADOS),
        ("W: Voo_e_processamentos_2.0", PASTA_BACKEND),
        ("W: Voo_e_processamentosparagithub_2.0", PASTA_GITHUB),
        ("M: Estudo de Arrendamentos", r"M:\17-Estudo de Arrendamentos"),
    ]
    for lbl, p in pastas:
        acesso = os.path.exists(p)
        print(f"  - {lbl:<35}: {'[OK] Acessível' if acesso else '[ERRO] Inacessível'}")
    print("=" * 70)

def main():
    parser = argparse.ArgumentParser(description="Clealco Geotecnologia - Orquestrador do Pipeline")
    subparsers = parser.add_subparsers(dest="comando", help="Comandos disponíveis")

    subparsers.add_parser("status", help="Exibe o status completo de sincronização e créditos")

    p_cred = subparsers.add_parser("set-creditos", help="Ajusta o saldo de créditos Inforow de forma consistente")
    p_cred.add_argument("valor", type=float, help="Novo valor de créditos (ex: 10713.75)")
    p_cred.add_argument("--compile", action="store_true", help="Recompila automaticamente após atualizar")

    p_dec = subparsers.add_parser("decrypt", help="Descriptografa e inspeciona dados_injetados.js ou dados_mapa_voos.js")
    p_dec.add_argument("tipo", choices=["injetados", "mapa"], default="injetados", nargs="?", help="Qual arquivo descriptografar")
    p_dec.add_argument("--export", action="store_true", help="Exporta o JSON legível para arquivo de debug")

    subparsers.add_parser("compile", help="Executa atualizar_dados.py usando o Python do QGIS")
    subparsers.add_parser("sync", help="Executa atualizar_dados.bat (compilação + git commit/push)")
    subparsers.add_parser("diagnose", help="Diagnostica interpretadores e conectividade")

    args = parser.parse_args()
    if not args.comando or args.comando == "status":
        cmd_status(args)
    elif args.comando == "set-creditos":
        cmd_set_creditos(args)
    elif args.comando == "decrypt":
        cmd_decrypt(args)
    elif args.comando == "compile":
        cmd_compile(args)
    elif args.comando == "sync":
        cmd_sync(args)
    elif args.comando == "diagnose":
        cmd_diagnose(args)

if __name__ == "__main__":
    main()
