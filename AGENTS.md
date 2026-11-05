# Clealco Geotecnologia - Diretrizes do Ecossistema de Voos e Processamentos

Este documento define a arquitetura, regras de ouro, fluxos de dados e comandos para agentes de IA e desenvolvedores que interagem com o sistema de monitoramento de voos de drones, falhas de plantio e publicações web da Clealco Geotecnologia.

---

## 1. Mapa de Diretórios e Papéis

### A. Backend & Ingestão Local
- **Pasta:** `W:\00-SITE QUALIDADE APLICACAO\Voo_e_processamentos_2.0`
- **Script Principal:** `2.0importos.py` (Flask + Robô Watchdog)
- **Ativador:** `INICIAR_SISTEMA.bat`
- **Funções:**
  - Monitora novos voos em `W:\01-DRONE\01-IMAGENS VOO\2026\1 - Não Processados`.
  - Verifica ortomosaicos (.tif), KMLs de talhões e arquivos de linhas/falhas.
  - Abate hectares e calcula saldo em `creditos_inforow.txt`.
  - Registra atividades em `BASE DE PROCESSAMENTOS.xlsx`.
  - Compila `dados.xlsx` para publicação e dispara `atualizar_dados.bat`.

### B. Publicador Web & Criptografia (Repositório Git)
- **Pasta:** `W:\00-SITE QUALIDADE APLICACAO\Voo_e_processamentosparagithub_2.0`
- **Script Compilador:** `atualizar_dados.py`
- **Script Automatizador:** `atualizar_dados.bat` (Git Pull, Compilação QGIS, Git Commit/Push)
- **Orquestrador Auxiliar:** `pipeline_clealco.py`
- **Frontend:** `index.html`, `service-worker.js`
- **Repositório Remoto:** `https://github.com/geotecnologiaclealco/PAINEL_FALHAS.git` (branch `main`)

### C. Fontes de Entrada na Rede Corporativa
- `W:\01-DRONE\01-IMAGENS VOO\2026\1 - Não Processados\creditos_inforow.txt` -> **FONTE RAIZ DO SALDO INFOROW**.
- `W:\01-DRONE\01-IMAGENS VOO\2026\1 - Não Processados\BASE DE PROCESSAMENTOS.xlsx` -> Fila e histórico de voos.
- `M:\17-Estudo de Arrendamentos\Status de Estudos.xlsx` -> Cadastro oficial de FAs, fazendas e parceiros.
- `M:\17-Estudo de Arrendamentos\SHP_CLEALCO.shp` -> Base cartográfica de talhões.
- `W:\01-DRONE\08-BANCO DE DADOS\2026\Falhas` -> Bases consolidadas de falhas.

---

## 2. As Quatro Regras de Ouro (Anti-Perda)

### 🥇 Regra 1: Saldo de Créditos Inforow
- **NUNCA** edite apenas `dados.xlsx` ou `index.html`. O script `2.0importos.py` lê periodicamente `creditos_inforow.txt` e sobrescreve `dados.xlsx`.
- **Como alterar com segurança:**
  Sempre use o utilitário:
  ```powershell
  & "C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe" pipeline_clealco.py set-creditos <VALOR> --compile
  ```
  Ele atualiza simultaneamente:
  1. `creditos_inforow.txt` (Fonte raiz)
  2. `dados.xlsx` (Aba `CONFIG`, célula `A2`)
  3. `index.html` (Fallbacks)
  4. `dados_injetados.js` (Criptografia AES-GCM)

### 🥈 Regra 2: Interpretador Python Correto
- **SEMPRE** execute os scripts geoespaciais e de compilação utilizando o Python do QGIS:
  `C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe`
- O Python padrão do Windows não possui `osgeo`, `ogr`, `gdal`, `geopandas` ou `cryptography`, o que causa falhas silenciosas na compilação do mapa e dos dados criptografados.

### 🥉 Regra 3: Criptografia e Service Worker
- Os dados do painel web são criptografados com **AES-GCM (PBKDF2-SHA256)** usando a chave de acesso corporativa (`Clealco@Geo2026`).
- Ao recompilar dados para a web, o `service-worker.js` deve ser incrementado (ex: `clealco-voos-v3715`) para que os navegadores dos usuários e celulares descarreguem o cache antigo.

### 🏅 Regra 4: dados.xlsx é Derivado e Ignorado
- O arquivo `dados.xlsx` é localmente gerado pelo backend a partir das bases oficiais.
- **NÃO** tente comitar `dados.xlsx` no Git (ele está no `.gitignore` por proteção de dados corporativos). Apenas os arquivos criptografados (`dados_injetados.js`, `dados_mapa_voos.js`, etc.) sobem para o GitHub.

### 💎 Regra 5: Fonte Oficial de Realizado de Falhas (W:\01-DRONE\08-BANCO DE DADOS\2026\Falhas)
- **FONTE DE VERDADE ABSOLUTA:** O Realizado de Falhas no Painel Web (Cards de Metas, Tabela de Acompanhamento e Gauges/Velocímetros) é alimentado e referenciado **exclusivamente** pelo Banco de Dados de Falhas:
  `W:\01-DRONE\08-BANCO DE DADOS\2026\Falhas`
- **Classificação:**
  - Arquivos `20XXXX.csv` (6 dígitos iniciando em `20`): **Próprias**.
  - Arquivos `30XXXX.csv` (6 dígitos iniciando em `30`): **Terceiros / Fornecedores**.
- **Coerência Geral:** Todo número de hectare realizado deve espelhar a soma dos CSVs dessa pasta, garantindo que os Cards de Metas, os Indicadores de Falhas no Plantio e os Velocímetros estejam sempre perfeitamente alinhados.


---

## 3. Comandos Rápidos com `pipeline_clealco.py`

| Objetivo | Comando |
| :--- | :--- |
| **Checar Saúde e Sincronia** | `& "C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe" pipeline_clealco.py status` |
| **Ajustar Créditos Inforow** | `& "C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe" pipeline_clealco.py set-creditos 10713.75 --compile` |
| **Inspecionar Dados Injetados** | `& "C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe" pipeline_clealco.py decrypt injetados` |
| **Inspecionar Mapa de Voos** | `& "C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe" pipeline_clealco.py decrypt mapa` |
| **Recompilar Dados Web** | `& "C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe" pipeline_clealco.py compile` |
| **Sincronizar e Enviar GitHub** | `& "C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe" pipeline_clealco.py sync` |
| **Diagnosticar Ambiente** | `& "C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe" pipeline_clealco.py diagnose` |

---

## 4. Fluxo de Vida do Dado (Resumo)
```text
Drone no Campo ➔ Pasta '1 - Não Processados' ➔ 2.0importos.py ➔ dados.xlsx ➔ atualizar_dados.bat ➔ atualizar_dados.py ➔ dados_injetados.js (AES) ➔ GitHub ➔ GitHub Pages (PWA)
```
