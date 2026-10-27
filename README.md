# 🛰️ Singularity AI - Painel Clealco Geotecnologia (Voo & Processamentos 2.0)

Painel executivo corporativo de alto desempenho para gestão de voos de drones, auditoria agronômica, extração de linhas e processamento de falhas de plantio da **Usina Clealco**.

Desenvolvido para publicação direta no **GitHub Pages** como uma aplicação 100% estática, com arquitetura segura via **criptografia cliente (AES-GCM 256 bits)** e suporte a **Progressive Web App (PWA)** offline.

---

## 📑 Módulos do Sistema

O painel conta com **8 módulos operacionais integrados**:

1. **🛰️ Visão Geral:**
   - Métricas globais da safra, balanço de cobertura de voos (Área Própria vs Parcerias/Terceiros).
   - **Cards de Metas da Safra com indicador 'Falta Fazer':** Acompanhamento em tempo real de hectares e percentual restantes para 100% da meta, com alerta visual de meta batida.
   - Ranking de produtividade e alertas de inconsistências.

2. **📋 Acompanhamento de Voos & Processamentos:**
   - Tabela analítica completa com filtros rápidos por Unidade, Fazenda, Data, Piloto e Status.
   - Auditoria de ortomosaicos (TIF), linhas geradas (SHP/GeoJSON) e links diretos para a rede interna.

3. **🌾 Falhas de Plantio (Safra Atual):**
   - Dashboard analítico da safra corrente com velocímetros (*gauges*), semáforo corporativo e distribuição percentual de falhas.

4. **📊 Histórico de Falhas Multi-Safra (2021 a 2026):**
   - Série histórica de **6 safras consecutivas** compilando mais de **53.978 hectares**, **344.560 km de linhas** e **4,45 milhões de metros de falhas** em **630 fazendas** e **5.298 talhões**.
   - **Seletor Retrátil de Safras (*Accordion*):** Menu retrátil inteligente que mantém o visual limpo, com indicador dinâmico da safra ativa no topo.
   - **Semáforo Corporativo de Risco:**
     - *Áreas Próprias:* Aceitável $\le 8\%$ | Atenção $8,01\% - 12\%$ | Crítico $> 12\%$
     - *Parcerias/Terceiros:* Aceitável $\le 5\%$ | Atenção $5,01\% - 10\%$ | Crítico $> 10\%$
   - Gráficos dinâmicos (evolução temporal de barras e rosca de distribuição) e tabela detalhada por fazenda e talhão com exportação para Excel (XLSX).

5. **✈️ Backlog (Faltam Voar):**
   - Motor de inteligência que cruza o planejamento com as áreas já voadas, identificando talhões pendentes com agrupamento por FA ou Talhão.
   - Exportação em lote de arquivos **KML** diretamente para os drones de mapeamento (DJI Pilot).

6. **🗺️ Mapa Interativo, Radar IPMet & GPS de Campo:**
   - Visualização geoespacial vetorial via **Leaflet** com limites de talhões, ortomosaicos e polígonos de plantio.
   - **Radar Meteorológico (IPMet Bauru / Presidente Prudente):** Detecção de nuvens de chuva em tempo real com camada semitransparente Doppler sobreposta aos talhões.
   - **GPS de Campo:** Localização em tempo real para os pilotos e fiscais na lavoura com precisão métrica.

7. **🌧️ Clima & Chuvas:**
   - Monitoramento das condições meteorológicas locais, pluviometria acumulada e previsão do tempo para planejamento de voos.

8. **📑 Biblioteca de Mapas de Safra em PDF:**
   - Catálogo de mais de **1.300 croquis operacionais** divididos em Próprias, Acionistas, Terceiros e Spot.
   - Visualizador de PDF em tela cheia integrado no navegador e detecção automática de mapas recém-atualizados (*badge* `[NOVO]`).

---

## ⚙️ Fluxo de Atualização de Dados (Passo a Passo)

A atualização de dados é 100% automatizada e não exige conhecimentos avançados de programação:

```
[Rede W: e F:]
   ├── dados.xlsx (Planilha Mestra)
   ├── 08-BANCO DE DADOS (2021 a 2026 / Falhas)
   └── F:\Mapas\01 - Mapas Safra (PDFs)
                │
                ▼
   [Executar: atualizar_dados.bat]
                │
                ▼
   [Script Python: atualizar_dados.py]
      • Minera dados.xlsx e bases de falhas 2021-2026
      • Valida cabeçalhos (descarta arquivos de pisoteio)
      • Criptografa payloads em AES-GCM (Clealco@Geo2026)
      • Gera dados_injetados.js, dados_historico_falhas.js, etc.
      • Incrementa a versão do cache no service-worker.js
      • Executa git add, commit e push para o GitHub
                │
                ▼
   [GitHub Pages: Online em ~1 minuto]
```

### 🚀 Para atualizar:
1. Abra a pasta do projeto: `W:\00-SITE QUALIDADE APLICACAO\Voo_e_processamentosparagithub_2.0`
2. Dê um **duplo clique no arquivo `atualizar_dados.bat`**.
3. O terminal executará a compilação completa, a criptografia dos dados e enviará automaticamente para o GitHub.
4. Ao abrir o painel online, o Service Worker detecta a nova versão e atualiza o painel instantaneamente.

---

## 🔒 Segurança & Blindagem de Dados

- **Zero Exposição de Dados Confidenciais:** Planilhas brutas, shapefiles e PDFs **não sobem** desprotegidos para o repositório público.
- **Criptografia Militar:** Todos os bancos de dados sensíveis (`dados_injetados.js`, `dados_mapa_voos.js`, `dados_mapas_safra.js` e `dados_historico_falhas.js`) são criptografados com **AES-GCM (256 bits)**.
- **Leitura Segura:** A sincronização com as unidades de rede (`W:` e `F:`) opera em modo estritamente **somente leitura**, garantindo a integridade dos arquivos originais da usina.

---

## 📖 Documentação Completa

Para visualizar o manual didático com fluxogramas interativos Mermaid, regras de negócio e checklist diário do operador, abra no navegador o arquivo:
👉 **`MANUAL_PASSO_A_PASSO_PROJETO.html`**
