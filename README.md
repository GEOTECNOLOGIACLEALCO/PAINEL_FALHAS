# Singularity AI - Painel Clealco Geotecnologia (GitHub Pages)

Painel executivo online 100% estático desenvolvido para publicação direta no **GitHub Pages**, com visualização de métricas da safra, orquestração de voos de drones, extração de linhas e processamento de falhas de plantio.

---

## 🚀 Como funciona

O painel foi adaptado para rodar diretamente no navegador sem necessidade de servidor backend (Node, Flask ou Python). 
Ele utiliza a biblioteca **SheetJS** para fazer a leitura e interpretação direta da planilha `dados.xlsx` contida no repositório:

1. **`BASE_PROCESSAMENTOS`**: Alimenta o painel principal, histórico de voos, ortomosaicos (TIF), linhas geradas, balanço de cobertura e filtros rápidos.
2. **`BACKLOG`**: Alimenta a janela do motor de backlog (Faltam Voar), com agrupamento por Talhão ou consolidado por FA.
3. **`FALHAS_PLANTIO`**: Alimenta o dashboard analítico de falhas com velocímetros (gauges) e donuts de atingimento de metas.
4. **`METAS_GLOBAIS`**: Define as metas de hectares para áreas próprias e de parceiros/fornecedores.

---

## 🌐 Como ativar no GitHub Pages

1. Suba todo o conteúdo desta pasta (`index.html`, `dados.xlsx`, `logo_clealco.png`, `manifest.json`, `service-worker.js`) para um repositório no seu GitHub.
2. No seu repositório no GitHub, clique em **Settings** (Configurações).
3. No menu lateral esquerdo, clique em **Pages**.
4. Em **Build and deployment** > **Branch**, selecione `main` (ou `master`) e a pasta `/(root)`.
5. Clique em **Save**.
6. Aguarde cerca de 1 a 2 minutos e o link online do seu painel estará disponível (exemplo: `https://seu-usuario.github.io/nome-do-repositorio/`).

---

## 🔄 Como atualizar os dados da safra

Sempre que houver novos voos e processamentos:
1. Atualize a planilha `dados.xlsx`.
2. Substitua o arquivo `dados.xlsx` no repositório do GitHub (via upload ou `git push`).
3. O painel online carregará automaticamente os dados atualizados!
