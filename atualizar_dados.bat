@echo off
cd /d "%~dp0"

:: 1. Previne conflitos baixando eventuais alteracoes da nuvem primeiro
git pull --rebase origin main >nul 2>&1

echo Atualizando dados locais...
:: 2. Prioriza o interpretador Python do QGIS (com suporte a osgeo/ogr e shapefiles) com fallback para python padrao
set "PY_EXE=C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe"
if not exist "%PY_EXE%" set "PY_EXE=python"

"%PY_EXE%" atualizar_dados.py || exit /b 1
echo.

echo Preparando envio para o GitHub...
:: 3. Adiciona apenas os arquivos de dados CRIPTOGRAFADOS e cache do PWA (dados.xlsx fica protegido localmente)
git add dados_injetados.js dados_mapa_voos.js dados_historico_falhas.js dados_mapas_safra.js service-worker.js

:: 4. Verifica se existe algo novo. Se os dados forem iguais, sai silenciosamente sem dar erro de "nothing to commit"
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo.
    echo Nenhuma alteracao de dados detectada. Sincronizacao ignorada.
    exit /b 0
)

:: 5. Commit e Push blindados
git commit -m "Atualizacao automatica de dados: %date% %time%"
git push origin main
echo.
echo Painel atualizado com sucesso no GitHub!