@echo off
cd /d "%~dp0"
echo Atualizando dados locais...

:: 1. Executa o Python e corta o fluxo imediatamente se houver erro de script
python atualizar_dados.py || exit /b 1
echo.

echo Preparando envio para o GitHub...
:: 2. Previne conflitos baixando eventuais alteracoes da nuvem primeiro
git pull --rebase origin main

:: 3. Adiciona apenas os arquivos de dados CRIPTOGRAFADOS (dados.xlsx fica protegido localmente)
git add dados_injetados.js dados_mapa_voos.js

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