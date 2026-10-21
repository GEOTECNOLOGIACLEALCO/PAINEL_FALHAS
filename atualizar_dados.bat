@echo off
cd /d " %~dp0\
echo Atualizando dados locais...
python atualizar_dados.py
echo.
echo Enviando para o GitHub...
git add dados.xlsx dados_injetados.js
git commit -m \Atualizacao automatica da planilha dados.xlsx\
git push origin main
echo.
echo Painel atualizado com sucesso no GitHub!
