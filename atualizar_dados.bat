@echo off
cd /d "%~dp0"
echo Atualizando dados locais...
python atualizar_dados.py
echo.
echo Enviando para o GitHub...
git add .
git commit -m "Atualizacao de dados: %date% %time%"
git push origin main
echo.
echo Painel atualizado com sucesso no GitHub!