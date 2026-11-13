@echo off
cd /d "%~dp0"

:: 1. Previne conflitos baixando eventuais alteracoes da nuvem primeiro
git pull --rebase origin main >nul 2>&1

echo Atualizando dados locais...
:: 2. Prioriza o interpretador Python do QGIS (com suporte a osgeo/ogr e shapefiles) com deteccao multi-estacao
set "PY_EXE="
if exist "C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe" (
    set "PY_EXE=C:\Program Files\QGIS 4.2.1\apps\Python312\python.exe"
    set "PYTHONHOME=C:\Program Files\QGIS 4.2.1\apps\Python312"
    set "PATH=C:\Program Files\QGIS 4.2.1\apps\Python312;C:\Program Files\QGIS 4.2.1\apps\Python312\Scripts;C:\Program Files\QGIS 4.2.1\bin;C:\Program Files\QGIS 4.2.1\apps\gdal\bin;C:\Program Files\QGIS 4.2.1\apps\qt5\bin;%PATH%"
    set "PROJ_DATA=C:\Program Files\QGIS 4.2.1\share\proj"
    set "PROJ_LIB=C:\Program Files\QGIS 4.2.1\share\proj"
    set "GDAL_DATA=C:\Program Files\QGIS 4.2.1\apps\gdal\share\gdal"
    set "PYTHONUTF8=1"
)
if not defined PY_EXE if exist "C:\Program Files\QGIS 3.22.6\bin\python-qgis-ltr.bat" set "PY_EXE=C:\Program Files\QGIS 3.22.6\bin\python-qgis-ltr.bat"
if not defined PY_EXE for /d %%G in ("C:\Program Files\QGIS*") do (
    if not defined PY_EXE if exist "%%G\bin\python-qgis-ltr.bat" set "PY_EXE=%%G\bin\python-qgis-ltr.bat"
    if not defined PY_EXE if exist "%%G\bin\python-qgis.bat" set "PY_EXE=%%G\bin\python-qgis.bat"
    if not defined PY_EXE if exist "%%G\apps\Python312\python.exe" set "PY_EXE=%%G\apps\Python312\python.exe"
    if not defined PY_EXE if exist "%%G\apps\Python311\python.exe" set "PY_EXE=%%G\apps\Python311\python.exe"
    if not defined PY_EXE if exist "%%G\apps\Python39\python.exe" set "PY_EXE=%%G\apps\Python39\python.exe"
)
if not defined PY_EXE if exist "C:\OSGeo4W\bin\python-qgis.bat" set "PY_EXE=C:\OSGeo4W\bin\python-qgis.bat"
if not defined PY_EXE set "PY_EXE=python"

cmd.exe /c ""%PY_EXE%" atualizar_dados.py" || exit /b 1
echo.

echo Preparando envio para o GitHub...
:: 3. Adiciona apenas os arquivos de dados CRIPTOGRAFADOS e cache do PWA (dados.xlsx fica protegido localmente)
git add dados_injetados.js dados_mapa_voos.js dados_historico_falhas.js dados_mapas_safra.js dados_arable_clima.js dados_coordenadas_fas.js service-worker.js

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
if defined PY_EXE (
    "%PY_EXE%" exibir_resumo_console.py
    "%PY_EXE%" notificar_windows.py "Clealco Geotecnologia" "Base sincronizada com o GitHub com sucesso!" >nul 2>&1
) else (
    python exibir_resumo_console.py
    python notificar_windows.py "Clealco Geotecnologia" "Base sincronizada com o GitHub com sucesso!" >nul 2>&1
)
exit /b 0