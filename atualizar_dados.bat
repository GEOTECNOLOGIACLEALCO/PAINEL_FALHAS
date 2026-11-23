@echo off
cd /d "%~dp0"

:: 0. Deteccao automatica do executavel do Git em qualquer estacao ou usuario
set "GIT_CMD="
where git >nul 2>&1 && set "GIT_CMD=git"
if not defined GIT_CMD if exist "C:\Program Files\Git\cmd\git.exe" set "GIT_CMD=C:\Program Files\Git\cmd\git.exe"
if not defined GIT_CMD if exist "C:\Program Files\Git\bin\git.exe" set "GIT_CMD=C:\Program Files\Git\bin\git.exe"
if not defined GIT_CMD if exist "C:\Program Files (x86)\Git\cmd\git.exe" set "GIT_CMD=C:\Program Files (x86)\Git\cmd\git.exe"
if not defined GIT_CMD if exist "C:\Program Files (x86)\Git\bin\git.exe" set "GIT_CMD=C:\Program Files (x86)\Git\bin\git.exe"
if not defined GIT_CMD if exist "%LOCALAPPDATA%\Programs\Git\cmd\git.exe" set "GIT_CMD=%LOCALAPPDATA%\Programs\Git\cmd\git.exe"
if not defined GIT_CMD if exist "%LOCALAPPDATA%\Programs\Git\bin\git.exe" set "GIT_CMD=%LOCALAPPDATA%\Programs\Git\bin\git.exe"
if not defined GIT_CMD if exist "C:\Users\fabianocastro\AppData\Local\Programs\Git\cmd\git.exe" set "GIT_CMD=C:\Users\fabianocastro\AppData\Local\Programs\Git\cmd\git.exe"
if not defined GIT_CMD if exist "C:\Git\cmd\git.exe" set "GIT_CMD=C:\Git\cmd\git.exe"
if not defined GIT_CMD if exist "C:\PortableGit\cmd\git.exe" set "GIT_CMD=C:\PortableGit\cmd\git.exe"
if not defined GIT_CMD for /d %%U in ("C:\Users\*") do (
    if not defined GIT_CMD if exist "%%U\AppData\Local\Programs\Git\cmd\git.exe" set "GIT_CMD=%%U\AppData\Local\Programs\Git\cmd\git.exe"
    if not defined GIT_CMD if exist "%%U\AppData\Local\Programs\Git\bin\git.exe" set "GIT_CMD=%%U\AppData\Local\Programs\Git\bin\git.exe"
)
if defined GIT_CMD (
    if not "%GIT_CMD%"=="git" for %%I in ("%GIT_CMD%") do set "PATH=%%~dpI;%PATH%"
) else (
    set "GIT_CMD=git"
)

:: 1. Previne conflitos baixando eventuais alteracoes da nuvem primeiro
"%GIT_CMD%" pull --rebase origin main >nul 2>&1

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
where "%GIT_CMD%" >nul 2>&1
if %errorlevel% neq 0 if not exist "%GIT_CMD%" (
    echo [AVISO] Git nao encontrado neste computador. O envio ao GitHub sera ignorado.
    if defined PY_EXE (
        "%PY_EXE%" exibir_resumo_console.py
    ) else (
        python exibir_resumo_console.py
    )
    exit /b 0
)

:: 3. Adiciona apenas os arquivos de dados CRIPTOGRAFADOS e cache do PWA (dados.xlsx fica protegido localmente)
"%GIT_CMD%" add dados_injetados.js dados_mapa_voos.js dados_historico_falhas.js dados_mapas_safra.js dados_arable_clima.js dados_coordenadas_fas.js service-worker.js index.html atualizar_dados.py exibir_resumo_console.py

:: 4. Verifica se existe algo novo. Se os dados forem iguais, sai silenciosamente sem dar erro de "nothing to commit"
"%GIT_CMD%" diff --cached --quiet
if %errorlevel% equ 0 (
    echo.
    echo Nenhuma nova alteracao de dados detectada. Base de dados 100%% atualizada.
    if defined PY_EXE (
        "%PY_EXE%" exibir_resumo_console.py
    ) else (
        python exibir_resumo_console.py
    )
    exit /b 0
)

:: 5. Commit e Push blindados
"%GIT_CMD%" commit -m "Atualizacao automatica de dados: %date% %time%"
"%GIT_CMD%" push origin main
echo.
if defined PY_EXE (
    "%PY_EXE%" exibir_resumo_console.py
    "%PY_EXE%" notificar_windows.py "Clealco Geotecnologia" "Base sincronizada com o GitHub com sucesso!" >nul 2>&1
) else (
    python exibir_resumo_console.py
    python notificar_windows.py "Clealco Geotecnologia" "Base sincronizada com o GitHub com sucesso!" >nul 2>&1
)
exit /b 0