@echo off
chcp 65001 >nul
title Robo de Consulta Processual - SEBRAE Contencioso
cd /d "%~dp0"

echo ============================================================
echo    ROBO DE CONSULTA PROCESSUAL - SEBRAE Contencioso
echo ============================================================
echo.
echo  Na primeira vez este programa instala tudo o que precisa
echo  automaticamente, DENTRO desta mesma pasta. Isso pode levar
echo  alguns minutos. Aguarde.
echo.

REM ================================================================
REM  1) Procura o Python. Se nao existir, instala automaticamente.
REM ================================================================
call :ACHAR_PYTHON
if defined PY goto TEM_PYTHON

echo  [*] Python nao encontrado. Instalando automaticamente...
echo.

REM ---- Tentativa A: winget (Windows 10/11) ----
where winget >nul 2>&1
if %errorlevel%==0 (
  echo  [*] Instalando via winget ^(pode aparecer uma janela de permissao^)...
  winget install -e --id Python.Python.3.12 --silent --scope user --accept-source-agreements --accept-package-agreements
)

call :ACHAR_PYTHON
if defined PY goto TEM_PYTHON

REM ---- Tentativa B: baixar o instalador oficial e instalar em silencio ----
echo  [*] Baixando o instalador oficial do Python...
set "PYEXE=%TEMP%\python-instalador.exe"
curl -L -o "%PYEXE%" "https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe"
if not exist "%PYEXE%" (
  powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe' -OutFile '%PYEXE%'"
)
if exist "%PYEXE%" (
  echo  [*] Instalando o Python ^(aguarde, pode demorar alguns minutos^)...
  "%PYEXE%" /quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1 Include_pip=1
  del "%PYEXE%" >nul 2>&1
)

call :ACHAR_PYTHON
if defined PY goto TEM_PYTHON

echo.
echo  [ATENCAO] O Python foi instalado, mas o Windows precisa reconhecer.
echo  Por favor FECHE esta janela e abra o arquivo novamente ^(duplo clique^).
echo.
pause
exit /b

:TEM_PYTHON
echo  [1/4] Python pronto.

REM ================================================================
REM  2) Cria o ambiente isolado (venv) DENTRO desta pasta: .venv
REM     Assim nada e instalado no sistema e nao precisa de admin.
REM ================================================================
set "VENV=%~dp0.venv"
set "VPY=%VENV%\Scripts\python.exe"

if not exist "%VPY%" (
  echo  [2/4] Criando ambiente isolado nesta pasta ^(.venv^)...
  "%PY%" -m venv "%VENV%"
)
if not exist "%VPY%" (
  echo  [ATENCAO] Nao foi possivel criar o ambiente isolado.
  echo  Tire uma foto desta tela e envie para o suporte.
  pause
  exit /b
)

REM ================================================================
REM  3) Instala as bibliotecas dentro do venv (apenas se faltarem)
REM ================================================================
"%VPY%" -c "import requests, pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo  [3/4] Instalando componentes necessarios ^(so na primeira vez^)...
  "%VPY%" -m pip install --upgrade --disable-pip-version-check pip >nul 2>&1
  "%VPY%" -m pip install --disable-pip-version-check requests pandas openpyxl
) else (
  echo  [3/4] Componentes ja instalados.
)

REM ================================================================
REM  4) Abre a janela do robo (usando o Python do venv)
REM ================================================================
echo  [4/4] Abrindo a janela do robo...
echo.
"%VPY%" "robo_consulta_gui.py"

if errorlevel 1 (
  echo.
  echo  [ATENCAO] Ocorreu um erro ao abrir o robo.
  echo  Tire uma foto desta tela e envie para o suporte.
  echo.
  pause
)
exit /b


REM ================================================================
REM  Sub-rotina: localiza o Python do sistema e guarda em %PY%
REM ================================================================
:ACHAR_PYTHON
set "PY="
where py >nul 2>&1 && set "PY=py"
if not defined PY (
  where python >nul 2>&1 && set "PY=python"
)
if not defined PY (
  for /d %%D in ("%LocalAppData%\Programs\Python\Python3*") do (
    if exist "%%D\python.exe" set "PY=%%D\python.exe"
  )
)
if not defined PY (
  for /d %%D in ("%ProgramFiles%\Python3*") do (
    if exist "%%D\python.exe" set "PY=%%D\python.exe"
  )
)
exit /b
