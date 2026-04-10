@echo off
title Build - Consulta Tributaria
echo.
echo  ============================================
echo   Build - Consulta Tributaria
echo  ============================================
echo.

:: Verificar se o venv existe
if not exist "venv\Scripts\pip.exe" (
    echo ERRO: Ambiente virtual nao encontrado.
    echo Execute primeiro: python -m venv venv
    pause
    exit /b 1
)

:: Instalar / atualizar PyInstaller
echo [1/4] Instalando PyInstaller...
venv\Scripts\pip install pyinstaller --quiet --upgrade
if errorlevel 1 (
    echo ERRO: Falha ao instalar PyInstaller.
    pause
    exit /b 1
)

:: Limpar builds anteriores
echo [2/4] Limpando builds anteriores...
if exist "build" rmdir /s /q "build"
if exist "dist"  rmdir /s /q "dist"

:: Gerar o executavel
echo [3/4] Gerando executavel (pode demorar alguns minutos)...
venv\Scripts\pyinstaller consultatributaria.spec --clean --noconfirm
if errorlevel 1 (
    echo.
    echo ERRO: Falha ao gerar executavel.
    echo Verifique os logs acima para mais detalhes.
    pause
    exit /b 1
)

:: Montar pasta de distribuicao
echo [4/4] Preparando distribuicao...
set DIST=dist\ConsultaTributaria

if exist "%DIST%" rmdir /s /q "%DIST%"
mkdir "%DIST%"

copy "dist\ConsultaTributaria.exe" "%DIST%\ConsultaTributaria.exe" >nul
copy "TUTORIAL.txt"                "%DIST%\TUTORIAL.txt"           >nul

:: Sempre gerar um .env em branco — nunca copiar o .env local (tem credenciais)
(
    echo # Caminho completo para o executavel do navegador
    echo # Exemplos:
    echo #   Brave:  C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe
    echo #   Chrome: C:\Program Files\Google\Chrome\Application\chrome.exe
    echo BROWSER_PATH=
) > "%DIST%\.env"

echo.
echo  ============================================
echo   Concluido!
echo.
echo   Pasta de distribuicao:
echo   %CD%\%DIST%\
echo.
echo   Arquivos gerados:
echo   - ConsultaTributaria.exe
echo   - .env       ^(em branco - usuario deve configurar^)
echo   - TUTORIAL.txt
echo  ============================================
echo.
pause
