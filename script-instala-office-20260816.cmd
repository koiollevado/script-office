@echo off
chcp 1252 >nul
SETLOCAL ENABLEDELAYEDEXPANSION
color 0A

:: Verifica privilégios de Administrador
net session >nul 2>&1
if %errorlevel% neq 0 (
    cls
    echo.
    echo ===========================================================
    echo [ERRO] Este script requer privilegios de Administrador.
    echo Por favor, feche o programa e execute-o como Administrador.
    echo ===========================================================
    echo.
    pause
    exit /b 1
)

cd /d "%~dp0"
echo Script rodando com privilegios de administrador.
echo.

:: Detecta arquitetura do sistema (forma mais confiável)
set "arquitetura="
if defined PROCESSOR_ARCHITEW6432 (
    set "arquitetura=64"
) else (
    if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
        set "arquitetura=64"
    ) else (
        set "arquitetura=32"
    )
)

echo Arquitetura do sistema: !arquitetura!-bit
echo.

:menu
set "office="
set "escolha="
set "confirm="

cls
echo.
echo ===============================================
echo Selecione a versao do Office
echo ===============================================
echo 1. Office 2019
echo 2. Office 2021
echo 3. Office 2024
echo 4. Office 365
echo 0. Sair
echo ===============================================
echo.
set /p "escolha= Digite a opcao desejada: "

if /i "!escolha!"=="1" set "office=2019"
if /i "!escolha!"=="2" set "office=2021"
if /i "!escolha!"=="3" set "office=2024"
if /i "!escolha!"=="4" set "office=365"
if /i "!escolha!"=="0" goto :sair

if not defined office (
    cls
    echo.
    echo ===============================================
    echo Escolha invalida. Tente novamente.
    echo ===============================================
    echo.
    pause
    goto :menu
)

:: Confirmação
cls
echo.
echo ===============================================
echo Voce selecionou o Office !office! de !arquitetura!-bit.
echo ===============================================
echo.
set /p "confirm= Deseja continuar? (S/N): "

if /i "!confirm!"=="S" (
    cls
    echo.
    echo ===============================================
    echo Iniciando a instalacao do Office !office!...
    echo ===============================================
    echo.

    set "config_file=Config\Config-!office!-!arquitetura!.xml"

    if not exist "!config_file!" (
        echo [ERRO] Arquivo de configuracao nao encontrado:
        echo        !config_file!
        echo.
        pause
        goto :menu
    )

    if not exist "setup.exe" (
        echo [ERRO] setup.exe nao encontrado na pasta atual.
        echo.
        pause
        goto :menu
    )

    echo Executando: setup.exe /configure "!config_file!"
    echo.
    setup.exe /configure "!config_file!"

    echo.
    echo Instalacao finalizada (ou cancelada pelo setup).
    pause
    goto :menu

) else if /i "!confirm!"=="N" (
    cls
    echo.
    echo ===============================================
    echo A instalacao foi cancelada. Retornando ao Menu.
    echo ===============================================
    echo.
    pause
    goto :menu
) else (
    cls
    echo.
    echo ===============================================
    echo Opcao invalida. Retornando ao menu.
    echo ===============================================
    echo.
    pause
    goto :menu
)

:sair
cls
echo.
echo ===============================================
echo Saindo do script.
echo ===============================================
echo.
pause
ENDLOCAL
exit /b 0