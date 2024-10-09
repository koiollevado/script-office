@echo off
chcp 1252 >nul
ENDLOCAL
SETLOCAL ENABLEDELAYEDEXPANSION

:: Define a cor do fundo como preto e o texto como verde
color 0A

:: Verifica se o script est� sendo executado como administrador
:: Tenta acessar a pasta "C:\Windows\System32", que requer privil�gios elevados
net session >nul 2>&1

:: Verifica o c�digo de erro
if %errorlevel% neq 0 (
    cls
    echo.
    echo ===========================================================
    echo [ERRO] Este script requer privil�gios de Administrador.
    echo Por favor, feche o programa e execute-o como Administrador.
    echo ===========================================================
    echo.
    pause
    exit /b
)

cd /d %~dp0

:: Se for administrador, continua a execu��o
echo Script rodando com privil�gios de administrador.

:: Verifica a arquitetura do sistema usando WMIC
for /f "tokens=1 delims=- " %%a in ('wmic os get osarchitecture ^| find "bit"') do (
    set arquitetura=%%a
)

:: Exibe a arquitetura detectada
echo Arquitetura do sistema: %arquitetura%

:: Determina se o sistema � 32 ou 64 bits
if "%arquitetura%"=="64-bit" (
    echo Sistema operacional de 64 bits.
) else if "%arquitetura%"=="32-bit" (
    echo Sistema operacional de 32 bits.
) else (
    echo N�o foi poss�vel determinar a arquitetura.
)

:menu
cls
echo.
echo ===============================================
echo          Selecione a vers�o do Office
echo ===============================================
echo                 1. Office 2019
echo                 2. Office 2021
echo                 3. Office 2024
echo                 4. Office 365
echo                 0. Sair
echo ===============================================
echo.
set /p escolha="           Digite a op��o desejada: "

:: Verifica a escolha do usu�rio
if /i "!escolha!"=="1" set "office=2019"
if /i "!escolha!"=="2" set "office=2021"
if /i "!escolha!"=="3" set "office=2024"
if /i "!escolha!"=="4" set "office=365"
if /i "!escolha!"=="0" goto sair

:: Se a escolha for inv�lida, retorna ao menu
if not defined office (
    cls
    echo.
    echo ===============================================
    echo        Escolha inv�lida. Tente novamente.
    echo ===============================================
    echo.
    pause
    goto menu
)

:: Exibe a escolha do usu�rio e pede confirma��o
cls
echo.
echo ===============================================
echo    Voc� selecionou o Office !office! de !arquitetura! bits.
echo ===============================================
echo.
set /p confirm="           Deseja continuar? (S/N): "

:: Verifica a confirma��o do usu�rio
if /i "!confirm!"=="S" (
    cls    
    echo.
    echo ===============================================
    echo    Iniciando a instala��o do Office !office!...
    echo ===============================================
    echo.
    :: Gera o comando de instala��o para o Office
    echo setup.exe /configure Config/Config-!office!-!arquitetura!.xml > install-office.cmd
    call install-office.cmd
    pause
) else if /i "!confirm!"=="N" (
    cls
    echo.
    echo ===============================================
    echo   A instala��o foi cancelada. Retorne ao Menu.
    echo ===============================================
    echo.
    pause
    goto menu
) else (
    cls
    echo.
    echo ===============================================
    echo        Op��o inv�lida. Retornando ao menu.
    echo ===============================================
    echo.
    pause
    goto menu
)

:sair
cls
echo.
echo ===============================================
echo       A instala��o foi cancelada. Saindo.
echo ===============================================
echo.
pause
exit /b
