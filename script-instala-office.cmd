@echo off
chcp 65001 >nul
SETLOCAL ENABLEDELAYEDEXPANSION

:: Define a cor do fundo como preto e o texto como verde
color 0A

:: Verifica se o script está sendo executado como administrador
:: Tenta acessar a pasta "C:\Windows\System32", que requer privilégios elevados
net session >nul 2>&1

:: Verifica o código de erro
if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Este script requer privilégios de administrador.
    echo Por favor, feche o programa e execute-o como Administrador.
    pause
    exit /b
)

cd /d %~dp0

:: Se for administrador, continua a execução
echo Script rodando com privilégios de administrador.

:: Verifica a arquitetura do sistema usando WMIC
for /f "tokens=1 delims=- " %%a in ('wmic os get osarchitecture ^| find "bit"') do (
    set arquitetura=%%a
)

:: Exibe a arquitetura detectada
echo Arquitetura do sistema: %arquitetura%

:: Determina se o sistema é 32 ou 64 bits
if "%arquitetura%"=="64-bit" (
    echo Sistema operacional de 64 bits.
) else if "%arquitetura%"=="32-bit" (
    echo Sistema operacional de 32 bits.
) else (
    echo Nao foi possivel determinar a arquitetura.
)

:menu
cls
echo ===============================================
echo           Selecione a versão do Office
echo ===============================================
echo                 1. Office 2013
echo                 2. Office 2016
echo                 3. Office 2019
echo                 4. Office 2021
echo                 5. Office 2024
echo                 6. Office 365
echo                 0. Sair
echo ===============================================
set /p escolha="           Digite a opção desejada: "

:: Verifica a escolha do usuário
if /i "!escolha!"=="1" set "office=2013"
if /i "!escolha!"=="2" set "office=2016"
if /i "!escolha!"=="3" set "office=2019"
if /i "!escolha!"=="4" set "office=2021"
if /i "!escolha!"=="5" set "office=2024"
if /i "!escolha!"=="6" set "office=365"
if /i "!escolha!"=="0" goto sair

:: Se a escolha for inválida, retorna ao menu
if not defined office (
    echo Escolha inválida. Tente novamente.
    pause
    goto menu
)

:: Exibe a escolha do usuário e pede confirmação
cls
echo Você selecionou o Office !office! de !arquitetura! bits.
set /p confirm="Deseja continuar? (S/N): "

:: Verifica a confirmação do usuário
if /i "!confirm!"=="S" (
    echo Iniciando a instalação do Office !office!...
    :: Gera o comando de instalação para o Office
    (
    echo setup.exe /configure Config/Config-!office!-!arquitetura!.xml
    ) > install-office.cmd
    call install-office.cmd
    pause
) else if /i "!confirm!"=="N" (
    echo A instalação foi cancelada. Retorne ao Menu.
    pause
    goto menu
) else (
    echo Opção inválida. Retornando ao menu.
    pause
    goto menu
)

:sair
echo A instalação foi cancelada. Saindo.
pause
exit /b
