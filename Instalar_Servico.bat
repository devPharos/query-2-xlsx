@echo off

echo ===============================================
echo   Desenvolvido por Denis Varella - Maio/2025
echo ===============================================
echo.

:: Verifica se o script esta rodando como Administrador
NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo ===============================================
    echo  ERRO: Este script precisa ser executado como administrador!
    echo  Clique com o botao direito no arquivo e selecione:
    echo     "Executar como administrador"
    echo ===============================================
    pause
    exit /b
)

REM === PEGA O DIRETORIO DO SCRIPT ATUAL ===
SET "ROOT_DIR=%~dp0"
SET "APP_DIR=%ROOT_DIR%"
SET "NSSM_DIR=%ROOT_DIR%nssm-2.24\win64"

REM === MUDA PARA O DIRETORIO ONDE ESTA O NSSM.EXE ===
cd /d "%NSSM_DIR%"

REM === SOLICITA O NOME DO SERVICO AO USUARIO ===
set /p NOME_SERVICO=Digite o nome do servico que deseja criar: 

REM === DEFINE O SCRIPT EXECUTAVEL ===
SET "SCRIPT_EXECUTAVEL=%APP_DIR%service.bat"

REM === CRIACAO DO SERVICO ===
nssm install "%NOME_SERVICO%" "C:\Windows\System32\cmd.exe" "/C \"%SCRIPT_EXECUTAVEL%\""

REM === CONFIGURACOES ADICIONAIS DO SERVICO ===
nssm set "%NOME_SERVICO%" AppDirectory "%APP_DIR%"
nssm set "%NOME_SERVICO%" Description "Gerador de arquivo XLSX por query."
nssm set "%NOME_SERVICO%" Start "SERVICE_AUTO_START"

REM === INSTRUCOES FINAIS ===
echo.
echo ===============================================
echo Servico "%NOME_SERVICO%" criado e configurado com sucesso!
echo.
echo Agora:
echo 1. Abra o "Servicos" do Windows.
echo 2. Localize o servico "%NOME_SERVICO%".
echo 3. Clique com o botao direito > Propriedades > Aba "Logon"
echo 4. Substitua a conta "Sistema local" (SYSTEM) por um usuario com permissoes suficientes.
echo    Esse usuario deve ter acesso a rede, pastas e permissoes adequadas para executar o script.
echo ===============================================
echo.

REM === CONFIRMACAO FINAL ===
choice /M "Deseja manter o servico criado?"
IF ERRORLEVEL 2 (
    echo.
    echo Removendo o servico "%NOME_SERVICO%"...
    nssm remove "%NOME_SERVICO%" confirm
    echo Servico removido.
) ELSE (
    echo.
    echo O servico foi mantido.
)

pause
