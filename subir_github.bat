@echo off
title Enviar Projeto para o GitHub - Gratitude Textil
cd /d "%~dp0"
echo ============================================================
echo   ENVIANDO PROJETO PARA O GITHUB: FINANCEIROMTKPLACEGT
echo ============================================================
echo.
git push -u origin main
echo.
echo ============================================================
if %ERRORLEVEL% EQU 0 (
    echo [SUCESSO] Projeto enviado com sucesso para o GitHub!
) else (
    echo [AVISO] Se solicitou login, faca a autenticacao no navegador.
)
echo ============================================================
pause
