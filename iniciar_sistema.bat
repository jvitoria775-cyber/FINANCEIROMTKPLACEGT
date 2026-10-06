@echo off
title Controle Financeiro Multicanal
cd /d "%~dp0"
echo ========================================================
echo INICIANDO CONTROLE FINANCEIRO E CONCILIACAO MULTICANAL
echo ========================================================
echo Abrindo navegador em http://localhost:8501 ...
start http://localhost:8501
.\.venv\Scripts\streamlit.exe run app.py
pause
