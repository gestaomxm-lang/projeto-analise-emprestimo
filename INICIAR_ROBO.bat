@echo off
title Robo de Analise - Agendamento
echo ========================================================
echo   INICIANDO ROBO DE AUTOMACAO DE ANALISE
echo ========================================================
echo.
echo Este terminal deve ficar aberto para que a automacao funcione.
echo Você pode minimizar esta janela.
echo.
echo Iniciando scheduler_service.py...
echo.

cd /d "%~dp0"
python scheduler_service.py

pause
