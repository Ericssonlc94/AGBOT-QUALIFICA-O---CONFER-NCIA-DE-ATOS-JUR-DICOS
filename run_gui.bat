@echo off
echo Ativando ambiente virtual...
call venv\Scripts\activate.bat

echo Iniciando o Preenchedor de Qualificação (GUI)...
python app_gui.py

echo.
echo Processo finalizado.
pause