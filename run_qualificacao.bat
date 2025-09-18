@echo off

echo Ativando ambiente virtual...
call venv\Scripts\activate.bat

echo Executando o Preenchedor de Qualificacao...
python preenchedor_qualificacao.py

echo.
pause