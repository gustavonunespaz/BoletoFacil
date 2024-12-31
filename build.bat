@echo off
echo ===========================================
echo Criando o executável para BoletoFacil.py
echo ===========================================
pyinstaller --onefile --noconsole --icon=icone.ico BoletoFacil.py

echo ===========================================
echo Criando o executável para Compilador.py
echo ===========================================
pyinstaller --onefile --noconsole --icon=iconecompilador.ico Compilador.py

echo ===========================================
echo Concluído! Verifique os arquivos na pasta "dist".
echo ===========================================
pause
