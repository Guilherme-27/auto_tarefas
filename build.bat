@echo off
echo Finalizando qualquer .exe aberto...
taskkill /F /IM LancaNotas.exe >nul 2>&1
taskkill /F /IM app_gui.exe >nul 2>&1

echo Limpando arquivos antigos...
rmdir /s /q dist
rmdir /s /q build
del *.spec

echo Compilando .exe novamente...
python -m PyInstaller --name "LancaNotas" --onefile --noconsole --icon=icone.ico app_gui.py

echo.
echo ✅ Se não deu erro acima, o build foi feito com sucesso.
pause
