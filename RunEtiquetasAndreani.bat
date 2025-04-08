@echo off
echo Ejecutando Script Imprimir_Etiquetas_Andreani.py
echo Activando entorno virtual

call env/Scripts/activate.bat
echo Ejecutando Script
cd ScriptPY/src
python main.py
pause
