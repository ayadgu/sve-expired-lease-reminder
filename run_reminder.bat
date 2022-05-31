SET mypath="T:\INFORMATIQUE\IT\Bot Rappel Fin de Bail"
echo %mypath:~0,-1%
cd /d %mypath%
CALL .\venv\Scripts\activate.bat
> where python.exe -u rappel_fin_baux.py