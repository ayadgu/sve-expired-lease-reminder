@REM Never run because UCN Link not accepted
@REM This repo is located on a share drive
@REM Hence this exact same bat is run from c:\

SET mypath="\\SRV-DC01.sve.local\Sve_Datas\INFORMATIQUE\IT\Bot Rappel Fin de Bail"

pushd %mypath%
echo [Guillaume] Lancement du script de verification des fins de baux...
rem cd /d %mypath%
CALL .\venv\Scripts\activate.bat
python "rappel_fin_baux.py"
popd
