@echo off
echo.
echo ========== %date% %time% ==========
echo.
echo ========== �妸ñ�J�Χ�s�{���K�n���� ==========
echo.
del /f /q .\log\SCRIPTS_MCI.log
CScript //Nologo .\vbs\MCI.vbs .\param\TBB.param .\ini\SCRIPTS.ini .\csv\SCRIPTS.csv > .\log\SCRIPTS_MCI.log
type .\log\SCRIPTS_MCI.log
echo.
echo ========== %date% %time% ==========
echo.
