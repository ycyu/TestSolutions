@echo off
echo.
echo ========== %date% %time% ==========
echo.
echo ========== §å¦¸Ã±¥X ==========
echo.
del /f /q .\log\SCRIPTS_MCO.log
CScript //Nologo .\vbs\MCO.vbs .\param\TBB.param .\ini\SCRIPTS.ini .\csv\SCRIPTS2.csv > .\log\SCRIPTS_MCO.log
type .\log\SCRIPTS_MCO.log
echo.
echo ========== %date% %time% ==========
echo.
