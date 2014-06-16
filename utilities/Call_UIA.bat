@echo off
echo.
echo ========== %date% %time% ==========
echo.
echo ========== 批次更新程式摘要說明 ==========
echo.
del /f /q .\log\SCRIPTS_UIA.log
CScript //Nologo .\vbs\UIA.vbs .\param\TBB.param .\ini\SCRIPTS.ini .\csv\SCRIPTS.csv > .\log\SCRIPTS_UIA.log
type .\log\SCRIPTS_UIA.log
echo.
echo ========== %date% %time% ==========
echo.
