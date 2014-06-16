@echo off
echo.
echo ========== %date% %time% ==========
echo.
echo ========== ²£¥XDEPLOY.DIM ==========
echo.
del /f /q .\log\SCRIPTS_DEPLOY.DIM
CScript //Nologo .\vbs\GenerateDeployDIM.vbs .\csv\SCRIPTS2.csv .\log\SCRIPTS_DEPLOY.DIM
type .\log\SCRIPTS_DEPLOY.DIM
echo.
echo ========== %date% %time% ==========
echo.
