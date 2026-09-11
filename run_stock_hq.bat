@echo off
cd /d "%~dp0"
echo ==== HAL Stock HQ - building today's report ====
python build_stock_hq.py --publish
echo.
echo ==== Done. Files are in this folder. Press any key to close. ====
pause >nul
