@echo off
chcp 65001
title PhotoDoc Creator - Офлайн сборка

echo ===============================================
echo    PhotoDoc Creator - Офлайн сборка EXE
echo ===============================================
echo.

set PYTHONPATH=.;core;utils

echo Сборка EXE файла...
pyinstaller --onefile --windowed ^
--name "PhotoDoc_Creator_v4.5" ^
--icon=icon.ico ^
--add-data "core;core" ^
--add-data "utils;utils" ^
--hidden-import=PIL ^
--hidden-import=PIL._imaging ^
--hidden-import=docx ^
--hidden-import=docx.shared ^
--hidden-import=docx.enum.text ^
main.py

echo.
echo Готово! EXE файл находится в папке dist/
pause