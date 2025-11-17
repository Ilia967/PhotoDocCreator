@echo off
chcp 65001
title PhotoDoc Creator - Сборка EXE

echo ===============================================
echo    PhotoDoc Creator - Сборка EXE файла
echo ===============================================
echo.

echo Установка auto-py-to-exe...
pip install auto-py-to-exe

echo.
echo Запуск auto-py-to-exe...
auto-py-to-exe

pause