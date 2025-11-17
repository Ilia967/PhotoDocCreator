@echo off
chcp 65001
title PhotoDoc Creator - Установка зависимостей

echo ===============================================
echo    Установка зависимостей PhotoDoc Creator
echo ===============================================
echo.

echo Установка python-docx...
pip install python-docx-1.1.0-py3-none-any.whl

echo Установка Pillow...
pip install Pillow-10.0.0-cp39-cp39-win32.whl

echo.
echo Все зависимости установлены!
pause