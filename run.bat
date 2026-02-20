@echo off
chcp 65001 > nul
title Авто-печать Excel

:: Проверяем наличие Python
where python >nul 2>&1
if errorlevel 1 (
    echo Python не найден. Установите Python 3.8+ и добавьте в PATH.
    pause
    exit /b 1
)

:: Устанавливаем зависимости, если не установлены
python -m pip show watchdog >nul 2>&1
if errorlevel 1 (
    echo Установка зависимостей...
    python -m pip install -r requirements.txt
)

echo Мониторинг папки Загрузки запущен.
echo Файлы, начинающиеся с "ОС-2" или "М11", будут выводиться на печать после подтверждения.
echo Закройте это окно для остановки программы.
echo.

pythonw auto_print.py
if errorlevel 1 (
    python auto_print.py
)
