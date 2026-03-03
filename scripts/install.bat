@echo off
chcp 65001 > nul
title DocToScan - Установка
color 0F

echo ════════════════════════════════════════════
echo    DocToScan - Установка интеграции
echo ════════════════════════════════════════════
echo.

:: Проверка прав администратора
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [✗] Ошибка: Недостаточно прав!
    echo.
    echo Для установки необходимы права администратора.
    echo Запустите этот файл от имени администратора.
    echo.
    pause
    exit /b 1
)

echo [✓] Права администратора получены

:: Определение пути к программе
set "SCRIPT_DIR=%~dp0"
set "EXE_PATH=%SCRIPT_DIR%DocToScan.exe"

:: Проверка существования exe
if not exist "%EXE_PATH%" (
    echo [✗] Ошибка: Не найден файл DocToScan.exe
    echo.
    echo Убедитесь, что install.bat находится в той же папке, что и DocToScan.exe
    echo.
    pause
    exit /b 1
)

echo [✓] Найден: %EXE_PATH%

:: Установка через программу
echo.
echo Установка интеграции в контекстное меню...
"%EXE_PATH%" /install

if %errorlevel% neq 0 (
    echo.
    echo [✗] Ошибка при установке!
    pause
    exit /b %errorlevel%
)

:: Создание папок для логов
if not exist "%SCRIPT_DIR%Logs" mkdir "%SCRIPT_DIR%Logs"
if not exist "%SCRIPT_DIR%Temp" mkdir "%SCRIPT_DIR%Temp"

:: Предложение добавить в PATH
echo.
set /p add_to_path="Добавить папку программы в PATH? (д/н): "
if /i "%add_to_path%"=="д" (
    setx PATH "%PATH%;%SCRIPT_DIR%" /M >nul
    if %errorlevel% equ 0 (
        echo [✓] Программа добавлена в PATH
    ) else (
        echo [✗] Не удалось добавить в PATH
    )
)

echo.
echo ════════════════════════════════════════════
echo    УСТАНОВКА ЗАВЕРШЕНА!
echo ════════════════════════════════════════════
echo.
echo ✓ Программа установлена в: %SCRIPT_DIR%
echo ✓ Интеграция добавлена в контекстное меню
echo.
echo Теперь вы можете:
echo   - Нажать ПКМ на PDF/DOCX файл
echo   - Выбрать "Создать скан-копию"
echo.
pause