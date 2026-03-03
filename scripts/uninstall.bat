@echo off
chcp 65001 > nul
title DocToScan - Удаление
color 0F

echo ════════════════════════════════════════════
echo    DocToScan - Удаление интеграции
echo ════════════════════════════════════════════
echo.

:: Проверка прав администратора
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [✗] Ошибка: Недостаточно прав!
    echo.
    echo Для удаления необходимы права администратора.
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
    echo [⚠] Внимание: Файл DocToScan.exe не найден
    echo.
    echo Будет выполнено только удаление из реестра.
    echo.
)

:: Удаление через программу
if exist "%EXE_PATH%" (
    echo.
    echo Удаление интеграции из контекстного меню...
    "%EXE_PATH%" /uninstall
) else (
    echo.
    echo Программа не найдена, удаляем записи реестра вручную...
    reg delete "HKCR\.pdf\shell\DocToScan" /f >nul 2>&1
    reg delete "HKCR\.docx\shell\DocToScan" /f >nul 2>&1
    reg delete "HKCR\Applications\DocToScan.exe" /f >nul 2>&1
)

:: Удаление из PATH
echo.
echo Удаление из PATH...
setx PATH "%PATH:;%SCRIPT_DIR%=%" /M >nul 2>&1

:: Вопрос о сохранении данных
echo.
echo Хотите сохранить папки с логами и конфигурацию?
echo (если нет - они будут удалены)
set /p keep_data="Сохранить данные? (д/н): "

if /i not "%keep_data%"=="д" (
    echo.
    echo Удаление файлов данных...
    if exist "%SCRIPT_DIR%Logs" rmdir /s /q "%SCRIPT_DIR%Logs"
    if exist "%SCRIPT_DIR%Temp" rmdir /s /q "%SCRIPT_DIR%Temp"
    if exist "%SCRIPT_DIR%config.xml" del /f /q "%SCRIPT_DIR%config.xml"
    echo [✓] Файлы данных удалены
) else (
    echo [✓] Файлы данных сохранены
)

:: Предложение удалить программу
echo.
set /p delete_exe="Удалить файлы программы? (д/н): "

if /i "%delete_exe%"=="д" (
    echo.
    echo Удаление файлов программы...
    if exist "%EXE_PATH%" del /f /q "%EXE_PATH%"
    
    :: Удаление других файлов
    for %%f in ("%SCRIPT_DIR%*.dll" "%SCRIPT_DIR%*.pdb") do (
        if exist "%%f" del /f /q "%%f"
    )
    
    :: Удаление самого uninstall.bat
    del /f /q "%SCRIPT_DIR%uninstall.bat"
    
    echo [✓] Файлы программы удалены
)

echo.
echo ════════════════════════════════════════════
echo    УДАЛЕНИЕ ЗАВЕРШЕНО!
echo ════════════════════════════════════════════
echo.
echo ✓ Интеграция удалена из системы
echo ✓ Записи реестра очищены
echo.
pause