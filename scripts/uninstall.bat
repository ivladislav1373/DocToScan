@echo off
chcp 65001 > nul
title DocToScan - Удаление
color 0F

:: Запрос прав администратора
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ════════════════════════════════════════════
    echo       ТРЕБУЮТСЯ ПРАВА АДМИНИСТРАТОРА
    echo ════════════════════════════════════════════
    echo.
    echo Для удаления интеграции из контекстного меню
    echo необходимы права администратора.
    echo.
    echo Запуск от имени администратора...
    echo.

    :: Создаём временный VBS скрипт для запуска с правами администратора
    set "tempVbs=%temp%\getAdmin.vbs"
    echo Set UAC = CreateObject^("Shell.Application"^) > "%tempVbs%"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%tempVbs%"
    "%tempVbs%"
    del "%tempVbs%"
    exit /b
)

:: Если дошли до сюда - права администратора есть
echo ════════════════════════════════════════════
echo    DocToScan - Удаление интеграции
echo ════════════════════════════════════════════
echo.

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
for /f "tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do set "systemPath=%%b"
set "newPath=%systemPath:;%SCRIPT_DIR%=%"
setx PATH "%newPath%" /M >nul 2>&1

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
    
    :: Удаляем exe
    if exist "%EXE_PATH%" del /f /q "%EXE_PATH%"
    
    :: Удаляем все DLL и другие файлы
    for %%f in ("%SCRIPT_DIR%*.dll" "%SCRIPT_DIR%*.pdb" "%SCRIPT_DIR%*.ico" "%SCRIPT_DIR%*.xml") do (
        if exist "%%f" del /f /q "%%f"
    )
    
    :: Удаляем папки x64 и runtimes если они есть
    if exist "%SCRIPT_DIR%x64" rmdir /s /q "%SCRIPT_DIR%x64"
    if exist "%SCRIPT_DIR%runtimes" rmdir /s /q "%SCRIPT_DIR%runtimes"
    
    echo [✓] Файлы программы удалены
    
    :: Вопрос об удалении самого uninstall.bat
    echo.
    set /p delete_self="Удалить также этот скрипт (uninstall.bat)? (д/н): "
    if /i "%delete_self%"=="д" (
        del /f /q "%~f0"
        echo [✓] uninstall.bat будет удалён после закрытия
    )
)

echo.
echo ════════════════════════════════════════════
echo    УДАЛЕНИЕ ЗАВЕРШЕНО!
echo ════════════════════════════════════════════
echo.
echo ✓ Интеграция удалена из системы
echo ✓ Записи реестра очищены
echo.

if "%delete_self%"=="д" (
    echo Скрипт самоудалится через 2 секунды...
    ping -n 3 127.0.0.1 > nul
    del /f /q "%~f0"
) else (
    pause
)