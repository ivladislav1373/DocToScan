#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Сборка DocToScan и создание дистрибутива
.DESCRIPTION
    Скрипт собирает проект, создает структуру дистрибутива и упаковывает в ZIP
#>

# Устанавливаем кодировку UTF-8 для консоли
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [System.Text.UTF8Encoding]::new()

# Цвета для красивого вывода
$Host.UI.RawUI.ForegroundColor = "White"
$script:successColor = "Green"
$script:errorColor = "Red"
$script:infoColor = "Yellow"
$script:detailColor = "Cyan"

function Write-Success { Write-Host $args -ForegroundColor $script:successColor }
function Write-Error { Write-Host $args -ForegroundColor $script:errorColor }
function Write-Info { Write-Host $args -ForegroundColor $script:infoColor }
function Write-Detail { Write-Host $args -ForegroundColor $script:detailColor }

# Функция для выхода с ошибкой
function Exit-WithError {
    param([string]$Message)
    
    Write-Host "`n" -NoNewline
    Write-Error "=" * 60
    Write-Error "ОШИБКА!"
    Write-Error "=" * 60
    Write-Error $Message
    Write-Error "=" * 60
    Write-Error "Сборка прервана!"
    exit 1
}

# Проверка наличия dotnet
try {
    $dotnetVersion = dotnet --version
    Write-Detail "Dotnet SDK версия: $dotnetVersion"
} catch {
    Exit-WithError "Dotnet SDK не найден! Установите .NET 8.0 SDK"
}

Write-Info "`n" + "=" * 60
Write-Info "DocToScan - Сборка дистрибутива"
Write-Info "=" * 60

# Определяем корневую директорию проекта
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Resolve-Path "$scriptPath/../.."
$solutionPath = "$repoRoot/DocToScan.sln"

Write-Detail "Корень проекта: $repoRoot"

# Получаем версию из сборки, если не указана
if ([string]::IsNullOrWhiteSpace($Version)) {
    Write-Info "`nПолучение версии из сборки..."
    
    # Пытаемся прочитать версию из Directory.Build.props или AssemblyInfo
    $versionFile = "$repoRoot/Directory.Build.props"
    if (Test-Path $versionFile) {
        $content = Get-Content $versionFile -Raw
        if ($content -match '<Version>(.*?)</Version>') {
            $Version = $matches[1].Trim()
            Write-Detail "Версия из Directory.Build.props: $Version"
        }
    }
    
    if ([string]::IsNullOrWhiteSpace($Version)) {
        $Version = "1.0.0"
        Write-Info "Версия не найдена, используется: $Version"
    }
}

# Очищаем версию от лишних символов
$cleanVersion = $Version -replace '[^0-9.]', ''
if ($cleanVersion -ne $Version) {
    Write-Info "Версия очищена: $cleanVersion (было: $Version)"
    $Version = $cleanVersion
}

# Создаем имя дистрибутива
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$distName = "DocToScan_v${Version}_$timestamp"
$distPath = "$repoRoot/$OutputDir/$distName"
$zipPath = "$repoRoot/$OutputDir/${distName}.zip"

Write-Info "`nПараметры сборки:"
Write-Detail "  Версия:      $Version"
Write-Detail "  Конфиг:      $Configuration"
Write-Detail "  Дистрибутив: $distName"
Write-Detail "  Тесты:       $(-not $SkipTests)"

# Создаем директории
Write-Info "`nСоздание директорий..."
$null = New-Item -Path $distPath -ItemType Directory -Force
$null = New-Item -Path "$distPath/Logs" -ItemType Directory -Force
$null = New-Item -Path "$distPath/Temp" -ItemType Directory -Force

# Восстановление пакетов
Write-Info "`nВосстановление NuGet пакетов..."
dotnet restore $solutionPath
if ($LASTEXITCODE -ne 0) {
    Exit-WithError "Ошибка восстановления пакетов!"
}

# Публикация приложения
Write-Info "`nПубликация приложения..."
$publishPath = "$repoRoot/src/DocToScan.Console/bin/$Configuration/net8.0/win-x64/publish"
dotnet publish "$repoRoot/src/DocToScan.Console/DocToScan.Console.csproj" `
    --configuration $Configuration `
    --runtime win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:PublishTrimmed=true `
    -p:Version=$Version `
    --output $publishPath

if ($LASTEXITCODE -ne 0) {
    Exit-WithError "Ошибка публикации!"
}

# Копирование файлов
Write-Info "`nКопирование файлов в дистрибутив..."

# Основной exe
Copy-Item "$publishPath/DocToScan.Console.exe" "$distPath/DocToScan.exe"

# Нативные библиотеки
Copy-Item "$publishPath/*.dll" "$distPath/" -Exclude "*.pdb" -ErrorAction SilentlyContinue

# Конфигурация по умолчанию
Copy-Item "$repoRoot/assets/config/default-config.xml" "$distPath/config.xml"

# Скрипты установки
Copy-Item "$repoRoot/scripts/deploy/install.bat" "$distPath/"
Copy-Item "$repoRoot/scripts/deploy/uninstall.bat" "$distPath/"

# Иконка
Copy-Item "$repoRoot/assets/icons/app.ico" "$distPath/" -ErrorAction SilentlyContinue

# Документация
Copy-Item "$repoRoot/README.md" "$distPath/README.txt"
Copy-Item "$repoRoot/docs/user-guide.md" "$distPath/Руководство пользователя.txt"

# Создание README для дистрибутива
Write-Info "`nСоздание README для дистрибутива..."
@"
DocToScan v$Version
===================

Программа для преобразования документов в "скан-копии"

УСТАНОВКА:
----------
1. Распакуйте архив в любую папку (рекомендуется C:\Program Files\DocToScan)
2. Запустите install.bat от имени администратора для интеграции в контекстное меню
3. Для удаления запустите uninstall.bat от имени администратора

ИСПОЛЬЗОВАНИЕ:
--------------
- Через контекстное меню: ПКМ на PDF/DOCX файл -> "Создать скан-копию"
- Через командную строку: DocToScan.exe "путь\к\файлу.docx"

КОНФИГУРАЦИЯ:
-------------
Настройки программы находятся в файле config.xml:
- Brightness: настройка яркости
- Rotation: случайный поворот страниц
- Grayscale: черно-белый режим
- ImageQuality: качество изображений (DPI, сжатие JPEG)

ТРЕБОВАНИЯ:
-----------
- Windows 10/11 (64-bit)
- .NET 8.0 Runtime (включен в standalone версию)
- Microsoft Word (только для конвертации DOCX файлов)

ФАЙЛЫ ПРОГРАММЫ:
----------------
- DocToScan.exe - основная программа
- config.xml - файл конфигурации
- Logs\ - папка для логов (создается автоматически)
- Temp\ - временные файлы (создается автоматически)

Версия: $Version
Дата сборки: $(Get-Date -Format "dd.MM.yyyy HH:mm")
"@ | Out-File "$distPath/README.txt" -Encoding UTF8

# Упаковка в ZIP
Write-Info "`nСоздание ZIP-архива..."
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

try {
    # Используем .NET для лучшей компрессии
    Add-Type -Assembly System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($distPath, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
} catch {
    # Fallback на Compress-Archive
    Compress-Archive -Path "$distPath/*" -DestinationPath $zipPath -Force
}

# Проверка создания архива
if (-not (Test-Path $zipPath)) {
    Exit-WithError "Не удалось создать ZIP-архив!"
}

# Получаем размер
$zipSize = [math]::Round((Get-Item $zipPath).Length / 1MB, 2)

# Удаляем временную папку
Remove-Item $distPath -Recurse -Force

# Вывод результатов
Write-Success "`n" + "=" * 60
Write-Success "СБОРКА УСПЕШНО ЗАВЕРШЕНА!"
Write-Success "=" * 60
Write-Detail "Архив:     $zipPath"
Write-Detail "Размер:    $zipSize MB"
Write-Detail "Версия:    $Version"
Write-Detail "Дата:      $(Get-Date -Format 'dd.MM.yyyy HH:mm')"
Write-Success "=" * 60

# Показываем все сборки
Write-Info "`nВсе сборки в папке $OutputDir :"
$builds = Get-ChildItem "$repoRoot/$OutputDir" -Filter "*.zip" | Sort-Object LastWriteTime -Descending
if ($builds.Count -gt 0) {
    $builds | ForEach-Object {
        $fileSize = [math]::Round($_.Length / 1MB, 2)
        $fileDate = $_.LastWriteTime.ToString('dd.MM.yyyy HH:mm')
        Write-Detail "  - $($_.Name) ($fileSize MB, $fileDate)"
    }
} else {
    Write-Detail "  (нет архивов)"
}

Write-Success "`nГотово!"