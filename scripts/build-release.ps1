#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Сборка DocToScan и создание дистрибутива
#>

param(
    [string]$Configuration = "Release",
    [string]$Version = "",
    [string]$OutputDir = "builds"
)

# Устанавливаем кодировку UTF-8 для консоли
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [System.Text.UTF8Encoding]::new()

# Цвета для вывода
$Host.UI.RawUI.ForegroundColor = "White"
$script:successColor = "Green"
$script:errorColor = "Red"
$script:infoColor = "Yellow"
$script:detailColor = "Cyan"

function Write-Success { Write-Host $args -ForegroundColor $script:successColor }
function Write-Error { Write-Host $args -ForegroundColor $script:errorColor }
function Write-Info { Write-Host $args -ForegroundColor $script:infoColor }
function Write-Detail { Write-Host $args -ForegroundColor $script:detailColor }

function Exit-WithError {
    param([string]$Message)
    Write-Host "`n" -NoNewline
    Write-Error ("=" * 60)
    Write-Error "ОШИБКА!"
    Write-Error ("=" * 60)
    Write-Error $Message
    Write-Error ("=" * 60)
    Write-Error "Сборка прервана!"
    exit 1
}

# Проверка dotnet
try {
    $dotnetVersion = dotnet --version
    Write-Detail "Dotnet SDK версия: $dotnetVersion"
} catch {
    Exit-WithError "Dotnet SDK не найден! Установите .NET 10.0 SDK"
}

Write-Info ("=" * 60)
Write-Info "DocToScan - Сборка дистрибутива"
Write-Info ("=" * 60)

# Определяем пути
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Resolve-Path "$scriptPath/.."
$solutionPath = "$repoRoot/DocToScan.slnx"
$csprojPath = "$repoRoot/DocToScan.csproj"

Write-Detail "Корень проекта: $repoRoot"

# Получаем версию
if ([string]::IsNullOrWhiteSpace($Version)) {
    Write-Info "`nПолучение версии из сборки..."
    
    $buildPropsPath = "$repoRoot/Directory.Build.props"
    if (Test-Path $buildPropsPath) {
        $content = Get-Content $buildPropsPath -Raw -Encoding UTF8
        if ($content -match '<Version>(.*?)</Version>') {
            $Version = $matches[1].Trim()
            Write-Detail "Версия из Directory.Build.props: $Version"
        }
    }
    
    if ([string]::IsNullOrWhiteSpace($Version)) {
        $Version = Get-Date -Format "yyyy.MM.dd"
        Write-Info "Версия не найдена, используется: $Version"
    }
}

# Очищаем версию
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

if (-not (Test-Path "$repoRoot/$OutputDir")) {
    $null = New-Item -Path "$repoRoot/$OutputDir" -ItemType Directory -Force
}

Write-Info "`nПараметры сборки:"
Write-Detail "  Версия:      $Version"
Write-Detail "  Конфиг:      $Configuration"
Write-Detail "  Дистрибутив: $distName"

# Создаем директории дистрибутива
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

Write-Success "Восстановление завершено"

# Публикация приложения
Write-Info "`nПубликация приложения через dotnet publish..."

$publishPath = "$repoRoot/$OutputDir/publish_output"
if (Test-Path $publishPath) {
    Remove-Item $publishPath -Recurse -Force
}

dotnet publish $csprojPath `
    --configuration $Configuration `
    --runtime win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:PublishTrimmed=false `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:Version=$Version `
    --output $publishPath

if ($LASTEXITCODE -ne 0) {
    Exit-WithError "Ошибка публикации!"
}

Write-Success "Публикация завершена"

# Проверяем, что exe создан
$exeFile = Get-ChildItem "$publishPath" -Filter "*.exe" -Recurse | Select-Object -First 1
if (-not $exeFile) {
    Exit-WithError "Исполняемый файл не найден после публикации!"
}

# Копирование файлов в дистрибутив
Write-Info "`nКопирование файлов в дистрибутив..."

# Основной exe
Copy-Item $exeFile.FullName "$distPath/DocToScan.exe"
Write-Detail "Скопирован DocToScan.exe"

# Копируем все DLL и связанные файлы из publish_output
$fileCount = 0
Get-ChildItem "$publishPath" -File -Recurse | Where-Object {
    $_.Name -notlike "*.pdb" -and
    $_.Name -notlike "*.xml" -and
    $_.Name -notlike "*.exe"
} | ForEach-Object {
    $relativePath = $_.FullName.Substring($publishPath.Length).TrimStart('\')
    $targetPath = "$distPath/$relativePath"
    $targetDir = Split-Path $targetPath -Parent
    
    if (-not (Test-Path $targetDir)) {
        $null = New-Item -Path $targetDir -ItemType Directory -Force
    }
    
    Copy-Item $_.FullName $targetPath -ErrorAction SilentlyContinue
    $fileCount++
}
Write-Detail "Скопировано файлов: $fileCount"

# Копирование нативных библиотек Pdfium
Write-Info "`nКопирование нативных библиотек Pdfium..."

# Создаём папку x64
if (-not (Test-Path "$distPath/x64")) {
    $null = New-Item -Path "$distPath/x64" -ItemType Directory -Force
    Write-Detail "Создана папка x64"
}

# Путь к NuGet пакетам
$nugetPackagesPath = "$env:USERPROFILE\.nuget\packages"
$pdfiumNativePackage = "pdfiumviewer.native.x86_64.v8-xfa"
$pdfiumVersion = "2018.4.8.256"

$pdfiumNativePath = "$nugetPackagesPath\$pdfiumNativePackage\$pdfiumVersion\build\x64\pdfium.dll"

if (Test-Path $pdfiumNativePath) {
    # Копируем в папку x64
    Copy-Item $pdfiumNativePath "$distPath/x64/pdfium.dll" -ErrorAction SilentlyContinue
    Write-Detail "Скопирована pdfium.dll в папку x64"
    
    # Также копируем в корень для надёжности
    Copy-Item $pdfiumNativePath "$distPath/pdfium.dll" -ErrorAction SilentlyContinue
    Write-Detail "Скопирована pdfium.dll в корень"
} else {
    Write-Error "Не найдена pdfium.dll в NuGet пакетах!"
    
    # Ищем в папке публикации
    $pdfiumFiles = Get-ChildItem $publishPath -Filter "pdfium.dll" -Recurse
    foreach ($pdfiumFile in $pdfiumFiles) {
        Copy-Item $pdfiumFile.FullName "$distPath/pdfium.dll"
        Copy-Item $pdfiumFile.FullName "$distPath/x64/pdfium.dll"
        Write-Detail "Найдена и скопирована pdfium.dll: $($pdfiumFile.FullName)"
    }
}

# Конфигурация
$configSource = "$repoRoot/config.xml"
if (Test-Path $configSource) {
    Copy-Item $configSource "$distPath/config.xml"
    Write-Detail "Скопирован config.xml"
} else {
    $defaultConfig = "$repoRoot/assets/config/default-config.xml"
    if (Test-Path $defaultConfig) {
        Copy-Item $defaultConfig "$distPath/config.xml"
        Write-Detail "Скопирован default-config.xml"
    }
}

# Скрипты установки
$installBat = "$repoRoot/scripts/install.bat"
if (Test-Path $installBat) {
    Copy-Item $installBat "$distPath/install.bat"
    Write-Detail "Скопирован install.bat"
}

$uninstallBat = "$repoRoot/scripts/uninstall.bat"
if (Test-Path $uninstallBat) {
    Copy-Item $uninstallBat "$distPath/uninstall.bat"
    Write-Detail "Скопирован uninstall.bat"
}

# Иконка
$iconSource = "$repoRoot/assets/icons/apps_icon.ico"
if (Test-Path $iconSource) {
    Copy-Item $iconSource "$distPath/app.ico"
    Write-Detail "Скопирована иконка"
}

# Документация
if (Test-Path "$repoRoot/README.md") {
    Copy-Item "$repoRoot/README.md" "$distPath/README.txt"
    Write-Detail "Скопирован README.md"
}

# Лицензия
if (Test-Path "$repoRoot/LICENSE.txt") {
    Copy-Item "$repoRoot/LICENSE.txt" "$distPath/LICENSE.txt"
    Write-Detail "Скопирована лицензия"
}

# Создание README для дистрибутива
Write-Info "`nСоздание README для дистрибутива..."

$readmeContent = @"
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

ТРЕБОВАНИЯ:
-----------
- Windows 10/11 (64-bit)
- Microsoft Word (только для конвертации DOCX файлов)

Версия: $Version
Дата сборки: $(Get-Date -Format "dd.MM.yyyy HH:mm")
"@

$readmeContent | Out-File "$distPath/README.txt" -Encoding UTF8
Write-Detail "Создан README.txt"

# Проверяем наличие pdfium.dll
if (Test-Path "$distPath/pdfium.dll") {
    Write-Success "✓ pdfium.dll присутствует в дистрибутиве"
} else {
    Write-Error "⚠ pdfium.dll отсутствует! PDF функциональность может не работать."
}

# Подсчет файлов
$distFileCount = (Get-ChildItem $distPath -Recurse -File).Count
Write-Detail "Всего файлов в дистрибутиве: $distFileCount"

# Создание ZIP
Write-Info "`nСоздание ZIP-архива..."
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

try {
    Add-Type -Assembly System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($distPath, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
    Write-Detail "Архив создан через System.IO.Compression"
} catch {
    Write-Detail "Ошибка: $_"
    Write-Detail "Пробуем Compress-Archive..."
    Compress-Archive -Path "$distPath/*" -DestinationPath $zipPath -Force
}

if (-not (Test-Path $zipPath)) {
    Exit-WithError "Не удалось создать ZIP-архив!"
}

$zipSize = [math]::Round((Get-Item $zipPath).Length / 1MB, 2)

# Очистка
# Write-Info "`nОчистка временных файлов..."
# Remove-Item $distPath -Recurse -Force -ErrorAction SilentlyContinue
# Remove-Item $publishPath -Recurse -Force -ErrorAction SilentlyContinue

# Вывод результатов
Write-Success "`n" + ("=" * 60)
Write-Success "СБОРКА УСПЕШНО ЗАВЕРШЕНА!"
Write-Success ("=" * 60)
Write-Detail "Архив:     $zipPath"
Write-Detail "Размер:    $zipSize MB"
Write-Detail "Версия:    $Version"
Write-Detail "Дата:      $(Get-Date -Format 'dd.MM.yyyy HH:mm')"
Write-Success ("=" * 60)

# Показываем все сборки
Write-Info "`nВсе сборки в папке $OutputDir :"
$builds = Get-ChildItem "$repoRoot/$OutputDir" -Filter "*.zip" | Sort-Object LastWriteTime -Descending
if ($builds.Count -gt 0) {
    $builds | ForEach-Object {
        $fileSize = [math]::Round($_.Length / 1MB, 2)
        $fileDate = $_.LastWriteTime.ToString('dd.MM.yyyy HH:mm')
        Write-Detail ("  - " + $_.Name + " (" + $fileSize + " MB, " + $fileDate + ")")
    }
}

Write-Success "`nГотово!"