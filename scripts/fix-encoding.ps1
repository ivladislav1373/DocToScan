# Этот скрипт нужно сохранить в ANSI кодировке
$scriptsPath = "G:\Projects\C#\DocToScan\scripts"

# Создаем временную папку
$tempPath = "$env:TEMP\doc2scan_scripts"
New-Item -ItemType Directory -Path $tempPath -Force | Out-Null

# Функция для создания файла с правильной кодировкой
function Create-File {
    param($filename, $content)
    
    $filePath = "$scriptsPath\$filename"
    $tempFilePath = "$tempPath\$filename"
    
    # Сохраняем во временный файл в UTF-8 без BOM
    [System.IO.File]::WriteAllText($tempFilePath, $content, [System.Text.UTF8Encoding]::new($false))
    
    # Копируем
    Copy-Item $tempFilePath $filePath -Force
    
    Write-Host "Создан: $filename" -ForegroundColor Green
}

# Создаем fix-all.ps1
$fixAllContent = @'
#!/usr/bin/env pwsh

# Устанавливаем кодировку UTF-8 для консоли
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [System.Text.UTF8Encoding]::new()

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  DocToScan - Исправление всех проблем" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

$rootPath = "G:\Projects\C#\DocToScan"

# 1. Очистка старых сборок
Write-Host "Шаг 1: Очистка старых сборок..." -ForegroundColor Yellow
if (Test-Path "$rootPath\bin") {
    Remove-Item -Recurse -Force "$rootPath\bin" -ErrorAction SilentlyContinue
    Write-Host "  Папка bin удалена" -ForegroundColor Green
}
if (Test-Path "$rootPath\obj") {
    Remove-Item -Recurse -Force "$rootPath\obj" -ErrorAction SilentlyContinue
    Write-Host "  Папка obj удалена" -ForegroundColor Green
}

# 2. Очистка кэша NuGet
Write-Host "`nШаг 2: Очистка кэша NuGet..." -ForegroundColor Yellow
dotnet nuget locals all --clear
Write-Host "  Кэш очищен" -ForegroundColor Green

# 3. Восстановление пакетов
Write-Host "`nШаг 3: Восстановление пакетов..." -ForegroundColor Yellow
dotnet restore "$rootPath\DocToScan.csproj"
Write-Host "  Пакеты восстановлены" -ForegroundColor Green

# 4. Сборка проекта
Write-Host "`nШаг 4: Сборка проекта..." -ForegroundColor Yellow
dotnet build "$rootPath\DocToScan.csproj" -c Debug -p:Platform=x64
Write-Host "  Проект собран" -ForegroundColor Green

# 5. Проверка наличия config.xml
Write-Host "`nШаг 5: Проверка config.xml..." -ForegroundColor Yellow
if (-not (Test-Path "$rootPath\config.xml")) {
    @"
<?xml version="1.0" encoding="utf-8"?>
<Configuration>
  <Brightness>
    <Enable>true</Enable>
    <Level>15</Level>
  </Brightness>
  <Rotation>
    <Enable>true</Enable>
    <MinAngle>-3</MinAngle>
    <MaxAngle>3</MaxAngle>
  </Rotation>
  <Grayscale>
    <Enable>true</Enable>
  </Grayscale>
  <ImageQuality>
    <Dpi>150</Dpi>
    <JpegCompression>85</JpegCompression>
  </ImageQuality>
</Configuration>
"@ | Out-File "$rootPath\config.xml" -Encoding UTF8
    Write-Host "  config.xml создан" -ForegroundColor Green
} else {
    Write-Host "  config.xml уже существует" -ForegroundColor Green
}

# 6. Проверка launchSettings.json
Write-Host "`nШаг 6: Проверка launchSettings.json..." -ForegroundColor Yellow
$launchPath = "$rootPath\Properties\launchSettings.json"
if (-not (Test-Path $launchPath)) {
    if (-not (Test-Path "$rootPath\Properties")) {
        New-Item -ItemType Directory -Path "$rootPath\Properties" -Force | Out-Null
    }
    
    @"
{
  "profiles": {
    "DocToScan": {
      "commandName": "Project",
      "commandLineArgs": "",
      "workingDirectory": "$(Convert-Path $rootPath)"
    }
  }
}
"@ | Out-File $launchPath -Encoding UTF8
    Write-Host "  launchSettings.json создан" -ForegroundColor Green
} else {
    Write-Host "  launchSettings.json уже существует" -ForegroundColor Green
}

# 7. Копирование сборки в ожидаемое место
Write-Host "`nШаг 7: Копирование сборки в ожидаемое место..." -ForegroundColor Yellow
$sourceExe = "$rootPath\bin\x64\Debug\net10.0\DocToScan.exe"
$targetExe = "$rootPath\bin\Debug\net10.0\DocToScan.exe"

if (Test-Path $sourceExe) {
    $targetDir = Split-Path $targetExe -Parent
    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }
    Copy-Item $sourceExe $targetExe -Force
    Write-Host "  Сборка скопирована в $targetExe" -ForegroundColor Green
} else {
    Write-Warning "  Исходная сборка не найдена: $sourceExe"
}

Write-Host "`n==================================================" -ForegroundColor Cyan
Write-Host "  ИСПРАВЛЕНИЕ ЗАВЕРШЕНО!" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Теперь попробуйте:" -ForegroundColor Yellow
Write-Host "  1. Открыть решение в Visual Studio" -ForegroundColor White
Write-Host "  2. Выбрать конфигурацию: Debug | x64" -ForegroundColor White
Write-Host "  3. Нажать F5 для запуска" -ForegroundColor White
Write-Host ""
Read-Host "Нажмите Enter для выхода"
'@

Create-File -filename "fix-all.ps1" -content $fixAllContent

# Создаем check-office.ps1
$checkOfficeContent = @'
#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Проверка установленной версии Microsoft Office
.DESCRIPTION
    Диагностический скрипт для определения версии и архитектуры Office
#>

# Устанавливаем кодировку UTF-8 для консоли
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [System.Text.UTF8Encoding]::new()

# Функции для цветного вывода
function Write-Title { Write-Host $args -ForegroundColor Cyan }
function Write-Success { Write-Host $args -ForegroundColor Green }
function Write-Warning { Write-Host $args -ForegroundColor Yellow }
function Write-Error { Write-Host $args -ForegroundColor Red }
function Write-Info { Write-Host $args -ForegroundColor Gray }

# Очистка экрана
Clear-Host

Write-Title "=================================================="
Write-Title "  DocToScan - Диагностика Microsoft Office"
Write-Title "=================================================="
Write-Info "Дата проверки: $(Get-Date -Format 'dd.MM.yyyy HH:mm')"
Write-Info ""

# Функция для получения версии Office из реестра
function Get-OfficeVersionFromRegistry {
    $versions = @{}
    
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
        "HKLM:\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot",
        "HKLM:\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot",
        "HKLM:\SOFTWARE\Microsoft\Office\12.0\Common\InstallRoot",
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot"
    )
    
    Write-Info "Поиск Microsoft Office в реестре..."
    
    foreach ($path in $registryPaths) {
        if (Test-Path $path) {
            try {
                $props = Get-ItemProperty $path -ErrorAction Stop
                if ($props) {
                    $version = if ($path -match "Office\\(\d+\.\d+)") { $matches[1] } else { "Unknown" }
                    
                    if ($props.Path) {
                        $versions[$version] = $props.Path
                        Write-Success "  Найдена версия $version"
                    }
                    
                    if ($props.VersionToReport) {
                        $versions["Microsoft 365"] = $props.VersionToReport
                        Write-Success "  Найдена версия Microsoft 365 ($($props.VersionToReport))"
                    }
                }
            } catch {
                Write-Info "  Пропускаем $path"
            }
        }
    }
    
    return $versions
}

# Проверка через COM
function Test-WordComObject {
    try {
        Write-Info "`nПроверка COM-объекта Word..."
        
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $version = $word.Version
        $build = $word.Build
        
        Write-Success "  COM-объект создан успешно"
        Write-Success "  Версия: $version (Build: $build)"
        
        # Определение понятного названия версии
        $friendlyVersion = switch ($version) {
            "16.0" { 
                if ($build -gt 16000) { "Microsoft 365" } 
                else { "Office 2016/2019/2021/2024" } 
            }
            "15.0" { "Office 2013" }
            "14.0" { "Office 2010" }
            "12.0" { "Office 2007" }
            default { "Неизвестная версия" }
        }
        Write-Success "  Версия (понятная): $friendlyVersion"
        
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
        
        return $true, $version, $build
    } catch {
        Write-Error "  Ошибка создания COM-объекта: $_"
        return $false, $null, $null
    }
}

# Проверка архитектуры
function Test-OfficeArchitecture {
    param([string]$Version)
    
    try {
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $path = $word.Path
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
        
        if ($path -match "Program Files \(x86\)") {
            Write-Warning "  Архитектура: 32-bit"
            return "x86"
        } else {
            Write-Success "  Архитектура: 64-bit"
            return "x64"
        }
    } catch {
        Write-Error "  Не удалось определить архитектуру"
        return "Unknown"
    }
}

# Основная диагностика
Write-Info "Шаг 1: Проверка реестра..."
$versions = Get-OfficeVersionFromRegistry

if ($versions.Count -eq 0) {
    Write-Warning "  Office не найден в реестре"
} else {
    Write-Success "  Найдено версий в реестре: $($versions.Count)"
}

Write-Info "`nШаг 2: Проверка COM-объекта..."
$comResult, $comVersion, $comBuild = Test-WordComObject

if ($comResult) {
    Write-Info "`nШаг 3: Определение архитектуры..."
    $arch = Test-OfficeArchitecture -Version $comVersion
    
    Write-Info "`n"
    Write-Title "=================================================="
    Write-Title "  РЕЗУЛЬТАТЫ ДИАГНОСТИКИ"
    Write-Title "=================================================="
    
    Write-Success "  Microsoft Office УСТАНОВЛЕН"
    Write-Success "  Версия COM: $comVersion (Build: $comBuild)"
    Write-Success "  Архитектура: $arch"
    
    if ($arch -eq "x64") {
        Write-Success "  Оптимально для работы"
    } else {
        Write-Warning "  32-битная версия - может работать медленнее"
    }
    
    Write-Info ""
    Write-Success "  Ваша программа DocToScan будет работать корректно!"
} else {
    Write-Info "`n"
    Write-Title "=================================================="
    Write-Title "  РЕЗУЛЬТАТЫ ДИАГНОСТИКИ"
    Write-Title "=================================================="
    
    Write-Error "  Microsoft Office НЕ НАЙДЕН"
    Write-Error "  Программа не сможет обрабатывать DOCX файлы"
    
    Write-Info ""
    Write-Warning "  Для работы с DOCX файлами требуется:"
    Write-Warning "     - Microsoft Office 2010 или новее"
    Write-Warning "     - Microsoft 365"
    Write-Warning "     - Или альтернатива с поддержкой COM-автоматизации"
}

Write-Info "`n"
Write-Title "=================================================="
Write-Info "  Нажмите Enter для выхода..."
Read-Host
'@

Create-File -filename "check-office.ps1" -content $checkOfficeContent

Write-Host "`nВсе скрипты созданы с правильной кодировкой!" -ForegroundColor Green
Write-Host "Теперь запустите:" -ForegroundColor Yellow
Write-Host "  fix-all.ps1" -ForegroundColor White