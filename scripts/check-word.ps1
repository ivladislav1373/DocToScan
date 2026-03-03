#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Проверка наличия и доступности Microsoft Word
.DESCRIPTION
    Скрипт проверяет установлен ли Microsoft Word, его версию и доступность для COM-взаимодействия
.PARAMETER Quiet
    Тихий режим - только код возврата, без вывода сообщений
.PARAMETER Version
    Показать только версию Word
#>

param(
    [switch]$Quiet,
    [switch]$Version
)

# Устанавливаем кодировку UTF-8
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [System.Text.UTF8Encoding]::new()

# Цвета для вывода
$Host.UI.RawUI.ForegroundColor = "White"
$script:successColor = "Green"
$script:errorColor = "Red"
$script:infoColor = "Yellow"
$script:detailColor = "Cyan"

function Write-Success { if (-not $Quiet) { Write-Host $args -ForegroundColor $script:successColor } }
function Write-Error { if (-not $Quiet) { Write-Host $args -ForegroundColor $script:errorColor } }
function Write-Info { if (-not $Quiet) { Write-Host $args -ForegroundColor $script:infoColor } }
function Write-Detail { if (-not $Quiet) { Write-Host $args -ForegroundColor $script:detailColor } }

# Функция для проверки Word через реестр
function Test-WordRegistry {
    $wordPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE",
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
    )

    foreach ($path in $wordPaths) {
        if (Test-Path $path) {
            $values = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
            if ($values -and $values.'(default)') {
                $exePath = $values.'(default)'
                if (Test-Path $exePath) {
                    return @{
                        Found = $true
                        Path = $exePath
                        Source = "Registry"
                    }
                }
            }
        }
    }
    return $null
}

# Функция для проверки Word через COM
function Test-WordCom {
    try {
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $version = $word.Version
        $visible = $word.Visible
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        
        return @{
            Found = $true
            Version = $version
            Source = "COM"
        }
    } catch {
        return $null
    }
}

# Функция для получения информации о версии Word
function Get-WordVersionInfo {
    param([string]$VersionCode)
    
    $versions = @{
        "16.0" = "Microsoft Office 2016/2019/2021/365"
        "15.0" = "Microsoft Office 2013"
        "14.0" = "Microsoft Office 2010"
        "12.0" = "Microsoft Office 2007"
        "11.0" = "Microsoft Office 2003"
    }
    
    foreach ($key in $versions.Keys) {
        if ($VersionCode -like "$key*") {
            return $versions[$key]
        }
    }
    return "Неизвестная версия"
}

# Основная логика
if (-not $Quiet) {
    Write-Info ("=" * 60)
    Write-Info "DocToScan - Проверка Microsoft Word"
    Write-Info ("=" * 60)
}

# Проверка через реестр
$regInfo = Test-WordRegistry
$comInfo = Test-WordCom

# Определяем результат
$wordFound = ($regInfo -ne $null -or $comInfo -ne $null)
$wordPath = if ($regInfo) { $regInfo.Path } else { "Не найден" }
$wordVersion = if ($comInfo) { $comInfo.Version } else { "Не определен" }
$wordVersionName = Get-WordVersionInfo -VersionCode $wordVersion

# Если запрошена только версия
if ($Version -and $wordFound) {
    Write-Host $wordVersion
    exit 0
} elseif ($Version) {
    Write-Host "0.0"
    exit 1
}

# Вывод информации
if (-not $Quiet) {
    Write-Detail "`nРезультаты проверки:"
    Write-Detail "-" * 40
    
    if ($wordFound) {
        Write-Success "✓ Microsoft Word найден"
        
        if ($wordVersion -ne "Не определен") {
            Write-Detail "  Версия:     $wordVersion ($wordVersionName)"
        }
        
        if ($regInfo) {
            Write-Detail "  Путь:       $wordPath"
            Write-Detail "  Источник:   Реестр Windows"
        }
        
        if ($comInfo) {
            Write-Detail "  COM-доступ: Доступен"
        }
        
        # Проверка разрядности
        if ($wordPath -like "*Program Files (x86)*") {
            Write-Detail "  Разрядность: 32-битная"
        } elseif ($wordPath -like "*Program Files*") {
            Write-Detail "  Разрядность: 64-битная"
        }
        
        # Проверка минимальных требований
        if ($comInfo -and $comInfo.Version -ge "12.0") {
            Write-Success "`n✓ Word совместим с DocToScan (версия $($comInfo.Version) >= 12.0)"
        } elseif ($comInfo) {
            Write-Error "`n✗ Word версии $($comInfo.Version) может быть несовместим. Требуется версия 12.0 или выше"
        }
    } else {
        Write-Error "✗ Microsoft Word не найден!"
        Write-Info "`nDocToScan требует Microsoft Word для конвертации DOCX файлов."
        Write-Info "Возможные решения:"
        Write-Info "  1. Установите Microsoft Word"
        Write-Info "  2. Используйте только PDF файлы"
        Write-Info "  3. Убедитесь, что Word установлен и доступен"
    }
    
    Write-Detail "-" * 40
}

# Возвращаем код результата
if ($wordFound) {
    if ($comInfo -and $comInfo.Version -ge "12.0") {
        exit 0  # Всё отлично
    } elseif ($comInfo) {
        exit 2  # Word найден, но версия старая
    } else {
        exit 1  # Word найден, но COM недоступен
    }
} else {
    exit 3  # Word не найден
}