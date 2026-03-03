# Этот скрипт должен быть сохранен в ANSI кодировке

$scriptsPath = "G:\Projects\C#\DocToScan\scripts"

# Функция для создания файла с правильной кодировкой
function Create-File {
    param($filename, $content)
    
    $filePath = "$scriptsPath\$filename"
    
    # Сохраняем в UTF-8 без BOM
    [System.IO.File]::WriteAllText($filePath, $content, [System.Text.UTF8Encoding]::new($false))
    
    Write-Host "Created: $filename" -ForegroundColor Green
}

# Создаем fix-all.ps1 - БЕЗ РУССКИХ БУКВ В КОДЕ
$fixAllContent = @'
#!/usr/bin/env pwsh

# Set UTF-8 encoding for console
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [System.Text.UTF8Encoding]::new()

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  DocToScan - Fix All Problems" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

$rootPath = "G:\Projects\C#\DocToScan"

# Step 1: Clean old builds
Write-Host "Step 1: Cleaning old builds..." -ForegroundColor Yellow
if (Test-Path "$rootPath\bin") {
    Remove-Item -Recurse -Force "$rootPath\bin" -ErrorAction SilentlyContinue
    Write-Host "  bin folder deleted" -ForegroundColor Green
}
if (Test-Path "$rootPath\obj") {
    Remove-Item -Recurse -Force "$rootPath\obj" -ErrorAction SilentlyContinue
    Write-Host "  obj folder deleted" -ForegroundColor Green
}

# Step 2: Clean NuGet cache
Write-Host "`nStep 2: Cleaning NuGet cache..." -ForegroundColor Yellow
dotnet nuget locals all --clear
Write-Host "  Cache cleared" -ForegroundColor Green

# Step 3: Restore packages
Write-Host "`nStep 3: Restoring packages..." -ForegroundColor Yellow
dotnet restore "$rootPath\DocToScan.csproj"
Write-Host "  Packages restored" -ForegroundColor Green

# Step 4: Build project
Write-Host "`nStep 4: Building project..." -ForegroundColor Yellow
dotnet build "$rootPath\DocToScan.csproj" -c Debug -p:Platform=x64
Write-Host "  Project built" -ForegroundColor Green

# Step 5: Check config.xml
Write-Host "`nStep 5: Checking config.xml..." -ForegroundColor Yellow
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
    Write-Host "  config.xml created" -ForegroundColor Green
} else {
    Write-Host "  config.xml already exists" -ForegroundColor Green
}

# Step 6: Check launchSettings.json
Write-Host "`nStep 6: Checking launchSettings.json..." -ForegroundColor Yellow
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
    Write-Host "  launchSettings.json created" -ForegroundColor Green
} else {
    Write-Host "  launchSettings.json already exists" -ForegroundColor Green
}

# Step 7: Copy build to expected location
Write-Host "`nStep 7: Copying build to expected location..." -ForegroundColor Yellow
$sourceExe = "$rootPath\bin\x64\Debug\net10.0\DocToScan.exe"
$targetExe = "$rootPath\bin\Debug\net10.0\DocToScan.exe"

if (Test-Path $sourceExe) {
    $targetDir = Split-Path $targetExe -Parent
    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }
    Copy-Item $sourceExe $targetExe -Force
    Write-Host "  Build copied to $targetExe" -ForegroundColor Green
} else {
    Write-Warning "  Source build not found: $sourceExe"
}

Write-Host "`n==================================================" -ForegroundColor Cyan
Write-Host "  FIX COMPLETED!" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Now try:" -ForegroundColor Yellow
Write-Host "  1. Open solution in Visual Studio" -ForegroundColor White
Write-Host "  2. Select configuration: Debug | x64" -ForegroundColor White
Write-Host "  3. Press F5 to run" -ForegroundColor White
Write-Host ""
Read-Host "Press Enter to exit"
'@

Create-File -filename "fix-all.ps1" -content $fixAllContent

# Создаем check-office.ps1 - БЕЗ РУССКИХ БУКВ В КОДЕ
$checkOfficeContent = @'
#!/usr/bin/env pwsh

# Set UTF-8 encoding for console
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [System.Text.UTF8Encoding]::new()

# Color output functions
function Write-Title { Write-Host $args -ForegroundColor Cyan }
function Write-Success { Write-Host $args -ForegroundColor Green }
function Write-Warning { Write-Host $args -ForegroundColor Yellow }
function Write-Error { Write-Host $args -ForegroundColor Red }
function Write-Info { Write-Host $args -ForegroundColor Gray }

# Clear screen
Clear-Host

Write-Title "=================================================="
Write-Title "  DocToScan - Microsoft Office Diagnostics"
Write-Title "=================================================="
Write-Info "Check date: $(Get-Date -Format 'dd.MM.yyyy HH:mm')"
Write-Info ""

# Function to get Office version from registry
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
    
    Write-Info "Searching Microsoft Office in registry..."
    
    foreach ($path in $registryPaths) {
        if (Test-Path $path) {
            try {
                $props = Get-ItemProperty $path -ErrorAction Stop
                if ($props) {
                    $version = if ($path -match "Office\\(\d+\.\d+)") { $matches[1] } else { "Unknown" }
                    
                    if ($props.Path) {
                        $versions[$version] = $props.Path
                        Write-Success "  Found version $version"
                    }
                    
                    if ($props.VersionToReport) {
                        $versions["Microsoft 365"] = $props.VersionToReport
                        Write-Success "  Found version Microsoft 365 ($($props.VersionToReport))"
                    }
                }
            } catch {
                Write-Info "  Skipping $path"
            }
        }
    }
    
    return $versions
}

# Test via COM
function Test-WordComObject {
    try {
        Write-Info "`nTesting COM object Word..."
        
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $version = $word.Version
        $build = $word.Build
        
        Write-Success "  COM object created successfully"
        Write-Success "  Version: $version (Build: $build)"
        
        $friendlyVersion = switch ($version) {
            "16.0" { 
                if ($build -gt 16000) { "Microsoft 365" } 
                else { "Office 2016/2019/2021/2024" } 
            }
            "15.0" { "Office 2013" }
            "14.0" { "Office 2010" }
            "12.0" { "Office 2007" }
            default { "Unknown version" }
        }
        Write-Success "  Friendly version: $friendlyVersion"
        
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
        
        return $true, $version, $build
    } catch {
        Write-Error "  Error creating COM object: $_"
        return $false, $null, $null
    }
}

# Test architecture
function Test-OfficeArchitecture {
    param([string]$Version)
    
    try {
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $path = $word.Path
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
        
        if ($path -match "Program Files \(x86\)") {
            Write-Warning "  Architecture: 32-bit"
            return "x86"
        } else {
            Write-Success "  Architecture: 64-bit"
            return "x64"
        }
    } catch {
        Write-Error "  Could not determine architecture"
        return "Unknown"
    }
}

# Main diagnostics
Write-Info "Step 1: Registry check..."
$versions = Get-OfficeVersionFromRegistry

if ($versions.Count -eq 0) {
    Write-Warning "  Office not found in registry"
} else {
    Write-Success "  Found versions in registry: $($versions.Count)"
}

Write-Info "`nStep 2: COM object test..."
$comResult, $comVersion, $comBuild = Test-WordComObject

if ($comResult) {
    Write-Info "`nStep 3: Architecture check..."
    $arch = Test-OfficeArchitecture -Version $comVersion
    
    Write-Info "`n"
    Write-Title "=================================================="
    Write-Title "  DIAGNOSTICS RESULTS"
    Write-Title "=================================================="
    
    Write-Success "  Microsoft Office INSTALLED"
    Write-Success "  COM Version: $comVersion (Build: $comBuild)"
    Write-Success "  Architecture: $arch"
    
    if ($arch -eq "x64") {
        Write-Success "  Optimal for work"
    } else {
        Write-Warning "  32-bit version - may work slower"
    }
    
    Write-Info ""
    Write-Success "  Your DocToScan program will work correctly!"
} else {
    Write-Info "`n"
    Write-Title "=================================================="
    Write-Title "  DIAGNOSTICS RESULTS"
    Write-Title "=================================================="
    
    Write-Error "  Microsoft Office NOT FOUND"
    Write-Error "  Program will not process DOCX files"
    
    Write-Info ""
    Write-Warning "  To work with DOCX files you need:"
    Write-Warning "     - Microsoft Office 2010 or newer"
    Write-Warning "     - Microsoft 365"
    Write-Warning "     - Alternative with COM automation support"
}

Write-Info "`n"
Write-Title "=================================================="
Write-Info "  Press Enter to exit..."
Read-Host
'@

Create-File -filename "check-office.ps1" -content $checkOfficeContent

Write-Host "`nAll scripts created with correct encoding!" -ForegroundColor Green
Write-Host "Now run:" -ForegroundColor Yellow
Write-Host "  .\fix-all.ps1" -ForegroundColor White