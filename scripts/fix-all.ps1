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