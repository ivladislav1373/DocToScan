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