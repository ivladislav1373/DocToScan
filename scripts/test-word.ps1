# Тестирование доступа к Word Interop

Write-Host "Testing Word Interop access..." -ForegroundColor Cyan

# Пути где может быть сборка
$searchPaths = @(
    [System.IO.Path]::GetDirectoryName((Get-Process -Id $pid).Path),
    "C:\Windows\assembly\GAC_MSIL",
    "C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE16",
    "C:\Program Files (x86)\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\OFFICE16",
    "C:\Program Files\Microsoft Office\Office16",
    "C:\Program Files (x86)\Microsoft Office\Office16",
    "${env:ProgramFiles}\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE16",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\OFFICE16"
)

$found = $false

foreach ($path in $searchPaths) {
    if (Test-Path $path) {
        Write-Host "Searching in: $path" -ForegroundColor Gray
        
        # Ищем Microsoft.Office.Interop.Word.dll
        $files = Get-ChildItem $path -Filter "*.dll" -Recurse -ErrorAction SilentlyContinue | 
                 Where-Object { $_.Name -like "*Interop.Word*" -or $_.Name -like "*office.dll" }
        
        foreach ($file in $files) {
            Write-Host "  Found: $($file.FullName)" -ForegroundColor Green
            $found = $true
        }
    }
}

if (-not $found) {
    Write-Host "Microsoft.Office.Interop.Word not found!" -ForegroundColor Red
    Write-Host "Try installing: Microsoft Office Developer Tools" -ForegroundColor Yellow
}

Read-Host "Press Enter to exit"