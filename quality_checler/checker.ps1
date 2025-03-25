$scriptPath = $PSScriptRoot
$logFolder = Join-Path $scriptPath "logs"
$computerName = $env:COMPUTERNAME

# Areas the code checks for existence
$teams = "C:\Program Files (x86)\Teams Installer\Teams.exe"
$word = "C:\Program Files\Microsoft Office\root\Office16"
$excel = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
$dell = "C:\Program Files (x86)\Dell\CommandUpdate\DellCommandUpdate.exe"
$zscaler = "C:\Program Files\Zscaler\ZSATray\ZSATray.exe"
$intune = "C:\Program Files (x86)\Microsoft Intune Management Extension\AgentExecutor.exe"

# Stores the test results
$logFile = Join-Path $logFolder "$computerName`_$(Get-Date -Format 'yyyyMMdd').txt"
$results = @()
$failure = $false

Write-Host "Running system quality check..."
Start-Sleep -Seconds 1

# Teams checker
Write-Host "Checking for Microsoft Teams..."
Start-Sleep -Seconds 1
$results += if (Test-Path $teams) {
    "PASS - Teams"
} else {
    $failure = $true
    "FAIL - Teams"
}

# Office 365 checker
Write-Host "Checking for Office 365 applications..."
Start-Sleep -Seconds 1
$results += if ((Test-Path $word) -and (Test-Path $excel)) {
    "PASS - Office 365 Enterprise"
} else {
    $failure = $true
    "FAIL - Office 365 Enterprise"
}

# Dell Command Update checker
Write-Host "Checking for Dell Command Update..."
Start-Sleep -Seconds 1
$results += if (Test-Path $dell) {
    "PASS - Dell Command Update"
} else {
    $failure = $true
    "FAIL - Dell Command Update"
}

# Zscaler checker
Write-Host "Checking for Zscaler..."
Start-Sleep -Seconds 1
$results += if (Test-Path $zscaler) {
    "PASS - Zscaler"
} else {
    $failure = $true
    "FAIL - Zscaler"
}

# Wi-Fi adapter checker
Write-Host "Checking for Wi-Fi adapter..."
Start-Sleep -Seconds 1
$results += if (Get-NetAdapter -Name "*Wi-Fi*" -Physical) {
    "PASS - Wi-Fi adapter"
} else {
    $failure = $true
    "FAIL - Wi-Fi adapter"
}

# Bluetooth adapter checker
Write-Host "Checking for Bluetooth adapter..."
Start-Sleep -Seconds 1
$results += if (Get-PnpDevice -PresentOnly | Where-Object { $_.Class -eq "Bluetooth" }) {
    "PASS - Bluetooth adapter"
} else {
    $failure = $true
    "FAIL - Bluetooth adapter"
}

# Intune Managemnet Extension checker
Write-Host "Checking for Intune..."
Start-Sleep -Seconds 1
$results += if (Test-Path $intune){
    "PASS - Intune Management"
} else {
    $failure = $true
    "FAIL - Intune Management"
}

Start-Sleep -Seconds 1
Clear-Host

$logContent = @()
$logContent += "System check results - $(Get-Date)"
$logContent += "*--------------------------------------*"


foreach ($result in $results) {
    if ($result -like "PASS*") {
        Write-Host -ForegroundColor Green $result
        $logContent += $result
    } else {
        Write-Host -ForegroundColor Red $result
        $logContent += $result
    }
}
Write-Host "Log file created"

if ($failure) {
    Write-Host ""
    Write-Host -ForegroundColor Red "OVERALL: FAIL"
    Write-Host -ForegroundColor Yellow "Please remedy the failed check(s) and try again."
    $logContent += "WARNING: Some checks failed. Please remedy and retry."
} else {
    Write-Host ""
    Write-Host -ForegroundColor Green "OVERALL: PASS"
}

$logContent | Out-File -FilePath $logFile -Encoding UTF8

Write-Host ""
Write-Host "Log file created: $logFile" -ForegroundColor Cyan
