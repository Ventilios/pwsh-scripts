<#
.SYNOPSIS
    Wrapper script to ensure Get-PowerBiTabularObjects.ps1 runs with PowerShell 7

.DESCRIPTION
    This script checks if it's running in PowerShell 7+ and if not, relaunches itself
    using PowerShell 7 (pwsh.exe) to ensure compatibility with required modules.

.EXAMPLE
    .\Run-With-PowerShell7.ps1 -WorkspaceName "Sales Analytics" -DatasetName "Sales Model"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkspaceName,
    
    [Parameter(Mandatory = $true)]
    [string]$DatasetName,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("JSON", "CSV", "None")]
    [string]$ExportFormat = "None",
    
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = $PSScriptRoot,
    
    [Parameter(Mandatory = $false)]
    [int]$Locale = 1033
)

# Check if we're running PowerShell 7+
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    Write-Host "This script requires PowerShell 7+. Attempting to launch with pwsh.exe..." -ForegroundColor Yellow
    
    # Check if pwsh.exe is available
    $pwshPath = Get-Command pwsh -ErrorAction SilentlyContinue
    if (-not $pwshPath) {
        Write-Error "PowerShell 7+ (pwsh.exe) is not installed or not in PATH."
        Write-Host "Please install PowerShell 7+ from: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Red
        exit 1
    }
    
    # Build the command arguments
    $scriptPath = Join-Path $PSScriptRoot "Get-PowerBiTabularObjects.ps1"
    $arguments = @(
        "-File", "`"$scriptPath`""
        "-WorkspaceName", "`"$WorkspaceName`""
        "-DatasetName", "`"$DatasetName`""
        "-ExportFormat", "`"$ExportFormat`""
        "-ExportPath", "`"$ExportPath`""
        "-Locale", $Locale
    )
    
    Write-Host "Launching: pwsh.exe $($arguments -join ' ')" -ForegroundColor Green
    
    # Launch with PowerShell 7
    & pwsh.exe @arguments
    exit $LASTEXITCODE
}

# If we're already in PowerShell 7+, just run the main script
Write-Host "Running in PowerShell $($PSVersionTable.PSVersion) - proceeding..." -ForegroundColor Green
$scriptPath = Join-Path $PSScriptRoot "Get-PowerBiTabularObjects.ps1"
& $scriptPath -WorkspaceName $WorkspaceName -DatasetName $DatasetName -ExportFormat $ExportFormat -ExportPath $ExportPath -Locale $Locale