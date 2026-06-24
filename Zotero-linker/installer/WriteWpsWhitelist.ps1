param(
    [Parameter(Mandatory = $true)] [string] $Name
)

$ErrorActionPreference = "Stop"

$path = Join-Path "HKCU:\Software\Kingsoft\Office\wps\AddinsWL" $Name
New-Item -Path $path -Force | Out-Null
Write-Output "WPS add-in whitelist (HKCU): $path"
