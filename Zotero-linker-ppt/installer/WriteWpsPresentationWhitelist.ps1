param(
    [Parameter(Mandatory = $true)] [string] $Name
)

$ErrorActionPreference = "Stop"

$path = "HKCU:\Software\Kingsoft\Office\WPP\AddinsWL"
New-Item -Path $path -Force | Out-Null
New-ItemProperty -Path $path -Name $Name -Value "" -PropertyType String -Force | Out-Null
Write-Output "WPS Presentation add-in whitelist value (HKCU): $path\$Name"
