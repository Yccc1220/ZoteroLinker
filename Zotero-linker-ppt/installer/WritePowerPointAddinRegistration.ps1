param(
    [Parameter(Mandatory = $true)] [string] $AddInName,
    [Parameter(Mandatory = $true)] [string] $ManifestPath,
    [Parameter(Mandatory = $true)] [string] $FriendlyName,
    [Parameter(Mandatory = $true)] [string] $Description
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $ManifestPath)) {
    Write-Error "Manifest not found: $ManifestPath"
    exit 1
}

$manifestUri = "file:///" + $ManifestPath.Replace('\', '/') + "|vstolocal"
$paths = @(
    "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$AddInName",
    "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Addins\$AddInName"
)

foreach ($path in $paths) {
    New-Item -Path $path -Force | Out-Null
    New-ItemProperty -Path $path -Name "Description" -Value $Description -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $path -Name "FriendlyName" -Value $FriendlyName -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $path -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $path -Name "Manifest" -Value $manifestUri -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $path -Name "CommandLineSafe" -Value 1 -PropertyType DWord -Force | Out-Null
    Write-Output "PowerPoint add-in registration (HKCU): $path"
}
