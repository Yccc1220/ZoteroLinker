param(
    [string] $InstallRoot = "",
    [switch] $RemoveInstallDir
)

$ErrorActionPreference = "Continue"

$addInName = "Zotero-linker"
$displayName = "Zotero Linker"
$wpsWhitelistName = "Yccc1220.Zotero Linker"
$tokens = @("Zotero-linker", "Zotero Linker", "Zotero_linker")

function Remove-KeyIfExists {
    param([string] $Path)
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
        Write-Output "Removed key: $Path"
    }
}

function Remove-ValueIfExists {
    param(
        [string] $Path,
        [string] $Name
    )
    if (Test-Path -LiteralPath $Path) {
        Remove-ItemProperty -LiteralPath $Path -Name $Name -Force -ErrorAction SilentlyContinue
        Write-Output "Removed value: $Path\$Name"
    }
}

function Remove-MatchingSubkeys {
    param(
        [string] $Root,
        [string[]] $Needles
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        return
    }

    Get-ChildItem -LiteralPath $Root -ErrorAction SilentlyContinue | ForEach-Object {
        $remove = $false
        foreach ($needle in $Needles) {
            if ($_.Name -like "*$needle*" -or $_.PSChildName -like "*$needle*") {
                $remove = $true
            }
        }

        if (-not $remove) {
            foreach ($property in @($_.GetValueNames())) {
                $value = [string]$_.GetValue($property)
                foreach ($needle in $Needles) {
                    if ($value -like "*$needle*") {
                        $remove = $true
                    }
                }
            }
        }

        if ($remove) {
            Remove-Item -LiteralPath $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Output "Removed matching key: $($_.Name)"
        }
    }
}

$officeAddinRoots = @(
    "HKLM:\Software\Microsoft\Office\Word\Addins\$addInName",
    "HKLM:\Software\Microsoft\Office\16.0\Word\Addins\$addInName",
    "HKLM:\Software\WOW6432Node\Microsoft\Office\Word\Addins\$addInName",
    "HKLM:\Software\WOW6432Node\Microsoft\Office\16.0\Word\Addins\$addInName",
    "HKCU:\Software\Microsoft\Office\Word\Addins\$addInName",
    "HKCU:\Software\Microsoft\Office\16.0\Word\Addins\$addInName",
    "HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Word\Addins\$addInName",
    "HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Word\Addins\$addInName",
    "HKLM:\Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Word\Addins\$addInName",
    "HKLM:\Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Word\Addins\$addInName",
    "HKCU:\Software\Kingsoft\Office\wps\AddinsWL\$wpsWhitelistName",
    "HKLM:\Software\Kingsoft\Office\wps\AddinsWL\$wpsWhitelistName"
)
foreach ($path in $officeAddinRoots) {
    Remove-KeyIfExists -Path $path
}

$resiliencyRoots = @(
    "HKCU:\Software\Microsoft\Office\Word\Addins\$addInName",
    "HKCU:\Software\Microsoft\Office\16.0\Word\Addins\$addInName",
    "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems",
    "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency\StartupItems"
)
foreach ($path in $resiliencyRoots) {
    Remove-KeyIfExists -Path $path
}

$vstoRoots = @(
    "HKCU:\Software\Microsoft\VSTO\Security\Inclusion",
    "HKLM:\Software\Microsoft\VSTO\Security\Inclusion",
    "HKCU:\Software\Microsoft\VSTO\SolutionMetadata",
    "HKLM:\Software\Microsoft\VSTO\SolutionMetadata",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Deployment\SubscriptionStore",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Deployment\ActivationData"
)
foreach ($root in $vstoRoots) {
    Remove-MatchingSubkeys -Root $root -Needles $tokens
}

$uninstallRoots = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)
foreach ($root in $uninstallRoots) {
    Remove-MatchingSubkeys -Root $root -Needles $tokens
}

if ($InstallRoot -and $RemoveInstallDir -and (Test-Path -LiteralPath $InstallRoot)) {
    $resolved = (Resolve-Path -LiteralPath $InstallRoot).Path
    $allowedRoots = @(
        ${env:ProgramFiles},
        ${env:ProgramFiles(x86)},
        ${env:LOCALAPPDATA}
    ) | Where-Object { $_ }

    $isAllowed = $false
    foreach ($root in $allowedRoots) {
        if ($resolved.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
            $isAllowed = $true
        }
    }

    if ($isAllowed) {
        Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue
        Write-Output "Removed install directory: $resolved"
    }
    else {
        Write-Warning "Install directory was not removed because it is outside expected roots: $resolved"
    }
}
