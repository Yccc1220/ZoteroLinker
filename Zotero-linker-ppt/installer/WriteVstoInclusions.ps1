param(
    [Parameter(Mandatory = $true)] [string] $ManifestPath,
    [Parameter(Mandatory = $false)] [ValidateSet("HKCU", "HKLM", "Both")] [string] $Target = "Both"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $ManifestPath)) {
    Write-Error "Manifest not found: $ManifestPath"
    exit 1
}

[xml] $manifest = Get-Content -LiteralPath $ManifestPath

$ns = New-Object Xml.XmlNamespaceManager $manifest.NameTable
$ns.AddNamespace("dsig", "http://www.w3.org/2000/09/xmldsig#")

$rsaNode = $manifest.SelectSingleNode("//dsig:Signature/dsig:KeyInfo/dsig:KeyValue/dsig:RSAKeyValue", $ns)
if (-not $rsaNode) {
    Write-Error "Could not find RSAKeyValue in manifest signature"
    exit 1
}

$modulus = $rsaNode.Modulus
$exponent = $rsaNode.Exponent
$publicKey = "<RSAKeyValue><Modulus>$modulus</Modulus><Exponent>$exponent</Exponent></RSAKeyValue>"

# VSTO expects an unescaped file URI and does not match Uri-encoded spaces reliably.
$manifestUri = "file:///" + $ManifestPath.Replace('\', '/')

$urlBytes = [System.Text.Encoding]::UTF8.GetBytes($manifestUri.ToLowerInvariant())
$sha256 = [System.Security.Cryptography.SHA256]::Create()
$hash = $sha256.ComputeHash($urlBytes)
$guidBytes = [byte[]] ($hash[0..15])
$guid = ([Guid]::new($guidBytes)).ToString("D")

if ($Target -in @("HKCU", "Both")) {
    $hkcuPath = "HKCU:\Software\Microsoft\VSTO\Security\Inclusion\$guid"
    New-Item -Path $hkcuPath -Force -ErrorAction SilentlyContinue | Out-Null
    if (Test-Path $hkcuPath) {
        New-ItemProperty -Path $hkcuPath -Name "Url" -Value $manifestUri -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $hkcuPath -Name "PublicKey" -Value $publicKey -PropertyType String -Force | Out-Null
        Write-Output "VSTO inclusion (HKCU): $manifestUri"
    }
    else {
        Write-Warning "Could not create HKCU inclusion: $manifestUri"
    }
}

if ($Target -in @("HKLM", "Both")) {
    $hklmPath = "HKLM:\Software\Microsoft\VSTO\Security\Inclusion\$guid"
    New-Item -Path $hklmPath -Force -ErrorAction SilentlyContinue | Out-Null
    if (Test-Path $hklmPath) {
        New-ItemProperty -Path $hklmPath -Name "Url" -Value $manifestUri -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $hklmPath -Name "PublicKey" -Value $publicKey -PropertyType String -Force | Out-Null
        Write-Output "VSTO inclusion (HKLM): $manifestUri"
    }
    else {
        Write-Warning "Could not create HKLM inclusion: $manifestUri"
    }
}
