param(
    [string] $AssemblyPath = (Join-Path $PSScriptRoot '..\bin\Debug\Zotero-linker.dll')
)

$ErrorActionPreference = 'Stop'

$assembly = [System.Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$serviceType = $assembly.GetType('Zotero_linker.ZoteroLinkerService', $true)
$method = $serviceType.GetMethod(
    'TestExtractCitationSegments',
    [System.Reflection.BindingFlags] 'NonPublic, Static')
$bookmarkMethod = $serviceType.GetMethod(
    'TestBuildBookmarkName',
    [System.Reflection.BindingFlags] 'NonPublic, Static')
$bibliographyAnchorMethod = $serviceType.GetMethod(
    'TestExtractBibliographyAnchorText',
    [System.Reflection.BindingFlags] 'NonPublic, Static')
$legacyBookmarkMethod = $serviceType.GetMethod(
    'TestIsLegacyMacroBookmarkName',
    [System.Reflection.BindingFlags] 'NonPublic, Static')

if ($null -eq $method) {
    throw 'TestExtractCitationSegments was not found. Build Debug configuration first.'
}

if ($null -eq $bookmarkMethod) {
    throw 'TestBuildBookmarkName was not found. Build Debug configuration first.'
}

if ($null -eq $bibliographyAnchorMethod) {
    throw 'TestExtractBibliographyAnchorText was not found. Build Debug configuration first.'
}

if ($null -eq $legacyBookmarkMethod) {
    throw 'TestIsLegacyMacroBookmarkName was not found. Build Debug configuration first.'
}

function Invoke-ExtractSegments {
    param(
        [string] $Text,
        [int] $Count
    )

    $segments = $method.Invoke($null, @($Text, $Count))
    $items = @()
    foreach ($segment in $segments) {
        $segmentType = $segment.GetType()
        $startOffset = [int] $segmentType.GetProperty(
            'StartOffset',
            [System.Reflection.BindingFlags] 'NonPublic, Instance').GetValue($segment, $null)
        $endOffset = [int] $segmentType.GetProperty(
            'EndOffset',
            [System.Reflection.BindingFlags] 'NonPublic, Instance').GetValue($segment, $null)
        $visible = [bool] $segmentType.GetProperty(
            'Visible',
            [System.Reflection.BindingFlags] 'NonPublic, Instance').GetValue($segment, $null)

        $items += [pscustomobject] @{
            Start = $startOffset
            End = $endOffset
            Visible = $visible
            Text = $Text.Substring($startOffset, $endOffset - $startOffset)
        }
    }

    return $items
}

function Assert-Equal {
    param(
        [object] $Actual,
        [object] $Expected,
        [string] $Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected=[$Expected] Actual=[$Actual]"
    }
}

function Assert-Segments {
    param(
        [string] $Name,
        [string] $Text,
        [int] $Count,
        [object[]] $Expected
    )

    $segments = @(Invoke-ExtractSegments -Text $Text -Count $Count)
    Assert-Equal $segments.Count $Expected.Count "$Name segment count mismatch."

    for ($index = 0; $index -lt $Expected.Count; $index += 1) {
        Assert-Equal $segments[$index].Text $Expected[$index].Text "$Name segment[$index] text mismatch."
        Assert-Equal $segments[$index].Visible $Expected[$index].Visible "$Name segment[$index] visible mismatch."
    }

    Write-Host "PASS $Name"
}

function Assert-BookmarkNamesDiffer {
    $titlePrefix = 'A very long shared title prefix that would otherwise be truncated before the differentiating suffix '
    $firstArgs = [object[]] @(
        '100',
        ($titlePrefix + 'alpha'),
        '2020',
        'Smith')
    $secondArgs = [object[]] @(
        '101',
        ($titlePrefix + 'beta'),
        '2021',
        'Smith')
    $first = [string] ($bookmarkMethod.Invoke($null, $firstArgs))
    $second = [string] ($bookmarkMethod.Invoke($null, $secondArgs))

    if ($first -eq $second) {
        throw "bookmark names should not collide. Name=[$first]"
    }

    foreach ($name in @($first, $second)) {
        if ($name.Length -gt 40) {
            throw "bookmark name is too long. Name=[$name] Length=$($name.Length)"
        }
        if ($name -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') {
            throw "bookmark name contains invalid Word bookmark characters. Name=[$name]"
        }
    }

    Write-Host 'PASS bookmark names include collision-resistant hash'
}

function Assert-BibliographyAnchor {
    param(
        [string] $Name,
        [string] $Text,
        [string] $Expected
    )

    $args = [object[]] @($Text)
    $actual = [string] ($bibliographyAnchorMethod.Invoke($null, $args))
    Assert-Equal $actual $Expected "$Name bibliography anchor mismatch."
    Write-Host "PASS $Name"
}

function Assert-LegacyBookmarkName {
    param(
        [string] $Name,
        [string] $BookmarkName,
        [bool] $Expected
    )

    $args = [object[]] @($BookmarkName)
    $actual = [bool] ($legacyBookmarkMethod.Invoke($null, $args))
    Assert-Equal $actual $Expected "$Name legacy bookmark name mismatch."
    Write-Host "PASS $Name"
}

Assert-Segments `
    -Name 'numeric list' `
    -Text '[1, 2, 3]' `
    -Count 3 `
    -Expected @(
        @{ Text = '1'; Visible = $true },
        @{ Text = '2'; Visible = $true },
        @{ Text = '3'; Visible = $true }
    )

Assert-Segments `
    -Name 'numeric range bracketed endpoints' `
    -Text '[1]-[3]' `
    -Count 3 `
    -Expected @(
        @{ Text = '1'; Visible = $true },
        @{ Text = '1'; Visible = $false },
        @{ Text = '3'; Visible = $true }
    )

Assert-Segments `
    -Name 'numeric trailing compressed range' `
    -Text '[1, 2-4]' `
    -Count 4 `
    -Expected @(
        @{ Text = '1'; Visible = $true },
        @{ Text = '2'; Visible = $true },
        @{ Text = '2'; Visible = $false },
        @{ Text = '4'; Visible = $true }
    )

Assert-Segments `
    -Name 'numeric trailing high-number compressed range' `
    -Text '[38, 41-43]' `
    -Count 4 `
    -Expected @(
        @{ Text = '38'; Visible = $true },
        @{ Text = '41'; Visible = $true },
        @{ Text = '41'; Visible = $false },
        @{ Text = '43'; Visible = $true }
    )

Assert-Segments `
    -Name 'numeric fullwidth bracket compressed range' `
    -Text ([string]::Concat(([char]0xFF3B).ToString(), '1', ([char]0xFF0C).ToString(), '2', ([char]0x2013).ToString(), '4', ([char]0xFF3D).ToString())) `
    -Count 4 `
    -Expected @(
        @{ Text = '1'; Visible = $true },
        @{ Text = '2'; Visible = $true },
        @{ Text = '2'; Visible = $false },
        @{ Text = '4'; Visible = $true }
    )

Assert-Segments `
    -Name 'numeric minus-sign compressed range' `
    -Text ([string]::Concat('[1', ([char]0x2212).ToString(), '4]')) `
    -Count 4 `
    -Expected @(
        @{ Text = '1'; Visible = $true },
        @{ Text = '1'; Visible = $false },
        @{ Text = '1'; Visible = $false },
        @{ Text = '4'; Visible = $true }
    )

Assert-Segments `
    -Name 'author year semicolon' `
    -Text '(Smith, 2020; Jones, 2021)' `
    -Count 2 `
    -Expected @(
        @{ Text = 'Smith, 2020'; Visible = $true },
        @{ Text = 'Jones, 2021'; Visible = $true }
    )

Assert-Segments `
    -Name 'author year comma same author' `
    -Text '(Smith, 2020, 2021)' `
    -Count 2 `
    -Expected @(
        @{ Text = 'Smith, 2020'; Visible = $true },
        @{ Text = '2021'; Visible = $true }
    )

Assert-BookmarkNamesDiffer

Assert-BibliographyAnchor `
    -Name 'bibliography halfwidth bracket anchor' `
    -Text '[12] Smith title' `
    -Expected '[12]'

Assert-BibliographyAnchor `
    -Name 'bibliography halfwidth bracket arrow anchor' `
    -Text ([string]::Concat('[12]', ([char]0x2192).ToString(), ' Smith title')) `
    -Expected ([string]::Concat('[12]', ([char]0x2192).ToString()))

Assert-BibliographyAnchor `
    -Name 'bibliography fullwidth bracket anchor' `
    -Text ([string]::Concat(([char]0xFF3B).ToString(), '12', ([char]0xFF3D).ToString(), ' Smith title')) `
    -Expected ([string]::Concat(([char]0xFF3B).ToString(), '12', ([char]0xFF3D).ToString()))

Assert-BibliographyAnchor `
    -Name 'bibliography fullwidth dot anchor' `
    -Text ([string]::Concat('12', ([char]0xFF0E).ToString(), ' Smith title')) `
    -Expected ([string]::Concat('12', ([char]0xFF0E).ToString()))

Assert-LegacyBookmarkName `
    -Name 'legacy macro bookmark title anchor' `
    -BookmarkName 'Long_Title_From_Macro_123' `
    -Expected $true

Assert-LegacyBookmarkName `
    -Name 'legacy macro bookmark rejects numeric prefix' `
    -BookmarkName '123_Title_From_Macro' `
    -Expected $false

Assert-LegacyBookmarkName `
    -Name 'legacy macro bookmark rejects punctuation' `
    -BookmarkName 'Title-From-Macro' `
    -Expected $false

Write-Host 'All citation segment tests passed.'
