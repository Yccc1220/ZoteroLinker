param(
    [string] $AssemblyPath = (Join-Path $PSScriptRoot '..\bin\Debug\Zotero-linker.dll')
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

function New-CitationJson {
    param(
        [object[]] $Items
    )

    $citationItems = @()
    foreach ($item in $Items) {
        $citationItems += [ordered] @{
            id = [int] $item.Id
            itemData = [ordered] @{
                id = [int] $item.Id
                title = $item.Title
                author = @(
                    [ordered] @{
                        family = $item.Author
                    }
                )
                issued = [ordered] @{
                    'date-parts' = @(
                        @([int] $item.Year)
                    )
                }
            }
        }
    }

    return ([ordered] @{
        citationItems = $citationItems
    } | ConvertTo-Json -Depth 12 -Compress)
}

function ConvertTo-XmlText {
    param(
        [string] $Value
    )

    return [System.Security.SecurityElement]::Escape($Value)
}

function New-MinimalWordPackage {
    param(
        [string] $Path,
        [string] $CitationCode,
        [string] $CitationText,
        [string] $BibliographyText,
        [string[]] $CitationTextRuns = @()
    )

    $tempRoot = Join-Path $env:TEMP ('zotero-linker-docx-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempRoot | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot '_rels') | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot 'word') | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot 'word\_rels') | Out-Null

    try {
        @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
'@ | Set-Content -LiteralPath (Join-Path $tempRoot '[Content_Types].xml') -Encoding UTF8

        @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@ | Set-Content -LiteralPath (Join-Path $tempRoot '_rels\.rels') -Encoding UTF8

        @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
'@ | Set-Content -LiteralPath (Join-Path $tempRoot 'word\_rels\document.xml.rels') -Encoding UTF8

        $escapedCitationCode = ConvertTo-XmlText $CitationCode
        if ($CitationTextRuns.Count -eq 0) {
            $CitationTextRuns = @($CitationText)
        }

        $citationResultRuns = ($CitationTextRuns | ForEach-Object {
            '<w:r><w:t>' + (ConvertTo-XmlText $_) + '</w:t></w:r>'
        }) -join ''
        $bibliographyParagraphs = ($BibliographyText -split "`r?`n|`r") | Where-Object { $_ -ne '' }
        $bibliographyRuns = ($bibliographyParagraphs | ForEach-Object {
            '<w:p><w:r><w:t>' + (ConvertTo-XmlText $_) + '</w:t></w:r></w:p>'
        }) -join "`n"

        $documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>Citation </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> $escapedCitationCode </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>$citationResultRuns<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
<w:p><w:r><w:t>References</w:t></w:r></w:p>
<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_BIBL </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
$bibliographyRuns
<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>
"@
        Set-Content -LiteralPath (Join-Path $tempRoot 'word\document.xml') -Value $documentXml -Encoding UTF8

        $zipPath = [System.IO.Path]::ChangeExtension($Path, '.zip')
        if (Test-Path -LiteralPath $zipPath) {
            Remove-Item -LiteralPath $zipPath -Force
        }
        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Force
        }

        Compress-Archive -Path (Join-Path $tempRoot '*') -DestinationPath $zipPath
        Move-Item -LiteralPath $zipPath -Destination $Path
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

function New-MinimalWordPackageWithCitations {
    param(
        [string] $Path,
        [object[]] $CitationCases,
        [string] $BibliographyText
    )

    $tempRoot = Join-Path $env:TEMP ('zotero-linker-docx-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempRoot | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot '_rels') | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot 'word') | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tempRoot 'word\_rels') | Out-Null

    try {
        @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
'@ | Set-Content -LiteralPath (Join-Path $tempRoot '[Content_Types].xml') -Encoding UTF8

        @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@ | Set-Content -LiteralPath (Join-Path $tempRoot '_rels\.rels') -Encoding UTF8

        @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
'@ | Set-Content -LiteralPath (Join-Path $tempRoot 'word\_rels\document.xml.rels') -Encoding UTF8

        $citationRuns = @()
        for ($index = 0; $index -lt $CitationCases.Count; $index += 1) {
            $case = $CitationCases[$index]
            $escapedCitationCode = ConvertTo-XmlText $case.Code
            $escapedCitationText = ConvertTo-XmlText $case.Text
            $citationNumber = $index + 1
            $citationRuns += "<w:p><w:r><w:t>Citation $citationNumber </w:t></w:r><w:r><w:fldChar w:fldCharType=""begin""/></w:r><w:r><w:instrText xml:space=""preserve""> $escapedCitationCode </w:instrText></w:r><w:r><w:fldChar w:fldCharType=""separate""/></w:r><w:r><w:t>$escapedCitationText</w:t></w:r><w:r><w:fldChar w:fldCharType=""end""/></w:r></w:p>"
        }

        $bibliographyParagraphs = ($BibliographyText -split "`r?`n|`r") | Where-Object { $_ -ne '' }
        $bibliographyRuns = ($bibliographyParagraphs | ForEach-Object {
            '<w:p><w:r><w:t>' + (ConvertTo-XmlText $_) + '</w:t></w:r></w:p>'
        }) -join "`n"

        $documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
$($citationRuns -join "`n")
<w:p><w:r><w:t>References</w:t></w:r></w:p>
<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_BIBL </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
$bibliographyRuns
<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>
"@
        Set-Content -LiteralPath (Join-Path $tempRoot 'word\document.xml') -Value $documentXml -Encoding UTF8

        $zipPath = [System.IO.Path]::ChangeExtension($Path, '.zip')
        if (Test-Path -LiteralPath $zipPath) {
            Remove-Item -LiteralPath $zipPath -Force
        }
        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Force
        }

        Compress-Archive -Path (Join-Path $tempRoot '*') -DestinationPath $zipPath
        Move-Item -LiteralPath $zipPath -Destination $Path
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
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

function Get-InternalValue {
    param(
        [object] $Object,
        [string] $PropertyName
    )

    return $Object.GetType().GetProperty(
        $PropertyName,
        [System.Reflection.BindingFlags] 'NonPublic, Instance').GetValue($Object, $null)
}

function New-TestItem {
    param(
        [int] $Id,
        [string] $Title,
        [string] $Author,
        [int] $Year
    )

    return [pscustomobject] @{
        Id = $Id
        Title = $Title
        Author = $Author
        Year = $Year
    }
}

function New-BibliographyText {
    param(
        [object[]] $Items
    )

    $lines = @()
    for ($index = 0; $index -lt $Items.Count; $index += 1) {
        $item = $Items[$index]
        $number = $index + 1
        $lines += "[$number]→ $($item.Author). $($item.Title). $($item.Year)."
    }

    return ($lines -join "`r")
}

function Invoke-WordIntegrationCase {
    param(
        [string] $Name,
        [string] $CitationText,
        [object[]] $Items,
        [int] $ExpectedLinked,
        [int] $ExpectedBacklinks,
        [string[]] $ExpectedCitationTexts = @(),
        [int[]] $ExpectedCitationItemIndexes = @(),
        [string[]] $CitationTextRuns = @()
    )

    $word = $null
    $document = $null
    $docxPath = Join-Path $env:TEMP ('zotero-linker-integration-' + [guid]::NewGuid().ToString('N') + '.docx')
    try {
        $citationCode = 'ADDIN ZOTERO_ITEM CSL_CITATION ' + (New-CitationJson -Items $Items)
        New-MinimalWordPackage `
            -Path $docxPath `
            -CitationCode $citationCode `
            -CitationText $CitationText `
            -BibliographyText (New-BibliographyText -Items $Items) `
            -CitationTextRuns $CitationTextRuns

        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $document = $word.Documents.Open($docxPath)

        $citationField = $document.Fields.Item(1)
        $bibliographyField = $document.Fields.Item(2)

        $scanBefore = $scanMethod.Invoke($service, [object[]] @($document))
        Assert-Equal (Get-InternalValue $scanBefore 'CitationFields') 1 "$Name scan before link citation count mismatch."
        Assert-Equal (Get-InternalValue $scanBefore 'BibliographyFields') 1 "$Name scan before link bibliography count mismatch."

        $red = [System.Drawing.Color]::FromArgb(255, 0, 0)
        $linkResult = $linkMethod.Invoke($service, [object[]] @($document, $red, [single] 10))
        $linked = Get-InternalValue $linkResult 'Linked'
        $linkedBacklinks = Get-InternalValue $linkResult 'LinkedBacklinks'
        $failedBibliographyMatch = Get-InternalValue $linkResult 'FailedBibliographyMatch'
        $failedCitationRange = Get-InternalValue $linkResult 'FailedCitationRange'
        $skippedMultiItem = Get-InternalValue $linkResult 'SkippedMultiItem'
        $skippedMissingBibliography = Get-InternalValue $linkResult 'SkippedMissingBibliography'
        if ($linked -ne $ExpectedLinked -or $linkedBacklinks -ne $ExpectedBacklinks -or $failedBibliographyMatch -ne 0 -or $failedCitationRange -ne 0) {
            Write-Host "DEBUG [$Name] citation result text=[$($citationField.Result.Text)]"
            Write-Host "DEBUG [$Name] bibliography result text=[$($bibliographyField.Result.Text)]"
            Write-Host "DEBUG [$Name] linked=$linked backlinks=$linkedBacklinks failedBib=$failedBibliographyMatch failedRange=$failedCitationRange skippedMulti=$skippedMultiItem missingBib=$skippedMissingBibliography"
        }

        Assert-Equal $linked $ExpectedLinked "$Name linked visible citation count mismatch."
        Assert-Equal $linkedBacklinks $ExpectedBacklinks "$Name linked bibliography backlink count mismatch."
        Assert-Equal $failedBibliographyMatch 0 "$Name bibliography match failure count mismatch."
        Assert-Equal $failedCitationRange 0 "$Name citation range failure count mismatch."

        $citeHyperlinks = 0
        $backlinks = 0
        $hyperlinkDebug = @()
        $actualCitationTexts = @()
        $actualCitationTargets = @()
        $actualBacklinkTexts = @()
        $firstBacklink = $null
        foreach ($hyperlink in $document.Hyperlinks) {
            $subAddress = [string] $hyperlink.SubAddress
            $hyperlinkDebug += "sub=[$subAddress] text=[$($hyperlink.Range.Text)]"
            if (-not $document.Bookmarks.Exists($subAddress)) {
                throw "$Name hyperlink target bookmark is missing. Text=[$($hyperlink.Range.Text)] target=[$subAddress]."
            }

            if ($subAddress.StartsWith('Cite_Back_', [System.StringComparison]::OrdinalIgnoreCase)) {
                $bookmarkRangeText = [string] $document.Bookmarks[$subAddress].Range.Text
                if ([string]::IsNullOrWhiteSpace($bookmarkRangeText)) {
                    throw "$Name backlink target bookmark has empty text. Backlink text=[$($hyperlink.Range.Text)] target=[$subAddress]."
                }

                if ($null -eq $firstBacklink) {
                    $firstBacklink = $hyperlink
                }

                $actualBacklinkTexts += ([string] $hyperlink.Range.Text)
                $backlinks += 1
            } elseif ($subAddress.StartsWith('Cite_', [System.StringComparison]::OrdinalIgnoreCase)) {
                $citeHyperlinks += 1
                $actualCitationTexts += ([string] $hyperlink.Range.Text)
                $actualCitationTargets += $subAddress
            }
        }

        if ($citeHyperlinks -ne $ExpectedLinked -or $backlinks -ne $ExpectedBacklinks) {
            $hyperlinkDebug | ForEach-Object { Write-Host "DEBUG [$Name] hyperlink $_" }
        }

        Assert-Equal $citeHyperlinks $ExpectedLinked "$Name document citation hyperlink count mismatch."
        Assert-Equal $backlinks $ExpectedBacklinks "$Name document bibliography backlink count mismatch."
        foreach ($backlinkText in $actualBacklinkTexts) {
            if ($backlinkText -notmatch '^\[[0-9]+\]$') {
                throw "$Name bibliography backlink should only cover the reference number. Actual=[$backlinkText]"
            }
        }

        if ($firstBacklink -ne $null) {
            $firstBacklink.Follow()
            $selectionText = [string] $word.Selection.Range.Text
            if ($selectionText -ne $ExpectedCitationTexts[0]) {
                throw "$Name bibliography backlink did not navigate to the first citation text. Expected=[$($ExpectedCitationTexts[0])] Actual=[$selectionText]"
            }
        }

        if ($ExpectedCitationTexts.Count -gt 0) {
            Assert-Equal $actualCitationTexts.Count $ExpectedCitationTexts.Count "$Name visible citation text count mismatch."
            for ($index = 0; $index -lt $ExpectedCitationTexts.Count; $index += 1) {
                Assert-Equal $actualCitationTexts[$index] $ExpectedCitationTexts[$index] "$Name visible citation text mismatch at index $index."
                $itemIndex = if ($ExpectedCitationItemIndexes.Count -gt $index) { $ExpectedCitationItemIndexes[$index] } else { $index }
                $expectedItemId = [string] $Items[$itemIndex].Id
                if (-not $actualCitationTargets[$index].Contains($expectedItemId)) {
                    throw "$Name citation hyperlink target mismatch at index $index. Expected target to contain item id [$expectedItemId], actual target=[$($actualCitationTargets[$index])]."
                }
            }
        }

        $removeResult = $removeMethod.Invoke($service, [object[]] @($document))
        Assert-Equal (Get-InternalValue $removeResult 'LinksRemoved') ($ExpectedLinked + $ExpectedBacklinks) "$Name removed hyperlink count mismatch."

        $remainingCiteHyperlinks = 0
        foreach ($hyperlink in $document.Hyperlinks) {
            $subAddress = [string] $hyperlink.SubAddress
            if ($subAddress.StartsWith('Cite_', [System.StringComparison]::OrdinalIgnoreCase)) {
                $remainingCiteHyperlinks += 1
            }
        }

        Assert-Equal $remainingCiteHyperlinks 0 "$Name remaining citation hyperlinks after remove mismatch."
        Write-Host "PASS $Name"
    }
    finally {
        if ($document -ne $null) {
            $document.Close($false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
        }
        if ($word -ne $null) {
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        if (Test-Path -LiteralPath $docxPath) {
            Remove-Item -LiteralPath $docxPath -Force
        }
    }
}

function Invoke-RepeatedRangeEndpointCase {
    $word = $null
    $document = $null
    $docxPath = Join-Path $env:TEMP ('zotero-linker-integration-' + [guid]::NewGuid().ToString('N') + '.docx')
    $name = 'Word integration repeated endpoint compressed range workflow'
    try {
        $item1 = New-TestItem -Id 6101 -Title 'Repeated alpha title' -Author 'Reed' -Year 2020
        $item2 = New-TestItem -Id 6102 -Title 'Repeated beta title' -Author 'Stone' -Year 2021
        $item3 = New-TestItem -Id 6103 -Title 'Repeated gamma title' -Author 'Turner' -Year 2022
        $item4 = New-TestItem -Id 6104 -Title 'Repeated delta title' -Author 'Young' -Year 2023
        $items = @($item1, $item2, $item3, $item4)

        $firstCode = 'ADDIN ZOTERO_ITEM CSL_CITATION ' + (New-CitationJson -Items @($item4))
        $secondCode = 'ADDIN ZOTERO_ITEM CSL_CITATION ' + (New-CitationJson -Items $items)
        New-MinimalWordPackageWithCitations `
            -Path $docxPath `
            -CitationCases @(
                [pscustomobject] @{ Code = $firstCode; Text = '[4]' },
                [pscustomobject] @{ Code = $secondCode; Text = '[1-4]' }
            ) `
            -BibliographyText (New-BibliographyText -Items $items)

        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $document = $word.Documents.Open($docxPath)

        $red = [System.Drawing.Color]::FromArgb(255, 0, 0)
        $linkResult = $linkMethod.Invoke($service, [object[]] @($document, $red, [single] 10))
        Assert-Equal (Get-InternalValue $linkResult 'FailedBibliographyMatch') 0 "$name bibliography match failure count mismatch."
        Assert-Equal (Get-InternalValue $linkResult 'FailedCitationRange') 0 "$name citation range failure count mismatch."

        $endpointLinks = @()
        foreach ($hyperlink in $document.Hyperlinks) {
            $subAddress = [string] $hyperlink.SubAddress
            if ($subAddress.StartsWith('Cite_', [System.StringComparison]::OrdinalIgnoreCase) -and
                -not $subAddress.StartsWith('Cite_Back_', [System.StringComparison]::OrdinalIgnoreCase) -and
                ([string] $hyperlink.Range.Text) -eq '4') {
                $endpointLinks += $hyperlink
            }
        }

        Assert-Equal $endpointLinks.Count 2 "$name visible citation endpoint link count mismatch."
        Write-Host "PASS $name"
    }
    finally {
        if ($document -ne $null) {
            $document.Close($false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
        }
        if ($word -ne $null) {
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        if (Test-Path -LiteralPath $docxPath) {
            Remove-Item -LiteralPath $docxPath -Force
        }
    }
}

function Invoke-ShortTitleBibliographyDisambiguationCase {
    $word = $null
    $document = $null
    $docxPath = Join-Path $env:TEMP ('zotero-linker-integration-' + [guid]::NewGuid().ToString('N') + '.docx')
    $name = 'Word integration short title bibliography disambiguation workflow'
    try {
        $items = @(
            (New-TestItem -Id 2919 -Title 'Macrophage mediated mesoscale brain mechanical homeostasis mechanically imaged via optical tweezers and brillouin microscopy in vivo' -Author 'So' -Year 2023),
            (New-TestItem -Id 633 -Title 'Brillouin microscopy' -Author 'Kabakova' -Year 2024),
            (New-TestItem -Id 2931 -Title 'Brillouin scattering, density and elastic properties of the lens and cornea of the eye' -Author 'Vaughan' -Year 1980)
        )

        $citationCode = 'ADDIN ZOTERO_ITEM CSL_CITATION ' + (New-CitationJson -Items @($items[1]))
        $bibliographyText = @(
            '[23] So W Y, Johnson B, Gordon P B, et al. Macrophage mediated mesoscale brain mechanical homeostasis mechanically imaged via optical tweezers and brillouin microscopy in vivo[M]. Biophysics, 2023.',
            '[24] Kabakova I, Zhang J, Xiang Y, et al. Brillouin microscopy[J]. Nature Reviews Methods Primers, 2024, 4(1): 8.',
            '[25] Vaughan J M, Randall J T. Brillouin scattering, density and elastic properties of the lens and cornea of the eye[J]. Nature, 1980, 284(5755): 489-491.'
        ) -join "`r"

        New-MinimalWordPackage `
            -Path $docxPath `
            -CitationCode $citationCode `
            -CitationText '[24]' `
            -BibliographyText $bibliographyText

        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $document = $word.Documents.Open($docxPath)

        $red = [System.Drawing.Color]::FromArgb(255, 0, 0)
        $linkResult = $linkMethod.Invoke($service, [object[]] @($document, $red, [single] 10))
        Assert-Equal (Get-InternalValue $linkResult 'Linked') 1 "$name linked visible citation count mismatch."
        Assert-Equal (Get-InternalValue $linkResult 'LinkedBacklinks') 1 "$name linked bibliography backlink count mismatch."
        Assert-Equal (Get-InternalValue $linkResult 'FailedBibliographyMatch') 0 "$name bibliography match failure count mismatch."
        Assert-Equal (Get-InternalValue $linkResult 'FailedCitationRange') 0 "$name citation range failure count mismatch."

        $citationLink = $null
        $backlink = $null
        foreach ($hyperlink in $document.Hyperlinks) {
            $subAddress = [string] $hyperlink.SubAddress
            if ($subAddress.StartsWith('Cite_Back_', [System.StringComparison]::OrdinalIgnoreCase)) {
                $backlink = $hyperlink
            } elseif ($subAddress.StartsWith('Cite_', [System.StringComparison]::OrdinalIgnoreCase)) {
                $citationLink = $hyperlink
            }
        }

        if ($null -eq $citationLink) {
            throw "$name citation hyperlink was not created."
        }

        $targetText = [string] $document.Bookmarks[[string] $citationLink.SubAddress].Range.Text
        if (-not $targetText.StartsWith('[24]', [System.StringComparison]::Ordinal)) {
            throw "$name citation [24] target mismatch. Expected bibliography entry [24], actual=[$targetText]"
        }

        if ($null -eq $backlink) {
            throw "$name bibliography backlink was not created."
        }

        Assert-Equal ([string] $backlink.Range.Text) '[24]' "$name bibliography backlink text mismatch."
        $backlink.Follow()
        Assert-Equal ([string] $word.Selection.Range.Text) '24' "$name bibliography backlink did not return to body citation."
        Write-Host "PASS $name"
    }
    finally {
        if ($document -ne $null) {
            $document.Close($false)
        }
        if ($word -ne $null) {
            $word.Quit()
        }
        if (Test-Path -LiteralPath $docxPath) {
            Remove-Item -LiteralPath $docxPath -Force
        }
    }
}

$assembly = [System.Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath))
$serviceType = $assembly.GetType('Zotero_linker.ZoteroLinkerService', $true)
$service = [Activator]::CreateInstance($serviceType, $true)
$linkMethod = $serviceType.GetMethod('LinkCitations', [System.Reflection.BindingFlags] 'NonPublic, Instance')
$removeMethod = $serviceType.GetMethod('RemoveCitationLinks', [System.Reflection.BindingFlags] 'NonPublic, Instance')
$scanMethod = $serviceType.GetMethod('ScanDocument', [System.Reflection.BindingFlags] 'NonPublic, Instance')

if ($null -eq $linkMethod -or $null -eq $removeMethod -or $null -eq $scanMethod) {
    throw 'Required ZoteroLinkerService methods were not found. Build Debug configuration first.'
}

Invoke-WordIntegrationCase `
    -Name 'Word integration numeric list link/remove workflow' `
    -CitationText '[1, 2]' `
    -Items @(
        (New-TestItem -Id 1001 -Title 'Alpha integration title' -Author 'Smith' -Year 2020),
        (New-TestItem -Id 1002 -Title 'Beta integration title' -Author 'Jones' -Year 2021)
    ) `
    -ExpectedLinked 2 `
    -ExpectedBacklinks 2 `
    -ExpectedCitationTexts @('1', '2') `
    -ExpectedCitationItemIndexes @(0, 1)

Invoke-WordIntegrationCase `
    -Name 'Word integration trailing compressed range workflow' `
    -CitationText '[1, 2-4]' `
    -Items @(
        (New-TestItem -Id 2001 -Title 'Alpha compressed title' -Author 'Smith' -Year 2020),
        (New-TestItem -Id 2002 -Title 'Beta compressed title' -Author 'Jones' -Year 2021),
        (New-TestItem -Id 2003 -Title 'Gamma compressed title' -Author 'Brown' -Year 2022),
        (New-TestItem -Id 2004 -Title 'Delta compressed title' -Author 'Miller' -Year 2023)
    ) `
    -ExpectedLinked 3 `
    -ExpectedBacklinks 4 `
    -ExpectedCitationTexts @('1', '2', '4') `
    -ExpectedCitationItemIndexes @(0, 1, 3)

Invoke-WordIntegrationCase `
    -Name 'Word integration whole compressed range workflow' `
    -CitationText '[1-4]' `
    -Items @(
        (New-TestItem -Id 3001 -Title 'Alpha whole range title' -Author 'Adams' -Year 2020),
        (New-TestItem -Id 3002 -Title 'Beta whole range title' -Author 'Baker' -Year 2021),
        (New-TestItem -Id 3003 -Title 'Gamma whole range title' -Author 'Clark' -Year 2022),
        (New-TestItem -Id 3004 -Title 'Delta whole range title' -Author 'Davis' -Year 2023)
    ) `
    -ExpectedLinked 2 `
    -ExpectedBacklinks 4 `
    -ExpectedCitationTexts @('1', '4') `
    -ExpectedCitationItemIndexes @(0, 3)

Invoke-WordIntegrationCase `
    -Name 'Word integration en dash compressed range workflow' `
    -CitationText ([string]::Concat('[1', ([char]0x2013).ToString(), '4]')) `
    -Items @(
        (New-TestItem -Id 3101 -Title 'Alpha en dash title' -Author 'Evans' -Year 2020),
        (New-TestItem -Id 3102 -Title 'Beta en dash title' -Author 'Fisher' -Year 2021),
        (New-TestItem -Id 3103 -Title 'Gamma en dash title' -Author 'Green' -Year 2022),
        (New-TestItem -Id 3104 -Title 'Delta en dash title' -Author 'Hall' -Year 2023)
    ) `
    -ExpectedLinked 2 `
    -ExpectedBacklinks 4 `
    -ExpectedCitationTexts @('1', '4') `
    -ExpectedCitationItemIndexes @(0, 3)

Invoke-WordIntegrationCase `
    -Name 'Word integration high-number trailing compressed range workflow' `
    -CitationText '[38, 41-43]' `
    -Items @(
        (New-TestItem -Id 4101 -Title 'Alpha high number title' -Author 'Irwin' -Year 2020),
        (New-TestItem -Id 4102 -Title 'Beta high number title' -Author 'Klein' -Year 2021),
        (New-TestItem -Id 4103 -Title 'Gamma high number title' -Author 'Lopez' -Year 2022),
        (New-TestItem -Id 4104 -Title 'Delta high number title' -Author 'Moore' -Year 2023)
    ) `
    -ExpectedLinked 3 `
    -ExpectedBacklinks 4 `
    -ExpectedCitationTexts @('38', '41', '43') `
    -ExpectedCitationItemIndexes @(0, 1, 3)

Invoke-WordIntegrationCase `
    -Name 'Word integration fullwidth compressed range workflow' `
    -CitationText ([string]::Concat(([char]0xFF3B).ToString(), '1', ([char]0xFF0C).ToString(), '2', ([char]0x2013).ToString(), '4', ([char]0xFF3D).ToString())) `
    -Items @(
        (New-TestItem -Id 5101 -Title 'Alpha fullwidth title' -Author 'Nash' -Year 2020),
        (New-TestItem -Id 5102 -Title 'Beta fullwidth title' -Author 'Owens' -Year 2021),
        (New-TestItem -Id 5103 -Title 'Gamma fullwidth title' -Author 'Perez' -Year 2022),
        (New-TestItem -Id 5104 -Title 'Delta fullwidth title' -Author 'Quinn' -Year 2023)
    ) `
    -ExpectedLinked 3 `
    -ExpectedBacklinks 4 `
    -ExpectedCitationTexts @('1', '2', '4') `
    -ExpectedCitationItemIndexes @(0, 1, 3)

Invoke-WordIntegrationCase `
    -Name 'Word integration split-run compressed range workflow' `
    -CitationText '[1, 2-4]' `
    -CitationTextRuns @('[1, 2', '-', '4]') `
    -Items @(
        (New-TestItem -Id 5201 -Title 'Alpha split run title' -Author 'Rice' -Year 2020),
        (New-TestItem -Id 5202 -Title 'Beta split run title' -Author 'Sanders' -Year 2021),
        (New-TestItem -Id 5203 -Title 'Gamma split run title' -Author 'Thomas' -Year 2022),
        (New-TestItem -Id 5204 -Title 'Delta split run title' -Author 'Underwood' -Year 2023)
    ) `
    -ExpectedLinked 3 `
    -ExpectedBacklinks 4 `
    -ExpectedCitationTexts @('1', '2', '4') `
    -ExpectedCitationItemIndexes @(0, 1, 3)

Invoke-WordIntegrationCase `
    -Name 'Word integration bracketed endpoint compressed range workflow' `
    -CitationText '[1]-[4]' `
    -Items @(
        (New-TestItem -Id 5301 -Title 'Alpha bracketed endpoint title' -Author 'Vance' -Year 2020),
        (New-TestItem -Id 5302 -Title 'Beta bracketed endpoint title' -Author 'White' -Year 2021),
        (New-TestItem -Id 5303 -Title 'Gamma bracketed endpoint title' -Author 'Xavier' -Year 2022),
        (New-TestItem -Id 5304 -Title 'Delta bracketed endpoint title' -Author 'Zimmer' -Year 2023)
    ) `
    -ExpectedLinked 2 `
    -ExpectedBacklinks 4 `
    -ExpectedCitationTexts @('1', '4') `
    -ExpectedCitationItemIndexes @(0, 3)

Invoke-WordIntegrationCase `
    -Name 'Word integration minus-sign compressed range workflow' `
    -CitationText ([string]::Concat('[1', ([char]0x2212).ToString(), '4]')) `
    -Items @(
        (New-TestItem -Id 5401 -Title 'Alpha minus sign title' -Author 'Avery' -Year 2020),
        (New-TestItem -Id 5402 -Title 'Beta minus sign title' -Author 'Blair' -Year 2021),
        (New-TestItem -Id 5403 -Title 'Gamma minus sign title' -Author 'Casey' -Year 2022),
        (New-TestItem -Id 5404 -Title 'Delta minus sign title' -Author 'Drew' -Year 2023)
    ) `
    -ExpectedLinked 2 `
    -ExpectedBacklinks 4 `
    -ExpectedCitationTexts @('1', '4') `
    -ExpectedCitationItemIndexes @(0, 3)

Invoke-RepeatedRangeEndpointCase
Invoke-ShortTitleBibliographyDisambiguationCase
