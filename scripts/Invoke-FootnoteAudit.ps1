[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputDocx,

    [Parameter(Mandatory)]
    [string]$BibliographyCsv,

    [Parameter(Mandatory)]
    [string]$PolicyJson,

    [Parameter(Mandatory)]
    [string]$OutputDirectory,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-RequiredFile {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label does not exist or is not a file: $Path"
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

function Assert-NoReparsePointInPath {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Label
    )

    $current = [System.IO.Path]::GetFullPath($Path)
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force
            if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "$Label traverses a reparse point: $current"
            }
        }
        $parent = [System.IO.Path]::GetDirectoryName($current)
        if ([string]::IsNullOrWhiteSpace($parent) -or
            [string]::Equals($parent, $current, [System.StringComparison]::OrdinalIgnoreCase)) {
            break
        }
        $current = $parent
    }
}

function Prepare-OutputFile {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$OutputRoot
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullRoot = [System.IO.Path]::GetFullPath($OutputRoot).TrimEnd('\', '/')
    $expectedParent = [System.IO.Path]::GetDirectoryName($fullPath).TrimEnd('\', '/')
    if (-not [string]::Equals($expectedParent, $fullRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to prepare an output outside OutputDirectory: $fullPath"
    }

    Assert-NoReparsePointInPath -Path $fullPath -Label 'Audit output path'
    if (Test-Path -LiteralPath $fullPath) {
        $item = Get-Item -LiteralPath $fullPath -Force
        if ($item.PSIsContainer) {
            throw "Audit output path is a directory, not a file: $fullPath"
        }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Refusing to replace a reparse-point audit output: $fullPath"
        }
        # Remove the directory entry before writing. This prevents writing through an existing hardlink.
        Remove-Item -LiteralPath $fullPath -Force
    }
}

function Normalize-Whitespace {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) {
        return ''
    }

    return (([regex]::Replace($Value, '\s+', ' ')).Trim()).Normalize(
        [System.Text.NormalizationForm]::FormC
    )
}

function Test-AliasMatch {
    param(
        [AllowNull()][string]$Text,
        [AllowNull()][string]$Alias
    )

    if ([string]::IsNullOrWhiteSpace($Text) -or [string]::IsNullOrWhiteSpace($Alias)) {
        return $false
    }

    $Text = $Text.Normalize([System.Text.NormalizationForm]::FormC)
    $Alias = $Alias.Normalize([System.Text.NormalizationForm]::FormC)
    $prefix = if ([regex]::IsMatch($Alias, '^[\p{L}\p{N}\p{M}]')) { '(?<![\p{L}\p{N}\p{M}])' } else { '' }
    $suffix = if ([regex]::IsMatch($Alias, '[\p{L}\p{N}\p{M}]$')) { '(?![\p{L}\p{N}\p{M}])' } else { '' }
    $pattern = $prefix + [regex]::Escape($Alias) + $suffix
    return [regex]::IsMatch(
        $Text,
        $pattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
        [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    )
}

function Protect-SpreadsheetCell {
    param([AllowNull()]$Value)

    if ($Value -is [string] -and $Value -match '^[\x00-\x20]*[=+\-@]') {
        return "'$Value"
    }
    return $Value
}

function Read-ZipEntryText {
    param(
        [Parameter(Mandatory)]
        [System.IO.Compression.ZipArchive]$Archive,

        [Parameter(Mandatory)]
        [string]$EntryName
    )

    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) {
        throw "DOCX package is missing required part: $EntryName"
    }

    $maximumUncompressedBytes = 64MB
    $maximumCompressionRatio = 500.0
    if ($entry.Length -gt $maximumUncompressedBytes) {
        throw "DOCX part exceeds the 64 MiB audit limit: $EntryName"
    }
    if ($entry.CompressedLength -gt 0 -and
        ($entry.Length / [double]$entry.CompressedLength) -gt $maximumCompressionRatio) {
        throw "DOCX part exceeds the permitted compression ratio: $EntryName"
    }

    $stream = $entry.Open()
    $reader = $null
    try {
        $reader = [System.IO.StreamReader]::new(
            $stream,
            [System.Text.Encoding]::UTF8,
            $true
        )
        return $reader.ReadToEnd()
    }
    finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        }
        else {
            $stream.Dispose()
        }
    }
}

function ConvertTo-SafeXmlDocument {
    param(
        [Parameter(Mandatory)]
        [string]$XmlText,

        [Parameter(Mandatory)]
        [string]$Label
    )

    $settings = [System.Xml.XmlReaderSettings]::new()
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $settings.MaxCharactersInDocument = 64MB
    $settings.MaxCharactersFromEntities = 0

    $stringReader = [System.IO.StringReader]::new($XmlText)
    $xmlReader = $null
    try {
        $xmlReader = [System.Xml.XmlReader]::Create($stringReader, $settings)
        $document = [System.Xml.XmlDocument]::new()
        $document.XmlResolver = $null
        $document.Load($xmlReader)
        return ,$document
    }
    catch {
        throw "$Label is not safe, valid XML: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $xmlReader) {
            $xmlReader.Dispose()
        }
        $stringReader.Dispose()
    }
}

function Get-LogicalFootnoteText {
    param(
        [Parameter(Mandatory)]
        [System.Xml.XmlElement]$FootnoteNode,

        [Parameter(Mandatory)]
        [System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    $paragraphTexts = [System.Collections.Generic.List[string]]::new()
    foreach ($paragraph in $FootnoteNode.SelectNodes('.//w:p', $NamespaceManager)) {
        $builder = [System.Text.StringBuilder]::new()
        foreach ($element in $paragraph.SelectNodes('.//*')) {
            switch ($element.LocalName) {
                't' {
                    [void]$builder.Append($element.InnerText)
                }
                'tab' {
                    [void]$builder.Append("`t")
                }
                'br' {
                    [void]$builder.Append("`n")
                }
                'cr' {
                    [void]$builder.Append("`n")
                }
            }
        }
        $paragraphTexts.Add($builder.ToString())
    }

    return ($paragraphTexts -join "`n")
}

function Write-CsvFile {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string[]]$Columns,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Rows
    )

    if ($Rows.Count -eq 0) {
        $header = ($Columns | ForEach-Object { '"' + ($_ -replace '"', '""') + '"' }) -join ','
        Set-Content -LiteralPath $Path -Value $header -Encoding utf8
        return
    }

    $safeRows = foreach ($row in $Rows) {
        $safeProperties = [ordered]@{}
        foreach ($column in $Columns) {
            $safeProperties[$column] = Protect-SpreadsheetCell $row.$column
        }
        [pscustomobject]$safeProperties
    }
    $csv = $safeRows | ConvertTo-Csv -NoTypeInformation
    Set-Content -LiteralPath $Path -Value $csv -Encoding utf8
}

if (-not [string]::Equals([System.IO.Path]::GetExtension($InputDocx), '.docx', [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "InputDocx must have a .docx extension: $InputDocx"
}

$resolvedInput = Resolve-RequiredFile -Path $InputDocx -Label 'InputDocx'
$resolvedBibliography = Resolve-RequiredFile -Path $BibliographyCsv -Label 'BibliographyCsv'
$resolvedPolicy = Resolve-RequiredFile -Path $PolicyJson -Label 'PolicyJson'
$resolvedOutput = [System.IO.Path]::GetFullPath($OutputDirectory)
Assert-NoReparsePointInPath -Path $resolvedOutput -Label 'OutputDirectory'
$outputPaths = [ordered]@{
    footnotes = [System.IO.Path]::GetFullPath((Join-Path $resolvedOutput 'footnotes.csv'))
    citations = [System.IO.Path]::GetFullPath((Join-Path $resolvedOutput 'citations.csv'))
    issues = [System.IO.Path]::GetFullPath((Join-Path $resolvedOutput 'issues.csv'))
    summary = [System.IO.Path]::GetFullPath((Join-Path $resolvedOutput 'summary.json'))
    report = [System.IO.Path]::GetFullPath((Join-Path $resolvedOutput 'report.md'))
}

$protectedInputs = @($resolvedInput, $resolvedBibliography, $resolvedPolicy)
foreach ($outputPath in $outputPaths.Values) {
    foreach ($protectedInput in $protectedInputs) {
        if ([string]::Equals($outputPath, $protectedInput, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to overwrite an input file with an audit output: $outputPath"
        }
    }
}

$outputDirectoryExists = Test-Path -LiteralPath $resolvedOutput
if ($outputDirectoryExists) {
    if (-not (Test-Path -LiteralPath $resolvedOutput -PathType Container)) {
        throw "OutputDirectory exists but is not a directory: $resolvedOutput"
    }
    if (-not $Force) {
        throw "OutputDirectory already exists. Use -Force to write audit outputs into it: $resolvedOutput"
    }
}

$inputHashBefore = (Get-FileHash -LiteralPath $resolvedInput -Algorithm SHA256).Hash.ToLowerInvariant()

try {
    $policy = Get-Content -Raw -Encoding utf8 -LiteralPath $resolvedPolicy | ConvertFrom-Json
}
catch {
    throw "PolicyJson is not valid JSON: $($_.Exception.Message)"
}

$requiredPolicyFields = @(
    'policy_name',
    'subsequent_citation',
    'consecutive_same_source',
    'new_sources'
)
foreach ($field in $requiredPolicyFields) {
    $property = $policy.PSObject.Properties[$field]
    if ($null -eq $property -or $null -eq $property.Value -or
        ($property.Value -is [string] -and [string]::IsNullOrWhiteSpace($property.Value))) {
        throw "PolicyJson is missing required field or has a blank value: $field"
    }
}

$bibliography = @(Import-Csv -LiteralPath $resolvedBibliography -Encoding utf8)
if ($bibliography.Count -eq 0) {
    throw 'BibliographyCsv must contain at least one data row.'
}

$requiredBibliographyColumns = @(
    'source_id',
    'source_type',
    'language',
    'author',
    'short_title',
    'aliases',
    'bibliography_entry'
)
$availableColumns = @($bibliography[0].PSObject.Properties.Name)
foreach ($column in $requiredBibliographyColumns) {
    if ($availableColumns -notcontains $column) {
        throw "BibliographyCsv is missing required column: $column"
    }
}

$sourceIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$sources = [System.Collections.Generic.List[object]]::new()
foreach ($row in $bibliography) {
    $sourceId = ([string]$row.source_id).Trim()
    if ([string]::IsNullOrWhiteSpace($sourceId)) {
        throw 'BibliographyCsv contains a blank source_id.'
    }
    if (-not $sourceIds.Add($sourceId)) {
        throw "BibliographyCsv contains duplicate source_id values: $sourceId"
    }

    $normalizedAliases = [System.Collections.Generic.List[string]]::new()
    foreach ($alias in ([string]$row.aliases -split '\|')) {
        $normalizedAlias = Normalize-Whitespace $alias
        if (-not [string]::IsNullOrWhiteSpace($normalizedAlias) -and
            -not $normalizedAliases.Contains($normalizedAlias)) {
            $normalizedAliases.Add($normalizedAlias)
        }
    }

    $sources.Add([pscustomobject][ordered]@{
        source_id = $sourceId
        source_type = [string]$row.source_type
        language = [string]$row.language
        author = [string]$row.author
        short_title = [string]$row.short_title
        aliases = @($normalizedAliases)
        bibliography_entry = [string]$row.bibliography_entry
    })
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$archive = [System.IO.Compression.ZipFile]::OpenRead($resolvedInput)
try {
    $documentText = Read-ZipEntryText -Archive $archive -EntryName 'word/document.xml'
    $footnotesText = Read-ZipEntryText -Archive $archive -EntryName 'word/footnotes.xml'
}
finally {
    $archive.Dispose()
}

$documentXml = ConvertTo-SafeXmlDocument -XmlText $documentText -Label 'word/document.xml'
$footnotesXml = ConvertTo-SafeXmlDocument -XmlText $footnotesText -Label 'word/footnotes.xml'

$documentWordNamespace = $documentXml.DocumentElement.NamespaceURI
$footnoteWordNamespace = $footnotesXml.DocumentElement.NamespaceURI
if ([string]::IsNullOrWhiteSpace($documentWordNamespace) -or [string]::IsNullOrWhiteSpace($footnoteWordNamespace)) {
    throw 'DOCX word/document.xml or word/footnotes.xml has no WordprocessingML namespace.'
}
$documentNamespaces = [System.Xml.XmlNamespaceManager]::new($documentXml.NameTable)
$documentNamespaces.AddNamespace('w', $documentWordNamespace)
$footnoteNamespaces = [System.Xml.XmlNamespaceManager]::new($footnotesXml.NameTable)
$footnoteNamespaces.AddNamespace('w', $footnoteWordNamespace)

$footnoteBodies = @{}
foreach ($footnoteNode in $footnotesXml.SelectNodes('/w:footnotes/w:footnote', $footnoteNamespaces)) {
    $id = $footnoteNode.GetAttribute('id', $footnoteWordNamespace)
    if ($id -in @('-1', '0')) {
        continue
    }
    $footnoteBodies[$id] = Get-LogicalFootnoteText -FootnoteNode $footnoteNode -NamespaceManager $footnoteNamespaces
}

$footnoteRows = [System.Collections.Generic.List[object]]::new()
$citationRows = [System.Collections.Generic.List[object]]::new()
$issueRows = [System.Collections.Generic.List[object]]::new()
$seenSources = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$citedSources = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$previousMatchIds = @()
$footnoteNumber = 0

foreach ($referenceNode in $documentXml.SelectNodes('//w:footnoteReference', $documentNamespaces)) {
    $referenceId = $referenceNode.GetAttribute('id', $documentWordNamespace)
    if ($referenceId -in @('-1', '0')) {
        continue
    }

    $footnoteNumber++
    $hasFootnoteBody = $footnoteBodies.ContainsKey($referenceId)
    $logicalText = if ($hasFootnoteBody) {
        [string]$footnoteBodies[$referenceId]
    }
    else {
        ''
    }
    $normalizedText = Normalize-Whitespace $logicalText
    $matchedSources = [System.Collections.Generic.List[object]]::new()

    foreach ($source in $sources) {
        $matched = $false
        foreach ($alias in $source.aliases) {
            if (Test-AliasMatch -Text $normalizedText -Alias $alias) {
                $matched = $true
                break
            }
        }
        if ($matched) {
            $matchedSources.Add($source)
        }
    }

    $matchIds = @($matchedSources | ForEach-Object { $_.source_id })
    $footnoteRows.Add([pscustomobject][ordered]@{
        footnote_number = $footnoteNumber
        reference_id = $referenceId
        text = $logicalText
        match_count = $matchedSources.Count
        matched_source_ids = $matchIds -join '|'
    })

    if (-not $hasFootnoteBody) {
        $issueRows.Add([pscustomobject][ordered]@{
            issue_type = 'missing_footnote_body'
            severity = 'review'
            footnote_number = $footnoteNumber
            reference_id = $referenceId
            source_id = ''
            message = 'The document references a footnote ID that is absent from word/footnotes.xml; OOXML structure requires review.'
        })
    }
    elseif ($matchedSources.Count -eq 0) {
        $issueRows.Add([pscustomobject][ordered]@{
            issue_type = 'unmatched_footnote'
            severity = 'review'
            footnote_number = $footnoteNumber
            reference_id = $referenceId
            source_id = ''
            message = 'No bibliography alias matched this footnote; human review is required.'
        })
    }
    elseif ($matchedSources.Count -gt 1) {
        $issueRows.Add([pscustomobject][ordered]@{
            issue_type = 'multiple_source_matches'
            severity = 'review'
            footnote_number = $footnoteNumber
            reference_id = $referenceId
            source_id = $matchIds -join '|'
            message = 'Multiple bibliography sources matched this footnote; human disambiguation is required.'
        })
    }

    foreach ($source in $matchedSources) {
        $classification = if (
            $matchedSources.Count -eq 1 -and
            $previousMatchIds.Count -eq 1 -and
            [string]::Equals($source.source_id, $previousMatchIds[0], [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            'ibid_candidate'
        }
        elseif ($seenSources.Contains($source.source_id)) {
            'repeat'
        }
        else {
            'first'
        }

        $citationRows.Add([pscustomobject][ordered]@{
            footnote_number = $footnoteNumber
            reference_id = $referenceId
            source_id = $source.source_id
            citation_classification = $classification
            review_status = if ($matchedSources.Count -gt 1) { 'review_required' } else { '' }
            matched_text = $logicalText
            bibliography_entry = $source.bibliography_entry
        })
        [void]$citedSources.Add($source.source_id)
    }

    if ($matchedSources.Count -eq 1) {
        [void]$seenSources.Add($matchedSources[0].source_id)
    }

    $previousMatchIds = $matchIds
}

foreach ($source in $sources) {
    if (-not $citedSources.Contains($source.source_id)) {
        $issueRows.Add([pscustomobject][ordered]@{
            issue_type = 'unused_bibliography_entry'
            severity = 'review'
            footnote_number = ''
            reference_id = ''
            source_id = $source.source_id
            message = 'This bibliography entry was not matched to a footnote. Review only; this is not a deletion instruction.'
        })
    }
}

$footnotesPath = $outputPaths.footnotes
$citationsPath = $outputPaths.citations
$issuesPath = $outputPaths.issues
$summaryPath = $outputPaths.summary
$reportPath = $outputPaths.report

if (-not $outputDirectoryExists) {
    New-Item -ItemType Directory -Path $resolvedOutput | Out-Null
}
Assert-NoReparsePointInPath -Path $resolvedOutput -Label 'OutputDirectory'
foreach ($outputPath in $outputPaths.Values) {
    Prepare-OutputFile -Path $outputPath -OutputRoot $resolvedOutput
}

Write-CsvFile -Path $footnotesPath -Columns @(
    'footnote_number', 'reference_id', 'text', 'match_count', 'matched_source_ids'
) -Rows @($footnoteRows)
Write-CsvFile -Path $citationsPath -Columns @(
    'footnote_number', 'reference_id', 'source_id', 'citation_classification', 'review_status', 'matched_text', 'bibliography_entry'
) -Rows @($citationRows)
Write-CsvFile -Path $issuesPath -Columns @(
    'issue_type', 'severity', 'footnote_number', 'reference_id', 'source_id', 'message'
) -Rows @($issueRows)

$inputHashAfter = (Get-FileHash -LiteralPath $resolvedInput -Algorithm SHA256).Hash.ToLowerInvariant()
if (-not [string]::Equals($inputHashBefore, $inputHashAfter, [System.StringComparison]::Ordinal)) {
    throw "Input DOCX changed during audit. Before: $inputHashBefore; after: $inputHashAfter"
}

$summary = [pscustomobject][ordered]@{
    schema_version = 1
    tool_name = 'thesis-footnote-normalizer-audit'
    mode = 'audit_only'
    input_file_name = [System.IO.Path]::GetFileName($resolvedInput)
    input_sha256_before = $inputHashBefore
    input_sha256_after = $inputHashAfter
    input_unchanged = $true
    policy = $policy
    counts = [pscustomobject][ordered]@{
        footnotes = $footnoteRows.Count
        citations = $citationRows.Count
        issues = $issueRows.Count
        unmatched_footnotes = @($issueRows | Where-Object issue_type -eq 'unmatched_footnote').Count
        multiple_source_matches = @($issueRows | Where-Object issue_type -eq 'multiple_source_matches').Count
        missing_footnote_bodies = @($issueRows | Where-Object issue_type -eq 'missing_footnote_body').Count
        unused_bibliography_entries = @($issueRows | Where-Object issue_type -eq 'unused_bibliography_entry').Count
    }
    limitations = @(
        'Version 1 audits true OOXML footnotes and does not rewrite the DOCX.',
        'Alias matches identify review candidates and do not establish bibliographic correctness.',
        'Aliases are Unicode-normalized and use letter/number/mark boundaries to reduce substring false positives.',
        'footnote_number is document-reference order and may differ from numbering rendered by Word.',
        'CSV cells that could be evaluated as spreadsheet formulas are prefixed with an apostrophe.',
        'No formatted replacement citations are generated.'
    )
}
Set-Content -LiteralPath $summaryPath -Value ($summary | ConvertTo-Json -Depth 20) -Encoding utf8

$reportLines = @(
    '# Footnote audit report',
    '',
    '> Audit-only output: the input DOCX was not rewritten. Alias matches are review candidates, not claims of bibliographic correctness.',
    '',
    '## Summary',
    '',
    "- Footnotes audited: $($footnoteRows.Count)",
    "- Matched citation rows: $($citationRows.Count)",
    "- Review issues: $($issueRows.Count)",
    "- Input SHA-256 before: ``$inputHashBefore``",
    "- Input SHA-256 after: ``$inputHashAfter``",
    '',
    '## Review boundaries',
    '',
    '- `ibid_candidate` means only that adjacent footnotes each matched exactly the same single source.',
    '- Unmatched and multiple-match footnotes require human review.',
    '- `footnote_number` is document-reference order, not a reconstruction of numbering rendered by Word.',
    '- Unused bibliography entries require review and are not deletion instructions.',
    '- Version 1 does not rewrite the DOCX or generate formatted replacement citations.'
)
Set-Content -LiteralPath $reportPath -Value $reportLines -Encoding utf8

$finalInputHash = (Get-FileHash -LiteralPath $resolvedInput -Algorithm SHA256).Hash.ToLowerInvariant()
if (-not [string]::Equals($inputHashBefore, $finalInputHash, [System.StringComparison]::Ordinal)) {
    throw "Input DOCX changed during audit. Before: $inputHashBefore; final: $finalInputHash"
}

[pscustomobject][ordered]@{
    output_directory = $resolvedOutput
    footnotes = $footnoteRows.Count
    citations = $citationRows.Count
    issues = $issueRows.Count
    input_unchanged = $true
}
