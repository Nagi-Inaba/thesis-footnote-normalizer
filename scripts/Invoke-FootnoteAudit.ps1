[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$InputDocx,

    [Parameter(Mandatory)]
    [string]$BibliographyCsv,

    [Parameter(Mandatory)]
    [string]$PolicyJson,

    [string]$BibliographyDocx,

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

function Normalize-CitationVariant {
    param([AllowNull()][string]$Value)

    $normalized = Normalize-Whitespace $Value

    $normalized = [regex]::Replace(
        $normalized,
        '\bpp\.\s*([0-9]+(?:\s*[-‒–—−]\s*[0-9]+)?)',
        'pp. <page>',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $normalized = [regex]::Replace(
        $normalized,
        '\bp\.\s*([0-9]+(?:\s*[-‒–—−]\s*[0-9]+)?)',
        'p. <page>',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $normalized = [regex]::Replace(
        $normalized,
        '(\d+\s*[-‒–—−]\s*\d+\s*(?:頁以下|頁))',
        '<page>',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $normalized = [regex]::Replace(
        $normalized,
        '(\d+\s*(?:頁以下|頁))',
        '<page>',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $normalized = [regex]::Replace($normalized, '(<page>)\s*[-‒–—−]\s*(<page>)', '<page>')
    $normalized = [regex]::Replace(
        $normalized,
        '(\d+)\s*頁(以下)?',
        ' ${1}頁$2',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    return [regex]::Replace($normalized, '([0-9])\s*[-‒–—−]\s*([0-9])', '$1-$2')
}

function Get-BibliographyValue {
    param(
        [Parameter(Mandatory)]
        [psobject]$Row,

        [Parameter(Mandatory)]
        [string]$FieldName
    )

    $property = $Row.PSObject.Properties[[string]$FieldName]
    if ($null -eq $property) {
        return ''
    }
    return [string]$property.Value
}

function Get-FieldList {
    param([AllowNull()]$Value)

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [string]) {
        $trimmed = $Value.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            return @()
        }
        return @($trimmed)
    }

    if ($Value -is [System.Array]) {
        $items = [System.Collections.Generic.List[string]]::new()
        foreach ($entry in $Value) {
            if ($null -eq $entry) {
                continue
            }
            $text = ([string]$entry).Trim()
            if (-not [string]::IsNullOrWhiteSpace($text) -and -not $items.Contains($text)) {
                $items.Add($text)
            }
        }
        return @($items)
    }

    return @([string]$Value)
}

function Get-PunctuationSignature {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }
    $normalized = Normalize-Whitespace $Value
    if ($normalized -match '([^\s0-9A-Za-z\u3040-\u30FF\u4E00-\u9FFF]+)\s*$') {
        return $matches[1]
    }
    return ''
}

function Get-TerminalMark {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }
    $trimmed = $Value.Trim()
    return [string]$trimmed[$trimmed.Length - 1]
}

function Get-ComponentListForPolicy {
    param([AllowNull()]$PolicyRule, [bool]$IsFirstUse)

    $fieldName = if ($IsFirstUse) { 'first_use_required_fields' } else { 'subsequent_use_required_fields' }
    if ($null -eq $PolicyRule) {
        return @()
    }

    $property = $PolicyRule.PSObject.Properties[$fieldName]
    if ($null -eq $property) {
        return @()
    }
    return Get-FieldList -Value $property.Value
}

function New-IssueRow {
    param(
        [Parameter(Mandatory)][string]$IssueType,
        [Parameter(Mandatory)][string]$Severity,
        [AllowNull()][string]$FootnoteNumber,
        [AllowNull()][string]$ReferenceId,
        [AllowNull()][string]$SourceId,
        [AllowNull()][string]$Message,
        [AllowNull()][string]$RelatedFootnoteNumber = '',
        [AllowNull()][string]$RelatedSourceId = '',
        [AllowNull()][string]$Evidence = ''
    )

    return [pscustomobject][ordered]@{
        issue_type = $IssueType
        severity = $Severity
        footnote_number = if ($null -eq $FootnoteNumber) { '' } else { [string]$FootnoteNumber }
        reference_id = if ($null -eq $ReferenceId) { '' } else { [string]$ReferenceId }
        source_id = if ($null -eq $SourceId) { '' } else { [string]$SourceId }
        message = if ($null -eq $Message) { '' } else { [string]$Message }
        related_footnote_number = if ($null -eq $RelatedFootnoteNumber) { '' } else { [string]$RelatedFootnoteNumber }
        related_source_id = if ($null -eq $RelatedSourceId) { '' } else { [string]$RelatedSourceId }
        evidence = if ($null -eq $Evidence) { '' } else { [string]$Evidence }
    }
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

function Assert-SafeZipContainer {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Label
    )

    $file = Get-Item -LiteralPath $Path -Force
    if ($file.Length -gt 128MB) {
        throw "$Label exceeds the 128 MiB compressed-file audit limit."
    }

    $stream = [System.IO.File]::Open($file.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
        $tailLength = [int][Math]::Min($stream.Length, 65557)
        $tail = [byte[]]::new($tailLength)
        [void]$stream.Seek(-$tailLength, [System.IO.SeekOrigin]::End)
        if ($stream.Read($tail, 0, $tail.Length) -ne $tail.Length) {
            throw "$Label ZIP end record could not be read safely."
        }

        $endRecordOffset = -1
        for ($index = $tail.Length - 22; $index -ge 0; $index--) {
            if ($tail[$index] -eq 0x50 -and $tail[$index + 1] -eq 0x4B -and
                $tail[$index + 2] -eq 0x05 -and $tail[$index + 3] -eq 0x06) {
                $commentLength = [BitConverter]::ToUInt16($tail, $index + 20)
                $candidateCentralDirectoryBytes = [BitConverter]::ToUInt32($tail, $index + 12)
                $candidateCentralDirectoryOffset = [BitConverter]::ToUInt32($tail, $index + 16)
                $candidateAbsoluteOffset = ($stream.Length - $tail.Length) + $index
                $endsAtFileBoundary = ($index + 22 + $commentLength) -eq $tail.Length
                $centralDirectoryEndsAtRecord = ([UInt64]$candidateCentralDirectoryOffset + [UInt64]$candidateCentralDirectoryBytes) -eq [UInt64]$candidateAbsoluteOffset
                $singleDisk = ([BitConverter]::ToUInt16($tail, $index + 4) -eq 0) -and
                    ([BitConverter]::ToUInt16($tail, $index + 6) -eq 0) -and
                    ([BitConverter]::ToUInt16($tail, $index + 8) -eq [BitConverter]::ToUInt16($tail, $index + 10))
                if ($endsAtFileBoundary -and $centralDirectoryEndsAtRecord -and $singleDisk) {
                    $endRecordOffset = $index
                    break
                }
            }
        }
        if ($endRecordOffset -lt 0) {
            throw "$Label is missing a supported ZIP end record."
        }

        $entryCount = [BitConverter]::ToUInt16($tail, $endRecordOffset + 10)
        $centralDirectoryBytes = [BitConverter]::ToUInt32($tail, $endRecordOffset + 12)
        $centralDirectoryOffset = [BitConverter]::ToUInt32($tail, $endRecordOffset + 16)
        if ($entryCount -eq [UInt16]::MaxValue -or
            $centralDirectoryBytes -eq [UInt32]::MaxValue -or
            $centralDirectoryOffset -eq [UInt32]::MaxValue) {
            throw "$Label uses ZIP64 metadata, which is not accepted by this audit tool."
        }
        if ($entryCount -gt 4096) {
            throw "$Label exceeds the permitted ZIP entry count of 4096."
        }
        if ($centralDirectoryBytes -gt 16MB) {
            throw "$Label central directory exceeds the 16 MiB audit limit."
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Open-SafeDocxArchive {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Label
    )

    Assert-SafeZipContainer -Path $Path -Label $Label
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        if ($archive.Entries.Count -gt 4096) {
            throw "$Label exceeds the permitted ZIP entry count of 4096."
        }
        [UInt64]$aggregateLength = 0
        foreach ($archiveEntry in $archive.Entries) {
            $aggregateLength += [UInt64]$archiveEntry.Length
            if ($aggregateLength -gt 256MB) {
                throw "$Label exceeds the 256 MiB aggregate uncompressed audit limit."
            }
        }
        return ,$archive
    }
    catch {
        $archive.Dispose()
        throw
    }
}

function Read-ZipEntryText {
    param(
        [Parameter(Mandatory)]
        [System.IO.Compression.ZipArchive]$Archive,

        [Parameter(Mandatory)]
        [string]$EntryName
    )

    $matchingEntries = @($Archive.Entries | Where-Object { $_.FullName -ceq $EntryName })
    if ($matchingEntries.Count -eq 0) {
        throw "DOCX package is missing required part: $EntryName"
    }
    if ($matchingEntries.Count -gt 1) {
        throw "DOCX package contains a duplicate required part: $EntryName"
    }
    $entry = $matchingEntries[0]

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

function Get-LogicalParagraphText {
    param(
        [Parameter(Mandatory)]
        [System.Xml.XmlElement]$ParagraphNode,

        [Parameter(Mandatory)]
        [System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    $builder = [System.Text.StringBuilder]::new()
    foreach ($element in $ParagraphNode.SelectNodes('.//*')) {
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

    return $builder.ToString()
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
    citation_variants = [System.IO.Path]::GetFullPath((Join-Path $resolvedOutput 'citation-variants.csv'))
    issues = [System.IO.Path]::GetFullPath((Join-Path $resolvedOutput 'issues.csv'))
    bibliography_reconciliation = [System.IO.Path]::GetFullPath((Join-Path $resolvedOutput 'bibliography-reconciliation.csv'))
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

$hasBibliographyDocumentPolicy = $policy.PSObject.Properties['bibliography_document'] -ne $null
$bibliographyDocumentPolicy = if ($hasBibliographyDocumentPolicy) { $policy.bibliography_document } else { $null }
$bibliographyDocumentPolicyEnabled = $false
$bibliographyPolicyStartMarker = ''
$bibliographyPolicyEndMarker = ''
$bibliographyPolicyIncludeHeading = $false
$bibliographyPolicyParagraphMatchMode = 'contains'

if ($null -ne $bibliographyDocumentPolicy) {
    if ($bibliographyDocumentPolicy.PSObject.Properties['enabled'] -ne $null) {
        $bibliographyDocumentPolicyEnabled = [bool]$bibliographyDocumentPolicy.enabled
    }
    if ($bibliographyDocumentPolicy.PSObject.Properties['start_marker'] -ne $null) {
        $bibliographyPolicyStartMarker = [string]$bibliographyDocumentPolicy.start_marker
    }
    if ($bibliographyDocumentPolicy.PSObject.Properties['end_marker'] -ne $null) {
        $bibliographyPolicyEndMarker = [string]$bibliographyDocumentPolicy.end_marker
    }
    if ($bibliographyDocumentPolicy.PSObject.Properties['include_heading'] -ne $null) {
        $bibliographyPolicyIncludeHeading = [bool]$bibliographyDocumentPolicy.include_heading
    }
    if ($bibliographyDocumentPolicy.PSObject.Properties['paragraph_match_mode'] -ne $null) {
        $bibliographyPolicyParagraphMatchMode = [string]$bibliographyDocumentPolicy.paragraph_match_mode
    }
}

$bibliographyReconciliationEnabled = $false
if ($bibliographyDocumentPolicyEnabled) {
    $bibliographyReconciliationEnabled = $true
}
elseif (-not [string]::IsNullOrWhiteSpace($BibliographyDocx)) {
    throw 'Bibliography docx was provided but bibliography_document.enabled is false or missing in policy.'
}

$resolvedBibliographyDocument = if ($bibliographyReconciliationEnabled) {
    if ([string]::IsNullOrWhiteSpace($BibliographyDocx)) {
        $resolvedInput
    }
    else {
        if (-not [string]::Equals([System.IO.Path]::GetExtension($BibliographyDocx), '.docx', [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "BibliographyDocx must have a .docx extension: $BibliographyDocx"
        }
        Resolve-RequiredFile -Path $BibliographyDocx -Label 'BibliographyDocx'
    }
} else {
    $null
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
$hasSourceTypePolicies = $policy.PSObject.Properties['source_type_policies'] -ne $null
$sourceTypePolicies = @{}
if ($hasSourceTypePolicies) {
    if ($policy.source_type_policies -isnot [psobject]) {
        throw 'policy.source_type_policies must be an object.'
    }
    if (-not ($policy.source_type_policies.PSObject.Properties.Name -contains 'default')) {
        throw 'policy.source_type_policies must contain a default rule.'
    }
    foreach ($policyEntry in $policy.source_type_policies.PSObject.Properties) {
        $sourceType = [string]$policyEntry.Name
        $sourceTypePolicies[$sourceType] = $policyEntry.Value
    }
}

$sourceIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$sources = [System.Collections.Generic.List[object]]::new()
$sourceTypePolicyMissingSources = [System.Collections.Generic.List[string]]::new()
foreach ($row in $bibliography) {
    $sourceId = ([string]$row.source_id).Trim()
    if ([string]::IsNullOrWhiteSpace($sourceId)) {
        throw 'BibliographyCsv contains a blank source_id.'
    }
    if (-not $sourceIds.Add($sourceId)) {
        throw "BibliographyCsv contains duplicate source_id values: $sourceId"
    }
    $sourceType = [string]$row.source_type
    $hasTypeRule = $sourceTypePolicies.ContainsKey($sourceType)
    if ($hasSourceTypePolicies -and -not $hasTypeRule) {
        $sourceTypePolicyMissingSources.Add($sourceId)
    }

    $sourceTypeRule = if ($hasTypeRule) { $sourceTypePolicies[$sourceType] } else { $sourceTypePolicies['default'] }
    $sourcePolicyKey = if ($hasTypeRule) { $sourceType } else { 'default' }
    $effectiveConsecutivePolicy = if ($null -ne $sourceTypeRule -and
        $sourceTypeRule.PSObject.Properties['consecutive_same_source'] -ne $null) {
        [string]$sourceTypeRule.consecutive_same_source
    } else {
        [string]$policy.consecutive_same_source
    }
    $requiredFieldForFirstUse = @()
    $requiredFieldForSubsequentUse = @()
    if ($null -ne $sourceTypeRule) {
        $firstUseProperty = $sourceTypeRule.PSObject.Properties['first_use_required_fields']
        $subsequentUseProperty = $sourceTypeRule.PSObject.Properties['subsequent_use_required_fields']
        if ($null -ne $firstUseProperty) {
            $requiredFieldForFirstUse = Get-FieldList -Value $firstUseProperty.Value
        }
        if ($null -ne $subsequentUseProperty) {
            $requiredFieldForSubsequentUse = Get-FieldList -Value $subsequentUseProperty.Value
        }
    }

    $normalizedAliases = [System.Collections.Generic.List[string]]::new()
    foreach ($alias in ([string]$row.aliases -split '\|')) {
        $normalizedAlias = Normalize-Whitespace $alias
        if (-not [string]::IsNullOrWhiteSpace($normalizedAlias) -and
            -not $normalizedAliases.Contains($normalizedAlias)) {
            $normalizedAliases.Add($normalizedAlias)
        }
    }

    $bibliographyAliases = [System.Collections.Generic.List[string]]::new()
    if ($row.PSObject.Properties.Name -contains 'bibliography_aliases') {
        foreach ($alias in ([string]$row.bibliography_aliases -split '\|')) {
            $normalizedAlias = Normalize-Whitespace $alias
            if (-not [string]::IsNullOrWhiteSpace($normalizedAlias) -and
                -not $bibliographyAliases.Contains($normalizedAlias)) {
                $bibliographyAliases.Add($normalizedAlias)
            }
        }
    }

    $sourceData = [ordered]@{
        source_id = $sourceId
        source_type = $sourceType
        language = [string]$row.language
        author = [string]$row.author
        short_title = [string]$row.short_title
        aliases = @($normalizedAliases)
        bibliography_entry = [string]$row.bibliography_entry
        source_policy_key = $sourcePolicyKey
        first_use_required_fields = @($requiredFieldForFirstUse)
        subsequent_use_required_fields = @($requiredFieldForSubsequentUse)
        consecutive_same_source = $effectiveConsecutivePolicy
        title = Get-BibliographyValue -Row $row -FieldName 'title'
        translator = Get-BibliographyValue -Row $row -FieldName 'translator'
        publication_place = Get-BibliographyValue -Row $row -FieldName 'publication_place'
        publisher = Get-BibliographyValue -Row $row -FieldName 'publisher'
        year = Get-BibliographyValue -Row $row -FieldName 'year'
        journal = Get-BibliographyValue -Row $row -FieldName 'journal'
        volume = Get-BibliographyValue -Row $row -FieldName 'volume'
        issue = Get-BibliographyValue -Row $row -FieldName 'issue'
        bibliography_aliases = @($bibliographyAliases)
    }
    foreach ($property in $row.PSObject.Properties) {
        if (-not $sourceData.Contains($property.Name)) {
            $sourceData[$property.Name] = [string]$property.Value
        }
    }
    $sources.Add([pscustomobject]$sourceData)
}

$bibliographyDocumentHashBefore = if ($bibliographyReconciliationEnabled -and $null -ne $resolvedBibliographyDocument) {
    (Get-FileHash -LiteralPath $resolvedBibliographyDocument -Algorithm SHA256).Hash.ToLowerInvariant()
} else {
    ''
}
$protectedInputs = @($resolvedInput, $resolvedBibliography, $resolvedPolicy)
if ($null -ne $resolvedBibliographyDocument) {
    $protectedInputs += $resolvedBibliographyDocument
}
foreach ($outputPath in $outputPaths.Values) {
    foreach ($protectedInput in $protectedInputs) {
        if ([string]::Equals($outputPath, $protectedInput, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to overwrite an input file with an audit output: $outputPath"
        }
    }
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$archive = Open-SafeDocxArchive -Path $resolvedInput -Label 'InputDocx'
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
$specialFootnoteIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($footnoteNode in $footnotesXml.SelectNodes('/w:footnotes/w:footnote', $footnoteNamespaces)) {
    $id = $footnoteNode.GetAttribute('id', $footnoteWordNamespace)
    $type = $footnoteNode.GetAttribute('type', $footnoteWordNamespace)
    if (-not [string]::IsNullOrWhiteSpace($type) -and
        -not [string]::Equals($type, 'normal', [System.StringComparison]::OrdinalIgnoreCase)) {
        [void]$specialFootnoteIds.Add($id)
        continue
    }
    $footnoteBodies[$id] = Get-LogicalFootnoteText -FootnoteNode $footnoteNode -NamespaceManager $footnoteNamespaces
}

$footnoteRows = [System.Collections.Generic.List[object]]::new()
$citationRows = [System.Collections.Generic.List[object]]::new()
$issueRows = [System.Collections.Generic.List[object]]::new()
$citationVariantRows = [System.Collections.Generic.List[object]]::new()
$referenceCandidateBySourceClass = @{}
$seenSources = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$citedSources = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

$bibliographyReconciliationRows = [System.Collections.Generic.List[object]]::new()
$matchedBibliographySourceIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$observedBibliographySourceIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$bibliographyReconciliationStatus = if ($bibliographyReconciliationEnabled) { 'evaluated' } else { 'not_observed' }
$bibliographyReconciliationParagraphCount = 0
$bibliographyReconciliationInputFile = if ($null -ne $resolvedBibliographyDocument) { [System.IO.Path]::GetFileName($resolvedBibliographyDocument) } else { '' }

if ($bibliographyReconciliationEnabled) {
    $bibliographyArchive = Open-SafeDocxArchive -Path $resolvedBibliographyDocument -Label 'BibliographyDocx'
    try {
        $bibliographyDocumentText = Read-ZipEntryText -Archive $bibliographyArchive -EntryName 'word/document.xml'
    }
    finally {
        $bibliographyArchive.Dispose()
    }

    $bibliographyDocumentXml = ConvertTo-SafeXmlDocument -XmlText $bibliographyDocumentText -Label 'bibliography word/document.xml'
    $bibliographyDocumentWordNamespace = $bibliographyDocumentXml.DocumentElement.NamespaceURI
    if ([string]::IsNullOrWhiteSpace($bibliographyDocumentWordNamespace)) {
        throw 'Bibliography DOCX word/document.xml has no WordprocessingML namespace.'
    }
    $bibliographyDocumentNamespaces = [System.Xml.XmlNamespaceManager]::new($bibliographyDocumentXml.NameTable)
    $bibliographyDocumentNamespaces.AddNamespace('w', $bibliographyDocumentWordNamespace)
    $documentBibliographyParagraphs = @()
    foreach ($paragraphNode in $bibliographyDocumentXml.SelectNodes('//w:body/w:p', $bibliographyDocumentNamespaces)) {
        $paragraphText = Normalize-Whitespace (Get-LogicalParagraphText -ParagraphNode $paragraphNode -NamespaceManager $bibliographyDocumentNamespaces)
        if (-not [string]::IsNullOrWhiteSpace($paragraphText)) {
            $documentBibliographyParagraphs += $paragraphText
        }
    }

    $normalizedStartMarker = Normalize-Whitespace $bibliographyPolicyStartMarker
    $normalizedEndMarker = Normalize-Whitespace $bibliographyPolicyEndMarker
    $matchMode = if ([string]::IsNullOrWhiteSpace($bibliographyPolicyParagraphMatchMode)) { 'contains' } else { [string]$bibliographyPolicyParagraphMatchMode }

    $startMatchIndex = -1
    for ($paragraphIndex = 0; $paragraphIndex -lt $documentBibliographyParagraphs.Count; $paragraphIndex++) {
        $paragraphText = $documentBibliographyParagraphs[$paragraphIndex]
        if ([string]::IsNullOrWhiteSpace($normalizedStartMarker)) {
            $startMatchIndex = -1
            break
        }
        if ([string]::Equals($matchMode, 'exact', [System.StringComparison]::OrdinalIgnoreCase)) {
            if ([string]::Equals($paragraphText, $normalizedStartMarker, [System.StringComparison]::OrdinalIgnoreCase)) {
                $startMatchIndex = $paragraphIndex
                break
            }
        }
        elseif ($paragraphText.IndexOf($normalizedStartMarker, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $startMatchIndex = $paragraphIndex
            break
        }
    }

    if ($startMatchIndex -lt 0) {
        $bibliographyReconciliationStatus = 'marker_not_found'
        $issueRows.Add((New-IssueRow -IssueType 'document_bibliography_not_found' -Severity 'review' -FootnoteNumber '' -ReferenceId '' -SourceId '' -Message 'Bibliography start marker not found in the designated DOCX.' -Evidence "start_marker=$normalizedStartMarker; paragraph_match_mode=$matchMode"))
    }
    else {
        $start = if ($bibliographyPolicyIncludeHeading) { $startMatchIndex } else { $startMatchIndex + 1 }
        if ($start -lt 0) {
            $start = 0
        }

        $end = $documentBibliographyParagraphs.Count
        if (-not [string]::IsNullOrWhiteSpace($normalizedEndMarker)) {
            for ($paragraphIndex = $start; $paragraphIndex -lt $documentBibliographyParagraphs.Count; $paragraphIndex++) {
                $paragraphText = $documentBibliographyParagraphs[$paragraphIndex]
                if ([string]::Equals($matchMode, 'exact', [System.StringComparison]::OrdinalIgnoreCase)) {
                    if ([string]::Equals($paragraphText, $normalizedEndMarker, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $end = $paragraphIndex
                        break
                    }
                }
                elseif ($paragraphText.IndexOf($normalizedEndMarker, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $end = $paragraphIndex
                    break
                }
            }
        }

        $paragraphNumber = 1
        for ($paragraphIndex = $start; $paragraphIndex -lt $end; $paragraphIndex++) {
            $paragraphText = $documentBibliographyParagraphs[$paragraphIndex]
            $matchedSources = [System.Collections.Generic.List[object]]::new()
            foreach ($source in $sources) {
                $aliasSources = if ($source.bibliography_aliases.Count -gt 0) { $source.bibliography_aliases } else { $source.aliases }
                $isMatched = $false
                foreach ($alias in $aliasSources) {
                    if (Test-AliasMatch -Text $paragraphText -Alias $alias) {
                        $isMatched = $true
                        break
                    }
                }
                if ($isMatched) {
                    $matchedSources.Add($source)
                }
            }

            $matchCount = $matchedSources.Count
            $matchIds = @($matchedSources | ForEach-Object { $_.source_id })
            $matchStatus = if ($matchCount -eq 1) {
                'matched'
            }
            elseif ($matchCount -gt 1) {
                'multiple_matches'
            }
            else {
                'unmatched'
            }

            if ($matchCount -eq 1) {
                [void]$matchedBibliographySourceIds.Add($matchIds[0])
                [void]$observedBibliographySourceIds.Add($matchIds[0])
            }
            elseif ($matchCount -gt 1) {
                foreach ($candidateId in $matchIds) {
                    [void]$observedBibliographySourceIds.Add($candidateId)
                }
                $issueRows.Add((New-IssueRow -IssueType 'document_bibliography_multiple_matches' -Severity 'review' -FootnoteNumber '' -ReferenceId '' -SourceId ($matchIds -join '|') -Message 'One bibliography paragraph matched multiple registry sources; human disambiguation is required.' -Evidence "paragraph_number=$paragraphNumber"))
            }
            else {
                $issueRows.Add((New-IssueRow -IssueType 'document_bibliography_unmatched' -Severity 'review' -FootnoteNumber '' -ReferenceId '' -SourceId '' -Message 'No bibliography source alias matched this bibliography paragraph.' -Evidence "paragraph_number=$paragraphNumber; text=$paragraphText"))
            }

            $bibliographyReconciliationRows.Add([pscustomobject][ordered]@{
                paragraph_number = [string]$paragraphNumber
                text = $paragraphText
                match_count = $matchCount
                matched_source_ids = $matchIds -join '|'
                status = $matchStatus
            })
            $paragraphNumber++
        }
        $bibliographyReconciliationParagraphCount = $paragraphNumber - 1
    }
}

$previousSingleMatchSourceId = ''
$footnoteNumber = 0
foreach ($missingSource in $sourceTypePolicyMissingSources) {
    $issueRows.Add((New-IssueRow -IssueType 'source_type_policy_missing' -Severity 'review' -FootnoteNumber '' -ReferenceId '' -SourceId $missingSource -Message 'No source-type policy exists for this source_type; default policy was used for component evaluation.'))
}

$japaneseTerminalMark = if ($policy.PSObject.Properties['japanese_note_terminal_mark'] -ne $null) {
    [string]$policy.japanese_note_terminal_mark
} else {
    ''
}
$foreignTerminalMark = if ($policy.PSObject.Properties['foreign_note_terminal_mark'] -ne $null) {
    [string]$policy.foreign_note_terminal_mark
} else {
    ''
}

function Get-ExpectedTerminalMark {
    param([AllowNull()][string]$Language)

    if ([string]::Equals(([string]$Language).Trim(), 'ja', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $japaneseTerminalMark
    }
    return $foreignTerminalMark
}

function Get-SourceComponentEvaluation {
    param(
        [Parameter(Mandatory)]
        [object]$Source,

        [Parameter(Mandatory)]
        [bool]$IsFirstUse,

        [Parameter(Mandatory)]
        [string]$NormalizedText
    )

    $requiredFields = @(
        if ($IsFirstUse) {
        Get-ComponentListForPolicy -PolicyRule $Source -IsFirstUse $true
        } else {
        Get-ComponentListForPolicy -PolicyRule $Source -IsFirstUse $false
        }
    )
    if ($null -eq $requiredFields) {
        $requiredFields = @()
    }
    $presentComponents = [System.Collections.Generic.List[string]]::new()
    $missingComponents = [System.Collections.Generic.List[string]]::new()
    $missingMetadataComponents = [System.Collections.Generic.List[string]]::new()
    $comparisonStatus = 'evaluated'
    if ($requiredFields.Count -eq 0) {
        return [pscustomobject][ordered]@{
            comparison_status = $comparisonStatus
            required_fields = @()
            present_fields = @()
            missing_fields = @()
            missing_metadata_fields = @()
            completeness_score = 0
            normalized_variant = Normalize-CitationVariant -Value $NormalizedText
            punctuation_signature = Get-PunctuationSignature -Value (Normalize-CitationVariant -Value $NormalizedText)
            terminal_mark = Get-TerminalMark -Value $NormalizedText
        }
    }

    foreach ($fieldName in $requiredFields) {
        $registryValue = Get-BibliographyValue -Row $Source -FieldName $fieldName
        if ([string]::IsNullOrWhiteSpace($registryValue)) {
            [void]$missingMetadataComponents.Add($fieldName)
            continue
        }
        if (Test-AliasMatch -Text $NormalizedText -Alias $registryValue) {
            [void]$presentComponents.Add($fieldName)
        }
        else {
            [void]$missingComponents.Add($fieldName)
        }
    }

    $presentCount = $presentComponents.Count
    $requiredCount = $requiredFields.Count
    $completeness = if ($requiredCount -eq 0) { 100 } else { [int]([double]($presentCount * 100) / $requiredCount) }

    return [pscustomobject][ordered]@{
        comparison_status = $comparisonStatus
        required_fields = @($requiredFields)
        present_fields = @($presentComponents)
        missing_fields = @($missingComponents + $missingMetadataComponents)
        missing_metadata_fields = @($missingMetadataComponents)
        completeness_score = $completeness
        normalized_variant = Normalize-CitationVariant -Value $NormalizedText
        punctuation_signature = Get-PunctuationSignature -Value (Normalize-CitationVariant -Value $NormalizedText)
        terminal_mark = Get-TerminalMark -Value $NormalizedText
    }
}

$latinShortFormPattern = '(?<![\p{L}\p{N}\p{M}])(?:ibidem|ibid\.?|op\.\s*cit\.?)(?![\p{L}\p{N}\p{M}])'
$shortFormPattern = "(?:同上|前掲(?:書|論文)?|$latinShortFormPattern)"
$contextualShortFormPattern = "(?:同上|(?<![\p{L}\p{N}\p{M}])(?:ibidem|ibid\.?)(?![\p{L}\p{N}\p{M}]))"

foreach ($referenceNode in $documentXml.SelectNodes('//w:footnoteReference', $documentNamespaces)) {
    $referenceId = $referenceNode.GetAttribute('id', $documentWordNamespace)
    if ($specialFootnoteIds.Contains($referenceId)) {
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
    $hasShortFormMarker = [regex]::IsMatch(
        $normalizedText,
        $shortFormPattern,
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
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

    $matchCount = $matchedSources.Count
    $matchIds = @($matchedSources | ForEach-Object { $_.source_id })

    $contextSourceId = ''
    $identityStatus = ''
    $comparisonStatus = 'not_evaluated_ambiguous_identity'
    if (-not $hasFootnoteBody) {
        $issueRows.Add((New-IssueRow -IssueType 'missing_footnote_body' -Severity 'review' -FootnoteNumber $footnoteNumber -ReferenceId $referenceId -SourceId '' -Message 'The document references a footnote ID that is absent from word/footnotes.xml; OOXML structure requires review.'))
    }
    elseif ($matchCount -eq 0) {
        if ($hasShortFormMarker) {
            $hasContextualShortForm = [regex]::IsMatch(
                $normalizedText,
                $contextualShortFormPattern,
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            )
            if ($hasContextualShortForm -and -not [string]::IsNullOrWhiteSpace($previousSingleMatchSourceId)) {
                $contextSourceId = $previousSingleMatchSourceId
                $identityStatus = 'context_inferred_review_required'
                $issueRows.Add((New-IssueRow -IssueType 'contextual_shorthand_candidate' -Severity 'review' -FootnoteNumber $footnoteNumber -ReferenceId $referenceId -SourceId $contextSourceId -Message 'Short form appears to inherit source context; identity requires review.'))
            }
            elseif ($hasContextualShortForm) {
                $issueRows.Add((New-IssueRow -IssueType 'unresolved_shorthand' -Severity 'review' -FootnoteNumber $footnoteNumber -ReferenceId $referenceId -SourceId '' -Message 'No explicit source alias was matched in context-only shorthand.'))
            }
            else {
                $issueRows.Add((New-IssueRow -IssueType 'unresolved_shorthand' -Severity 'review' -FootnoteNumber $footnoteNumber -ReferenceId $referenceId -SourceId '' -Message 'No explicit source alias was matched in shorthand citation.'))
            }
        }
        else {
            $issueRows.Add((New-IssueRow -IssueType 'unmatched_footnote' -Severity 'review' -FootnoteNumber $footnoteNumber -ReferenceId $referenceId -SourceId '' -Message 'No explicit source alias was matched in this footnote.'))
        }
        $comparisonStatus = 'not_configured'
    }
    elseif ($matchCount -gt 1) {
        $issueRows.Add((New-IssueRow -IssueType 'multiple_source_matches' -Severity 'review' -FootnoteNumber $footnoteNumber -ReferenceId $referenceId -SourceId ($matchIds -join '|') -Message 'Multiple bibliography sources matched this footnote; human disambiguation is required.'))
        foreach ($ambiguousSource in $matchedSources) {
            $ambiguousSourceId = [string]$ambiguousSource.source_id
            $ambiguousAdjacentSameSource = $false
            $citationRows.Add([pscustomobject][ordered]@{
                footnote_number = $footnoteNumber
                reference_id = $referenceId
                source_id = $ambiguousSourceId
                citation_classification = if ($seenSources.Contains($ambiguousSourceId)) { 'repeat' } else { 'first' }
                adjacent_same_source = [bool]$ambiguousAdjacentSameSource
                ibid_rewrite_candidate = [bool]$false
                review_status = 'review_required'
                matched_text = $logicalText
                bibliography_entry = $ambiguousSource.bibliography_entry
                identity_status = ''
                source_type = $ambiguousSource.source_type
                source_policy_key = $ambiguousSource.source_policy_key
                present_components = ''
                missing_required_components = ''
                unexpected_short_form = if ($hasShortFormMarker) { 'true' } else { 'false' }
                comparison_status = $comparisonStatus
                reference_footnote_number = ''
            })
        }
    }

    $footnoteRows.Add([pscustomobject][ordered]@{
        footnote_number = $footnoteNumber
        reference_id = $referenceId
        text = $logicalText
        match_count = $matchCount
        matched_source_ids = $matchIds -join '|'
        context_source_id = $contextSourceId
        identity_status = $identityStatus
    })

    if ($matchCount -eq 1) {
        $source = $matchedSources[0]
        $sourceId = $source.source_id
        $shortFormOnFirstUse = (-not $seenSources.Contains($sourceId)) -and $hasShortFormMarker
        $adjacentSameSource = [string]::Equals($sourceId, $previousSingleMatchSourceId, [System.StringComparison]::OrdinalIgnoreCase)
        $sourceIsFirstUse = (-not $seenSources.Contains($sourceId))
        $citationEval = Get-SourceComponentEvaluation -Source $source -IsFirstUse $sourceIsFirstUse -NormalizedText $normalizedText
        if (-not $hasSourceTypePolicies) {
            $citationEval.comparison_status = 'not_configured'
            $citationEval.present_fields = @()
            $citationEval.missing_fields = @()
            $citationEval.completeness_score = 0
        }
        $reviewStatus = if ($hasShortFormMarker -or $citationEval.missing_fields.Count -gt 0) { 'review_required' } else { '' }

        foreach ($missingField in $citationEval.missing_fields) {
            if ($citationEval.missing_metadata_fields -contains $missingField) {
                $issueRows.Add((New-IssueRow -IssueType 'bibliography_metadata_missing' -Severity 'review' -FootnoteNumber $footnoteNumber -ReferenceId $referenceId -SourceId $sourceId -Message "Required bibliography metadata field is blank: $missingField" -RelatedFootnoteNumber '' -RelatedSourceId $sourceId -Evidence "required_field=$missingField"))
            }
            else {
                $issueRows.Add((New-IssueRow -IssueType 'citation_required_component_missing' -Severity 'review' -FootnoteNumber $footnoteNumber -ReferenceId $referenceId -SourceId $sourceId -Message "Required component not present in citation text: $missingField" -RelatedFootnoteNumber '' -RelatedSourceId $sourceId -Evidence "required_field=$missingField; text=$normalizedText"))
            }
        }

        $terminalMark = Get-TerminalMark -Value $normalizedText
        $terminalExpected = Get-ExpectedTerminalMark -Language $source.language
        if (-not [string]::IsNullOrWhiteSpace($terminalExpected) -and
            -not [string]::IsNullOrWhiteSpace($terminalMark) -and
            -not [string]::Equals($terminalMark, $terminalExpected, [System.StringComparison]::Ordinal)) {
            $issueRows.Add((New-IssueRow -IssueType 'terminal_mark_mismatch' -Severity 'review' -FootnoteNumber $footnoteNumber -ReferenceId $referenceId -SourceId $sourceId -Message "Terminal mark '$terminalMark' does not match policy terminal mark '$terminalExpected'." -RelatedFootnoteNumber '' -RelatedSourceId $sourceId -Evidence "terminal_mark=$terminalMark; policy_mark=$terminalExpected"))
        }

        $citationClass = if ($sourceIsFirstUse) { 'first' } else { 'repeat' }
        $normalizedVariant = $citationEval.normalized_variant
        $punctuationSignature = $citationEval.punctuation_signature
        $completenessScore = [int]$citationEval.completeness_score

        $citationRows.Add([pscustomobject][ordered]@{
            footnote_number = $footnoteNumber
            reference_id = $referenceId
            source_id = $sourceId
            citation_classification = $citationClass
            adjacent_same_source = [bool]$adjacentSameSource
            ibid_rewrite_candidate = if (
                $adjacentSameSource -and
                ([string]::Equals($source.consecutive_same_source, 'ibid', [System.StringComparison]::OrdinalIgnoreCase) -or
                    [string]::Equals($source.consecutive_same_source, '同上', [System.StringComparison]::OrdinalIgnoreCase))
            ) { $true } else { $false }
            review_status = $reviewStatus
            matched_text = $logicalText
            bibliography_entry = $source.bibliography_entry
            identity_status = 'explicit_alias_match_review_candidate'
            source_type = $source.source_type
            source_policy_key = $source.source_policy_key
            present_components = ($citationEval.present_fields -join '|')
            missing_required_components = ($citationEval.missing_fields -join '|')
            unexpected_short_form = if ($hasShortFormMarker) { 'true' } else { 'false' }
            comparison_status = $citationEval.comparison_status
            reference_footnote_number = ''
            normalized_variant = $normalizedVariant
            punctuation_signature = $punctuationSignature
            completeness_score = $completenessScore
            terminal_mark = $terminalMark
        })
        if ($shortFormOnFirstUse) {
            $issueRows.Add((New-IssueRow -IssueType 'short_form_on_first_use' -Severity 'review' -FootnoteNumber $footnoteNumber -ReferenceId $referenceId -SourceId $sourceId -Message 'Citation appears to use short-form wording on first source use and requires review.'))
        }
        elseif ($hasShortFormMarker) {
            $issueRows.Add((New-IssueRow -IssueType 'contextual_shorthand_candidate' -Severity 'review' -FootnoteNumber $footnoteNumber -ReferenceId $referenceId -SourceId $sourceId -Message 'A context-dependent shorthand appears with an explicit alias match; identity still requires review.' -RelatedSourceId $sourceId -Evidence 'identity_status=explicit_alias_match_review_candidate'))
        }

        $key = "$sourceId|$citationClass"
        $referenceCitationCandidate = $referenceCandidateBySourceClass[$key]
        if ($null -eq $referenceCitationCandidate) {
            $currentReferenceWeight = -1
        }
        else {
            $currentReferenceWeight = [int]$referenceCitationCandidate.completeness_score
        }

        $isBetterReference = $false
        if ($completenessScore -gt $currentReferenceWeight) {
            $isBetterReference = $true
        }
        elseif ($completenessScore -eq $currentReferenceWeight) {
            $currentLength = [int][string]$referenceCitationCandidate.normalized_variant.Length
            $currentFootnote = [int][string]$referenceCitationCandidate.footnote_number
            if ($normalizedVariant.Length -gt $currentLength) {
                $isBetterReference = $true
            }
            elseif ($normalizedVariant.Length -eq $currentLength -and $footnoteNumber -lt $currentFootnote) {
                $isBetterReference = $true
            }
        }

        if ($isBetterReference) {
            $referenceCandidateBySourceClass[$key] = [pscustomobject][ordered]@{
                footnote_number = $footnoteNumber
                source_id = $sourceId
                citation_classification = $citationClass
                normalized_variant = $normalizedVariant
                punctuation_signature = $punctuationSignature
                completeness_score = $completenessScore
                length = $normalizedVariant.Length
                source_type = $source.source_type
            }
        }

        [void]$citedSources.Add($sourceId)
        [void]$seenSources.Add($sourceId)
        $previousSingleMatchSourceId = $sourceId
        $comparisonStatus = $citationEval.comparison_status
    }
    else {
        $previousSingleMatchSourceId = ''
    }
}

foreach ($citation in $citationRows) {
    if ($citation.identity_status -ne 'explicit_alias_match_review_candidate') {
        continue
    }
    $sourceId = [string]$citation.source_id
    $citationClass = [string]$citation.citation_classification
    $key = "$sourceId|$citationClass"
    $referenceCitation = $referenceCandidateBySourceClass[$key]
    $citation.reference_footnote_number = if ($null -eq $referenceCitation) { '' } else { [string]$referenceCitation.footnote_number }

    $citationVariantRows.Add([pscustomobject][ordered]@{
        footnote_number = $citation.footnote_number
        source_id = $citation.source_id
        citation_classification = $citation.citation_classification
        completeness_score = $citation.completeness_score
        present_components = $citation.present_components
        missing_required_components = $citation.missing_required_components
        normalized_variant = $citation.normalized_variant
        punctuation_signature = $citation.punctuation_signature
        terminal_mark = $citation.terminal_mark
        reference_footnote_number = $citation.reference_footnote_number
        comparison_status = $citation.comparison_status
    })
}

foreach ($key in $referenceCandidateBySourceClass.Keys) {
    $referenceCitation = $referenceCandidateBySourceClass[$key]
    foreach ($citation in @($citationRows | Where-Object {
        $_.identity_status -eq 'explicit_alias_match_review_candidate' -and "$($_.source_id)|$($_.citation_classification)" -eq $key
    })) {
        if ([string]$citation.footnote_number -eq [string]$referenceCitation.footnote_number) {
            continue
        }
        if (($citation.normalized_variant -ne $referenceCitation.normalized_variant) -or
            ($citation.punctuation_signature -ne $referenceCitation.punctuation_signature)) {
            $issueRows.Add((New-IssueRow -IssueType 'citation_variant' -Severity 'review' -FootnoteNumber $citation.footnote_number -ReferenceId $citation.reference_id -SourceId $citation.source_id -Message 'Citation variant differs from reference for source/classification.' -RelatedFootnoteNumber $referenceCitation.footnote_number -RelatedSourceId $referenceCitation.source_id -Evidence "normalized_variant=$($citation.normalized_variant); punctuation_signature=$($citation.punctuation_signature)"))
        }
    }
}

foreach ($source in $sources) {
    if (-not $citedSources.Contains($source.source_id)) {
        $issueRows.Add((New-IssueRow -IssueType 'unused_bibliography_entry' -Severity 'review' -FootnoteNumber '' -ReferenceId '' -SourceId $source.source_id -Message 'This bibliography entry was not matched to a footnote. Review only; this is not a deletion instruction.'))
    }
}

if ($bibliographyReconciliationStatus -eq 'evaluated') {
    foreach ($source in $sources) {
        if (($citedSources.Contains($source.source_id)) -and
            -not $observedBibliographySourceIds.Contains($source.source_id)) {
            $issueRows.Add((New-IssueRow -IssueType 'registry_missing_from_document_bibliography' -Severity 'review' -FootnoteNumber '' -ReferenceId '' -SourceId $source.source_id -Message 'A cited source was not found in the designated bibliography section.' -Evidence "source_id=$($source.source_id)"))
        }
    }
}

$footnotesPath = $outputPaths.footnotes
$citationsPath = $outputPaths.citations
$citationVariantPath = $outputPaths.citation_variants
$bibliographyReconPath = $outputPaths.bibliography_reconciliation
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
    'footnote_number', 'reference_id', 'text', 'match_count', 'matched_source_ids', 'context_source_id', 'identity_status'
) -Rows @($footnoteRows)
Write-CsvFile -Path $citationsPath -Columns @(
    'footnote_number',
    'reference_id',
    'source_id',
    'citation_classification',
    'adjacent_same_source',
    'ibid_rewrite_candidate',
    'review_status',
    'matched_text',
    'bibliography_entry',
    'identity_status',
    'source_type',
    'source_policy_key',
    'present_components',
    'missing_required_components',
    'unexpected_short_form',
    'comparison_status',
    'reference_footnote_number'
) -Rows @($citationRows)
Write-CsvFile -Path $citationVariantPath -Columns @(
    'footnote_number',
    'source_id',
    'citation_classification',
    'completeness_score',
    'present_components',
    'missing_required_components',
    'normalized_variant',
    'punctuation_signature',
    'terminal_mark',
    'reference_footnote_number',
    'comparison_status'
) -Rows @($citationVariantRows)
Write-CsvFile -Path $bibliographyReconPath -Columns @(
    'paragraph_number',
    'text',
    'match_count',
    'matched_source_ids',
    'status'
) -Rows @($bibliographyReconciliationRows)
Write-CsvFile -Path $issuesPath -Columns @(
    'issue_type', 'severity', 'footnote_number', 'reference_id', 'source_id', 'message', 'related_footnote_number', 'related_source_id', 'evidence'
) -Rows @($issueRows)

$inputHashAfter = (Get-FileHash -LiteralPath $resolvedInput -Algorithm SHA256).Hash.ToLowerInvariant()
if (-not [string]::Equals($inputHashBefore, $inputHashAfter, [System.StringComparison]::Ordinal)) {
    throw "Input DOCX changed during audit. Before: $inputHashBefore; after: $inputHashAfter"
}

$bibliographyDocumentHashAfter = if ($bibliographyReconciliationEnabled -and $null -ne $resolvedBibliographyDocument) {
    (Get-FileHash -LiteralPath $resolvedBibliographyDocument -Algorithm SHA256).Hash.ToLowerInvariant()
} else {
    ''
}
if ($bibliographyReconciliationEnabled -and -not [string]::Equals($bibliographyDocumentHashBefore, $bibliographyDocumentHashAfter, [System.StringComparison]::Ordinal)) {
    throw "Bibliography DOCX changed during audit. Before: $bibliographyDocumentHashBefore; after: $bibliographyDocumentHashAfter"
}

$summary = [pscustomobject][ordered]@{
    schema_version = 2
    tool_name = 'thesis-footnote-normalizer-audit'
    mode = 'audit_only'
    input_file_name = [System.IO.Path]::GetFileName($resolvedInput)
    input_sha256_before = $inputHashBefore
    input_sha256_after = $inputHashAfter
    input_unchanged = $true
    policy = $policy
    bibliography_reconciliation = [pscustomobject][ordered]@{
        status = $bibliographyReconciliationStatus
        input_file_name = $bibliographyReconciliationInputFile
        input_sha256_before = $bibliographyDocumentHashBefore
        input_sha256_after = $bibliographyDocumentHashAfter
        input_unchanged = if ($bibliographyReconciliationEnabled) { [string]::Equals($bibliographyDocumentHashBefore, $bibliographyDocumentHashAfter, [System.StringComparison]::Ordinal) } else { $false }
        paragraphs = $bibliographyReconciliationParagraphCount
    }
    counts = [pscustomobject][ordered]@{
        footnotes = $footnoteRows.Count
        citations = $citationRows.Count
        issues = $issueRows.Count
        bibliography_reconciliation_rows = $bibliographyReconciliationRows.Count
        unmatched_footnotes = @($issueRows | Where-Object issue_type -eq 'unmatched_footnote').Count
        multiple_source_matches = @($issueRows | Where-Object issue_type -eq 'multiple_source_matches').Count
        missing_footnote_bodies = @($issueRows | Where-Object issue_type -eq 'missing_footnote_body').Count
        unused_bibliography_entries = @($issueRows | Where-Object issue_type -eq 'unused_bibliography_entry').Count
        source_type_policy_missing = @($issueRows | Where-Object issue_type -eq 'source_type_policy_missing').Count
        bibliography_metadata_missing = @($issueRows | Where-Object issue_type -eq 'bibliography_metadata_missing').Count
        citation_required_component_missing = @($issueRows | Where-Object issue_type -eq 'citation_required_component_missing').Count
        citation_variant = @($issueRows | Where-Object issue_type -eq 'citation_variant').Count
        terminal_mark_mismatch = @($issueRows | Where-Object issue_type -eq 'terminal_mark_mismatch').Count
        document_bibliography_unmatched = @($issueRows | Where-Object issue_type -eq 'document_bibliography_unmatched').Count
        document_bibliography_multiple_matches = @($issueRows | Where-Object issue_type -eq 'document_bibliography_multiple_matches').Count
        document_bibliography_not_found = @($issueRows | Where-Object issue_type -eq 'document_bibliography_not_found').Count
        registry_missing_from_document_bibliography = @($issueRows | Where-Object issue_type -eq 'registry_missing_from_document_bibliography').Count
        citation_variants = $citationVariantRows.Count
    }
    limitations = @(
        'Version 2 adds citation variant tracking and source-type policy component checks.',
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
    '- `first` is the initial match in the doc for a source; `repeat` is any later single match.',
    '- Unmatched and multiple-match footnotes require human review.',
    '- `footnote_number` is document-reference order, not a reconstruction of numbering rendered by Word.',
    '- Unused bibliography entries require review and are not deletion instructions.',
    '- Document bibliography matching is review-only; it validates only marker-scoped presence and alias coverage, not truth of the cited material.',
    '- A bibliography `marker_not_found` status means no matching marker was detected under the configured policy, not an absence proof of a bibliography in the source document.',
    '- Version 2 adds citation variant normalization and explicit component/terminal-mark checks.',
    '- Version 2 does not rewrite the DOCX or generate formatted replacement citations.'
)
Set-Content -LiteralPath $reportPath -Value $reportLines -Encoding utf8

$finalInputHash = (Get-FileHash -LiteralPath $resolvedInput -Algorithm SHA256).Hash.ToLowerInvariant()
if (-not [string]::Equals($inputHashBefore, $finalInputHash, [System.StringComparison]::Ordinal)) {
    throw "Input DOCX changed during audit. Before: $inputHashBefore; final: $finalInputHash"
}

if ($bibliographyReconciliationEnabled -and $null -ne $resolvedBibliographyDocument) {
    $finalBibliographyHash = (Get-FileHash -LiteralPath $resolvedBibliographyDocument -Algorithm SHA256).Hash.ToLowerInvariant()
    if (-not [string]::Equals($bibliographyDocumentHashBefore, $finalBibliographyHash, [System.StringComparison]::Ordinal)) {
        throw "Bibliography DOCX changed during audit. Before: $bibliographyDocumentHashBefore; final: $finalBibliographyHash"
    }
}

[pscustomobject][ordered]@{
    output_directory = $resolvedOutput
    footnotes = $footnoteRows.Count
    citations = $citationRows.Count
    issues = $issueRows.Count
    input_unchanged = $true
}
