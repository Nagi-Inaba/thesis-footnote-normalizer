[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputPath,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not [string]::Equals([System.IO.Path]::GetExtension($OutputPath), '.docx', [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "OutputPath must have a .docx extension: $OutputPath"
}

$resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
$parentDirectory = Split-Path -Parent $resolvedOutput
if (-not (Test-Path -LiteralPath $parentDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $parentDirectory | Out-Null
}

if (Test-Path -LiteralPath $resolvedOutput) {
    if (-not $Force) {
        throw "Synthetic DOCX already exists: $resolvedOutput"
    }
    if (-not (Test-Path -LiteralPath $resolvedOutput -PathType Leaf)) {
        throw "OutputPath exists but is not a file: $resolvedOutput"
    }
    Remove-Item -LiteralPath $resolvedOutput
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Add-ZipTextEntry {
    param(
        [Parameter(Mandatory)]
        [System.IO.Compression.ZipArchive]$Archive,

        [Parameter(Mandatory)]
        [string]$EntryName,

        [Parameter(Mandatory)]
        [string]$Content
    )

    $entry = $Archive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $entry.LastWriteTime = [System.DateTimeOffset]::new(1980, 1, 1, 0, 0, 0, [System.TimeSpan]::Zero)
    $stream = $entry.Open()
    $writer = $null
    try {
        $utf8WithoutBom = [System.Text.UTF8Encoding]::new($false)
        $writer = [System.IO.StreamWriter]::new($stream, $utf8WithoutBom)
        $writer.Write($Content)
    }
    finally {
        if ($null -ne $writer) {
            $writer.Dispose()
        }
        else {
            $stream.Dispose()
        }
    }
}

$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>
'@

$packageRelationships = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@

$documentRelationships = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>
'@

$documentXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Alpha statement</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>
    <w:p><w:r><w:t>Adjacent alpha statement</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>
    <w:p><w:r><w:t>Beta statement</w:t></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>
    <w:p><w:r><w:t>Later alpha statement</w:t></w:r><w:r><w:footnoteReference w:id="4"/></w:r></w:p>
    <w:p><w:r><w:t>Explanatory statement</w:t></w:r><w:r><w:footnoteReference w:id="5"/></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>
'@

$footnotesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:t>Aster Vale, Clockwork Orchards, p. 11.</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>Aster</w:t><w:tab/><w:t>Vale, Clockwork Orchards, p. 18.</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="3"><w:p><w:r><w:t>Beryl North, Lantern Rivers, p. 7.</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="4"><w:p><w:r><w:t>Aster Vale,</w:t><w:br/><w:t>Clockwork Orchards, p. 29.</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="5"><w:p><w:r><w:t>=1+1 Artificial examples contain no bibliography alias.</w:t></w:r></w:p></w:footnote>
</w:footnotes>
'@

$fileStream = [System.IO.File]::Open($resolvedOutput, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
$archive = $null
try {
    $archive = [System.IO.Compression.ZipArchive]::new($fileStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)
    Add-ZipTextEntry -Archive $archive -EntryName '[Content_Types].xml' -Content $contentTypes
    Add-ZipTextEntry -Archive $archive -EntryName '_rels/.rels' -Content $packageRelationships
    Add-ZipTextEntry -Archive $archive -EntryName 'word/document.xml' -Content $documentXml
    Add-ZipTextEntry -Archive $archive -EntryName 'word/_rels/document.xml.rels' -Content $documentRelationships
    Add-ZipTextEntry -Archive $archive -EntryName 'word/footnotes.xml' -Content $footnotesXml
}
finally {
    if ($null -ne $archive) {
        $archive.Dispose()
    }
    else {
        $fileStream.Dispose()
    }
}

$signature = [System.IO.File]::ReadAllBytes($resolvedOutput)[0..1]
if ($signature[0] -ne 0x50 -or $signature[1] -ne 0x4B) {
    throw "Synthetic DOCX is not a ZIP-based OOXML package: $resolvedOutput"
}

Get-Item -LiteralPath $resolvedOutput
