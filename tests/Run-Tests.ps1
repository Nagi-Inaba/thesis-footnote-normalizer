[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-True {
    param(
        [Parameter(Mandatory)]
        [bool]$Condition,

        [Parameter(Mandatory)]
        [string]$Message
    )

    if (-not $Condition) {
        throw "Assertion failed: $Message"
    }
}

function Assert-Equal {
    param(
        [AllowNull()]$Expected,
        [AllowNull()]$Actual,
        [Parameter(Mandatory)]
        [string]$Message
    )

    if ($Expected -ne $Actual) {
        throw "Assertion failed: $Message. Expected '$Expected'; actual '$Actual'."
    }
}

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$fixtureGenerator = Join-Path $PSScriptRoot 'New-SyntheticDocx.ps1'
$auditScript = Join-Path $repositoryRoot 'scripts\Invoke-FootnoteAudit.ps1'
$tempBase = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
$testRoot = Join-Path $tempBase ("thesis-footnote-normalizer-tests-{0}" -f [guid]::NewGuid().ToString('N'))
$testRoot = [System.IO.Path]::GetFullPath($testRoot)

$tempPrefix = $tempBase.TrimEnd(
    [System.IO.Path]::DirectorySeparatorChar,
    [System.IO.Path]::AltDirectorySeparatorChar
) + [System.IO.Path]::DirectorySeparatorChar
if (-not $testRoot.StartsWith($tempPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or
    -not ([System.IO.Path]::GetFileName($testRoot)).StartsWith('thesis-footnote-normalizer-tests-', [System.StringComparison]::Ordinal)) {
    throw "Refusing to use an unvalidated test directory: $testRoot"
}

New-Item -ItemType Directory -Path $testRoot | Out-Null

try {
    $fixturePath = Join-Path $testRoot 'synthetic-footnotes.docx'
    $bibliographyPath = Join-Path $testRoot 'bibliography.csv'
    $policyPath = Join-Path $testRoot 'policy.json'
    $outputOne = Join-Path $testRoot 'audit-one'
    $outputTwo = Join-Path $testRoot 'audit-two'
    $outputAmbiguous = Join-Path $testRoot 'audit-ambiguous'

    & $fixtureGenerator -OutputPath $fixturePath | Out-Null

    $bibliographyRows = @(
        [pscustomobject][ordered]@{
            source_id = 'SRC-A'
            source_type = 'fictional_book'
            language = 'en'
            author = 'Aster Vale'
            short_title = 'Clockwork Orchards'
            aliases = 'Aster Vale|Clockwork Orchards'
            bibliography_entry = 'Aster Vale. Clockwork Orchards. Imaginary Press, 2042.'
        },
        [pscustomobject][ordered]@{
            source_id = 'SRC-B'
            source_type = 'fictional_book'
            language = 'en'
            author = 'Beryl North'
            short_title = 'Lantern Rivers'
            aliases = 'Beryl North|Lantern Rivers'
            bibliography_entry = 'Beryl North. Lantern Rivers. Fable House, 2043.'
        },
        [pscustomobject][ordered]@{
            source_id = 'SRC-UNUSED'
            source_type = 'fictional_article'
            language = 'en'
            author = 'Cinder Quill'
            short_title = 'Paper Moons'
            aliases = 'Art'
            bibliography_entry = 'Cinder Quill. Paper Moons. Invented Review 9, 2044.'
        }
    )
    Set-Content -LiteralPath $bibliographyPath -Value ($bibliographyRows | ConvertTo-Csv -NoTypeInformation) -Encoding utf8

    $policy = [pscustomobject][ordered]@{
        policy_name = 'synthetic-review-policy'
        subsequent_citation = 'short_form_review'
        consecutive_same_source = 'mark_ibid_candidate'
        new_sources = 'human_review_required'
    }
    Set-Content -LiteralPath $policyPath -Value ($policy | ConvertTo-Json -Depth 5) -Encoding utf8

    $hashBefore = (Get-FileHash -LiteralPath $fixturePath -Algorithm SHA256).Hash
    & $auditScript -InputDocx $fixturePath -BibliographyCsv $bibliographyPath -PolicyJson $policyPath -OutputDirectory $outputOne | Out-Null
    $hashAfterFirstRun = (Get-FileHash -LiteralPath $fixturePath -Algorithm SHA256).Hash

    $expectedFiles = @('footnotes.csv', 'citations.csv', 'issues.csv', 'summary.json', 'report.md')
    foreach ($fileName in $expectedFiles) {
        $path = Join-Path $outputOne $fileName
        Assert-True -Condition (Test-Path -LiteralPath $path -PathType Leaf) -Message "output exists: $fileName"
        Assert-True -Condition ((Get-Item -LiteralPath $path).Length -gt 0) -Message "output is non-empty: $fileName"
    }

    $footnotes = @(Import-Csv -LiteralPath (Join-Path $outputOne 'footnotes.csv') -Encoding utf8)
    $citations = @(Import-Csv -LiteralPath (Join-Path $outputOne 'citations.csv') -Encoding utf8)
    $issues = @(Import-Csv -LiteralPath (Join-Path $outputOne 'issues.csv') -Encoding utf8)
    $summary = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $outputOne 'summary.json') | ConvertFrom-Json
    $report = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $outputOne 'report.md')

    Assert-Equal -Expected 5 -Actual $footnotes.Count -Message 'five true footnotes are audited'
    Assert-Equal -Expected 4 -Actual $citations.Count -Message 'four source-matched citation rows are emitted'
    Assert-Equal -Expected 'first' -Actual (@($citations | Where-Object footnote_number -eq '1')[0].citation_classification) -Message 'Source A first use classification'
    Assert-Equal -Expected 'ibid_candidate' -Actual (@($citations | Where-Object footnote_number -eq '2')[0].citation_classification) -Message 'adjacent Source A classification'
    Assert-Equal -Expected 'first' -Actual (@($citations | Where-Object footnote_number -eq '3')[0].citation_classification) -Message 'Source B first use classification'
    Assert-Equal -Expected 'repeat' -Actual (@($citations | Where-Object footnote_number -eq '4')[0].citation_classification) -Message 'nonconsecutive Source A classification'
    Assert-Equal -Expected 1 -Actual @($issues | Where-Object issue_type -eq 'unmatched_footnote').Count -Message 'one unmatched footnote issue'
    Assert-Equal -Expected 1 -Actual @($issues | Where-Object issue_type -eq 'unused_bibliography_entry').Count -Message 'one unused bibliography issue'
    Assert-True -Condition ($footnotes[4].text.StartsWith("'=")) -Message 'spreadsheet-formula-like footnote text is neutralized in CSV'
    Assert-True -Condition ($report.Contains('not claims of bibliographic correctness')) -Message 'report states bibliographic limitation'
    Assert-True -Condition ([bool]$summary.input_unchanged) -Message 'summary records unchanged input'
    Assert-Equal -Expected 5 -Actual $summary.counts.footnotes -Message 'summary parses and records footnote count'
    Assert-Equal -Expected $hashBefore -Actual $hashAfterFirstRun -Message 'input hash remains unchanged after first audit'
    Assert-Equal -Expected $hashBefore.ToLowerInvariant() -Actual $summary.input_sha256_before -Message 'summary before hash'
    Assert-Equal -Expected $hashBefore.ToLowerInvariant() -Actual $summary.input_sha256_after -Message 'summary after hash'

    $existingDirectoryRejected = $false
    try {
        & $auditScript -InputDocx $fixturePath -BibliographyCsv $bibliographyPath -PolicyJson $policyPath -OutputDirectory $outputOne | Out-Null
    }
    catch {
        $existingDirectoryRejected = $_.Exception.Message -like '*already exists*'
    }
    Assert-True -Condition $existingDirectoryRejected -Message 'existing output directory is rejected without Force'

    $inputCollisionRejected = $false
    try {
        & $auditScript -InputDocx $fixturePath -BibliographyCsv $bibliographyPath -PolicyJson (Join-Path $outputOne 'summary.json') -OutputDirectory $outputOne -Force | Out-Null
    }
    catch {
        $inputCollisionRejected = $_.Exception.Message -like '*overwrite an input file*'
    }
    Assert-True -Condition $inputCollisionRejected -Message 'Force cannot overwrite a policy or bibliography input with an audit output'

    $externalHardlinkTarget = Join-Path $testRoot 'hardlink-target.txt'
    Set-Content -LiteralPath $externalHardlinkTarget -Value 'must remain unchanged' -Encoding utf8
    $linkedOutput = Join-Path $outputOne 'footnotes.csv'
    Remove-Item -LiteralPath $linkedOutput -Force
    New-Item -ItemType HardLink -Path $linkedOutput -Target $externalHardlinkTarget | Out-Null
    & $auditScript -InputDocx $fixturePath -BibliographyCsv $bibliographyPath -PolicyJson $policyPath -OutputDirectory $outputOne -Force | Out-Null
    Assert-Equal -Expected 'must remain unchanged' -Actual ((Get-Content -LiteralPath $externalHardlinkTarget -Raw -Encoding utf8).Trim()) -Message 'Force replaces a hardlinked output entry without modifying the other link target'
    Assert-Equal -Expected 5 -Actual @(Import-Csv -LiteralPath $linkedOutput -Encoding utf8).Count -Message 'hardlinked output entry is replaced by a fresh audit file'

    $junctionTarget = Join-Path $testRoot 'junction-target'
    $junctionOutput = Join-Path $testRoot 'junction-output'
    New-Item -ItemType Directory -Path $junctionTarget | Out-Null
    New-Item -ItemType Junction -Path $junctionOutput -Target $junctionTarget | Out-Null
    $junctionRejected = $false
    try {
        & $auditScript -InputDocx $fixturePath -BibliographyCsv $bibliographyPath -PolicyJson $policyPath -OutputDirectory $junctionOutput -Force | Out-Null
    }
    catch {
        $junctionRejected = $_.Exception.Message -like '*reparse point*'
    }
    Assert-True -Condition $junctionRejected -Message 'reparse-point OutputDirectory is rejected'
    Remove-Item -LiteralPath $junctionOutput -Force

    & $auditScript -InputDocx $fixturePath -BibliographyCsv $bibliographyPath -PolicyJson $policyPath -OutputDirectory $outputTwo | Out-Null
    $hashAfterSecondRun = (Get-FileHash -LiteralPath $fixturePath -Algorithm SHA256).Hash
    Assert-Equal -Expected $hashBefore -Actual $hashAfterSecondRun -Message 'input hash remains unchanged after second audit'

    foreach ($fileName in @('footnotes.csv', 'citations.csv', 'issues.csv', 'summary.json', 'report.md')) {
        $first = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $outputOne $fileName)
        $second = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $outputTwo $fileName)
        Assert-Equal -Expected $first -Actual $second -Message "deterministic semantic output: $fileName"
    }

    $ambiguousRows = @($bibliographyRows) + [pscustomobject][ordered]@{
        source_id = 'SRC-AMBIGUOUS'
        source_type = 'fictional_book'
        language = 'en'
        author = 'Dapple Reed'
        short_title = 'Eleven Pages'
        aliases = 'p. 11.'
        bibliography_entry = 'Dapple Reed. Eleven Pages. Imaginary House, 2045.'
    }
    $ambiguousBibliographyPath = Join-Path $testRoot 'bibliography-ambiguous.csv'
    Set-Content -LiteralPath $ambiguousBibliographyPath -Value ($ambiguousRows | ConvertTo-Csv -NoTypeInformation) -Encoding utf8
    & $auditScript -InputDocx $fixturePath -BibliographyCsv $ambiguousBibliographyPath -PolicyJson $policyPath -OutputDirectory $outputAmbiguous | Out-Null
    $ambiguousCitations = @(Import-Csv -LiteralPath (Join-Path $outputAmbiguous 'citations.csv') -Encoding utf8)
    Assert-Equal -Expected 2 -Actual @($ambiguousCitations | Where-Object footnote_number -eq '1').Count -Message 'multi-match footnote emits one row per candidate'
    Assert-True -Condition (@($ambiguousCitations | Where-Object footnote_number -eq '1' | Where-Object review_status -eq 'review_required').Count -eq 2) -Message 'multi-match rows require review'
    Assert-Equal -Expected 'first' -Actual (@($ambiguousCitations | Where-Object footnote_number -eq '2' | Where-Object source_id -eq 'SRC-A')[0].citation_classification) -Message 'ambiguous prior match does not advance confirmed source state'

    $nonDocxRejected = $false
    try {
        & $auditScript -InputDocx (Join-Path $testRoot 'not-a-docx.txt') -BibliographyCsv $bibliographyPath -PolicyJson $policyPath -OutputDirectory (Join-Path $testRoot 'invalid-extension-output') | Out-Null
    }
    catch {
        $nonDocxRejected = $_.Exception.Message -like '*.docx extension*'
    }
    Assert-True -Condition $nonDocxRejected -Message 'non-DOCX input is rejected'

    Write-Output 'PASS: 5 footnotes, 4 citations, 2 review issues, deterministic outputs, and unchanged input hash.'
}
finally {
    $validatedRoot = [System.IO.Path]::GetFullPath($testRoot)
    if (-not $validatedRoot.StartsWith($tempPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or
        -not ([System.IO.Path]::GetFileName($validatedRoot)).StartsWith('thesis-footnote-normalizer-tests-', [System.StringComparison]::Ordinal)) {
        throw "Refusing to clean an unvalidated test directory: $validatedRoot"
    }
    if (Test-Path -LiteralPath $validatedRoot) {
        Remove-Item -LiteralPath $validatedRoot -Recurse -Force
    }
}
