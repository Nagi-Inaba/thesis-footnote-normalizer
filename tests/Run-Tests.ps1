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
    $policyWithBibliographyPath = Join-Path $testRoot 'policy-with-bibliography.json'
    $outputOne = Join-Path $testRoot 'audit-one'
    $outputTwo = Join-Path $testRoot 'audit-two'
    $outputAmbiguous = Join-Path $testRoot 'audit-ambiguous'
    $outputBibliography = Join-Path $testRoot 'audit-bibliography'
    $outputBibliographyMissing = Join-Path $testRoot 'audit-bibliography-missing'
    $outputBibliographyAmbiguous = Join-Path $testRoot 'audit-bibliography-ambiguous'
    $outputNoPolicyReconciliation = Join-Path $testRoot 'audit-no-policy-reconciliation'
    $bibliographyDocxPath = Join-Path $testRoot 'synthetic-bibliography.docx'
    $bibliographyDocxPathNoMarker = Join-Path $testRoot 'synthetic-bibliography-no-marker.docx'
    $bibliographyAmbiguousPath = Join-Path $testRoot 'bibliography-reconciliation-ambiguous.csv'
    $duplicateDocumentPath = Join-Path $testRoot 'duplicate-document-part.docx'
    $duplicateFootnotesPath = Join-Path $testRoot 'duplicate-footnotes-part.docx'
    $duplicateBibliographyPath = Join-Path $testRoot 'duplicate-bibliography-part.docx'
    $excessiveEntriesPath = Join-Path $testRoot 'excessive-entries.docx'

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
            source_id = 'SRC-C'
            source_type = 'fictional_book'
            language = 'en'
            author = 'Cedar Vane'
            short_title = 'Ember Journal'
            aliases = 'Cedar Vane|Ember Journal'
            bibliography_entry = 'Cedar Vane. Ember Journal. Hearthstone Press, 2045.'
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
        consecutive_same_source = 'short-form'
        new_sources = 'human_review_required'
    }
    Set-Content -LiteralPath $policyPath -Value ($policy | ConvertTo-Json -Depth 5) -Encoding utf8

    $policyWithBibliography = [pscustomobject][ordered]@{
        policy_name = 'synthetic-bibliography-review-policy'
        subsequent_citation = 'short_form_review'
        consecutive_same_source = 'short-form'
        new_sources = 'human_review_required'
        bibliography_document = [ordered]@{
            enabled = $true
            start_marker = 'References'
            end_marker = ''
            include_heading = $false
            paragraph_match_mode = 'exact'
        }
    }
    Set-Content -LiteralPath $policyWithBibliographyPath -Value ($policyWithBibliography | ConvertTo-Json -Depth 8) -Encoding utf8

    & $fixtureGenerator -OutputPath $bibliographyDocxPath -IncludeBibliographySection | Out-Null
    & $fixtureGenerator -OutputPath $bibliographyDocxPathNoMarker -IncludeBibliographySection -OmitBibliographyMarker | Out-Null
    & $fixtureGenerator -OutputPath $duplicateDocumentPath -DuplicatePart document | Out-Null
    & $fixtureGenerator -OutputPath $duplicateFootnotesPath -DuplicatePart footnotes | Out-Null
    & $fixtureGenerator -OutputPath $duplicateBibliographyPath -IncludeBibliographySection -DuplicatePart document | Out-Null
    & $fixtureGenerator -OutputPath $excessiveEntriesPath -ExtraEntryCount 4097 | Out-Null

    $hashBefore = (Get-FileHash -LiteralPath $fixturePath -Algorithm SHA256).Hash
    & $auditScript -InputDocx $fixturePath -BibliographyCsv $bibliographyPath -PolicyJson $policyPath -OutputDirectory $outputOne | Out-Null
    $bibliographyFixtureHashBefore = (Get-FileHash -LiteralPath $bibliographyDocxPath -Algorithm SHA256).Hash
    $hashAfterFirstRun = (Get-FileHash -LiteralPath $fixturePath -Algorithm SHA256).Hash

    $expectedFiles = @('footnotes.csv', 'citations.csv', 'citation-variants.csv', 'issues.csv', 'summary.json', 'report.md', 'bibliography-reconciliation.csv')
    foreach ($fileName in $expectedFiles) {
        $path = Join-Path $outputOne $fileName
        Assert-True -Condition (Test-Path -LiteralPath $path -PathType Leaf) -Message "output exists: $fileName"
        Assert-True -Condition ((Get-Item -LiteralPath $path).Length -gt 0) -Message "output is non-empty: $fileName"
    }

    $footnotes = @(Import-Csv -LiteralPath (Join-Path $outputOne 'footnotes.csv') -Encoding utf8)
    $citations = @(Import-Csv -LiteralPath (Join-Path $outputOne 'citations.csv') -Encoding utf8)
    $bibliographyReconciliation = @(Import-Csv -LiteralPath (Join-Path $outputOne 'bibliography-reconciliation.csv') -Encoding utf8)
    $issues = @(Import-Csv -LiteralPath (Join-Path $outputOne 'issues.csv') -Encoding utf8)
    $summary = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $outputOne 'summary.json') | ConvertFrom-Json
    $report = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $outputOne 'report.md')

    Assert-Equal -Expected 8 -Actual $footnotes.Count -Message 'eight true footnotes are audited'
    Assert-Equal -Expected 6 -Actual $citations.Count -Message 'six source-matched citation rows are emitted'
    Assert-Equal -Expected 'first' -Actual (@($citations | Where-Object footnote_number -eq '1')[0].citation_classification) -Message 'Source A first use classification'
    Assert-Equal -Expected 'repeat' -Actual (@($citations | Where-Object footnote_number -eq '2')[0].citation_classification) -Message 'adjacent Source A classification'
    Assert-Equal -Expected 'first' -Actual (@($citations | Where-Object footnote_number -eq '3')[0].citation_classification) -Message 'Source B first use classification'
    Assert-Equal -Expected 'first' -Actual (@($citations | Where-Object footnote_number -eq '4')[0].citation_classification) -Message 'new source first-use classification'
    Assert-Equal -Expected 'true' -Actual ([string](@($citations | Where-Object footnote_number -eq '2')[0].adjacent_same_source)) -Message 'adjacent Source A classification tracks adjacency'
    Assert-Equal -Expected 'false' -Actual ([string](@($citations | Where-Object footnote_number -eq '2')[0].ibid_rewrite_candidate)) -Message 'ibid_rewrite_candidate defaults false for current policy'
    Assert-Equal -Expected 1 -Actual @($issues | Where-Object issue_type -eq 'short_form_on_first_use').Count -Message 'first-use short-form citation is flagged'
    Assert-Equal -Expected 2 -Actual @($issues | Where-Object issue_type -eq 'contextual_shorthand_candidate').Count -Message 'contextual shorthand candidates are flagged for inferred and explicit matches'
    Assert-Equal -Expected 0 -Actual @($issues | Where-Object issue_type -eq 'unresolved_shorthand').Count -Message 'contextual shorthand does not become unresolved'
    Assert-Equal -Expected 'SRC-C' -Actual (@($issues | Where-Object issue_type -eq 'contextual_shorthand_candidate')[0].source_id) -Message 'contextual shorthand issue records candidate source id'
    Assert-Equal -Expected 'context_inferred_review_required' -Actual (@($footnotes | Where-Object footnote_number -eq '5')[0].identity_status) -Message 'contextual shorthand marks identity status in footnotes output'
    Assert-Equal -Expected 'review_required' -Actual (@($citations | Where-Object footnote_number -eq '4')[0].review_status) -Message 'short-form first-use citation is review-required'
    Assert-Equal -Expected 'review_required' -Actual (@($citations | Where-Object footnote_number -eq '8')[0].review_status) -Message 'explicit repeat shorthand remains review-required'
    Assert-Equal -Expected 1 -Actual @($citations | Where-Object footnote_number -eq '7' | Where-Object source_id -eq 'SRC-A').Count -Message 'w:type normal footnote remains in citation analysis'
    Assert-Equal -Expected 0 -Actual @($citations | Where-Object footnote_number -eq '5').Count -Message 'bare same-up shorthand emits no citation row'
    Assert-Equal -Expected 1 -Actual @($issues | Where-Object issue_type -eq 'unmatched_footnote').Count -Message 'plain unmatched footnotes are explicitly reported as unmatched_footnote'
    $formulaFootnote = @($footnotes | Where-Object footnote_number -eq '6')
    Assert-Equal -Expected 1 -Actual $formulaFootnote.Count -Message 'formula-like synthetic footnote is emitted'
    Assert-Equal -Expected '=2+2 Flibidem Studies' -Actual $formulaFootnote[0].text.TrimStart("'") -Message 'unmatched synthetic footnote text preserved before formula neutralization'
    Assert-Equal -Expected '''=2+2 Flibidem Studies' -Actual $formulaFootnote[0].text -Message 'formula-like footnote text is formula-neutralized in CSV'
    Assert-Equal -Expected 0 -Actual @($issues | Where-Object issue_type -eq 'unresolved_shorthand' | Where-Object footnote_number -eq '6').Count -Message 'latin shorthand is not detected inside a longer word'
    Assert-Equal -Expected 0 -Actual @($issues | Where-Object issue_type -eq 'document_bibliography_not_found').Count -Message 'base run does not produce bibliography-not-found issues'
    Assert-Equal -Expected 1 -Actual @($issues | Where-Object issue_type -eq 'unused_bibliography_entry').Count -Message 'one unused bibliography issue'
    Assert-Equal -Expected 'false' -Actual ([string](@($citations | Where-Object footnote_number -eq '4')[0].adjacent_same_source)) -Message 'Source C first use is not adjacent to same source'
    Assert-Equal -Expected 0 -Actual @($issues | Where-Object issue_type -eq 'citation_variant' | Where-Object footnote_number -eq '7').Count -Message 'page-only differences do not create citation variant flags'
    Assert-Equal -Expected 1 -Actual @($issues | Where-Object issue_type -eq 'citation_variant' | Where-Object footnote_number -eq '8').Count -Message 'explicit shorthand variant is compared with the repeat reference'
    Assert-True -Condition ($report.Contains('not claims of bibliographic correctness')) -Message 'report states bibliographic limitation'
    Assert-True -Condition ([bool]$summary.input_unchanged) -Message 'summary records unchanged input'
    Assert-Equal -Expected 8 -Actual $summary.counts.footnotes -Message 'summary parses and records footnote count'
    Assert-Equal -Expected 2 -Actual $summary.schema_version -Message 'schema version is updated'
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
    Assert-Equal -Expected 8 -Actual @(Import-Csv -LiteralPath $linkedOutput -Encoding utf8).Count -Message 'hardlinked output entry is replaced by a fresh audit file'

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

    foreach ($fileName in @('footnotes.csv', 'citations.csv', 'citation-variants.csv', 'issues.csv', 'summary.json', 'report.md', 'bibliography-reconciliation.csv')) {
        $first = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $outputOne $fileName)
        $second = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $outputTwo $fileName)
        Assert-Equal -Expected $first -Actual $second -Message "deterministic semantic output: $fileName"
    }

    & $auditScript -InputDocx $fixturePath -BibliographyCsv $bibliographyPath -PolicyJson $policyWithBibliographyPath -BibliographyDocx $bibliographyDocxPath -OutputDirectory $outputBibliography | Out-Null
    $bibliographyReconciliation = @(Import-Csv -LiteralPath (Join-Path $outputBibliography 'bibliography-reconciliation.csv') -Encoding utf8)
    $bibliographyIssues = @(Import-Csv -LiteralPath (Join-Path $outputBibliography 'issues.csv') -Encoding utf8)
    $bibliographySummary = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $outputBibliography 'summary.json') | ConvertFrom-Json
    $citationVariants = @(Import-Csv -LiteralPath (Join-Path $outputBibliography 'citation-variants.csv') -Encoding utf8)

    Assert-Equal -Expected 2 -Actual $bibliographyReconciliation.Count -Message 'bibliography reconciliation extracts one matched and one unmatched paragraph'
    Assert-Equal -Expected 1 -Actual @($bibliographyReconciliation | Where-Object status -eq 'matched').Count -Message 'one bibliography paragraph is matched'
    Assert-Equal -Expected 1 -Actual @($bibliographyReconciliation | Where-Object status -eq 'unmatched').Count -Message 'one bibliography paragraph is unmatched'
    Assert-Equal -Expected 1 -Actual @($bibliographyIssues | Where-Object issue_type -eq 'document_bibliography_unmatched').Count -Message 'one document bibliography paragraph is unmatched'
    $registryMissing = @($bibliographyIssues | Where-Object issue_type -eq 'registry_missing_from_document_bibliography')
    Assert-Equal -Expected 2 -Actual $registryMissing.Count -Message 'cited sources absent from bibliography produce registry_missing issues'
    Assert-True -Condition (@($registryMissing | Where-Object source_id -eq 'SRC-B').Count -eq 1) -Message 'missing SRC-B is reported'
    Assert-True -Condition (@($registryMissing | Where-Object source_id -eq 'SRC-C').Count -eq 1) -Message 'missing SRC-C is reported'
    Assert-True -Condition (@($registryMissing | Where-Object source_id -eq 'SRC-UNUSED').Count -eq 0) -Message 'unused source is not flagged by registry_missing'
    Assert-Equal -Expected 0 -Actual @($bibliographyIssues | Where-Object issue_type -eq 'document_bibliography_not_found').Count -Message 'bibliography section start marker is found'
    Assert-Equal -Expected 2 -Actual $bibliographyReconciliation.Count -Message 'bibliography reconciliation yields two scoped rows'
    Assert-Equal -Expected 'evaluated' -Actual $bibliographySummary.bibliography_reconciliation.status -Message 'bibliography reconciliation status is evaluated with enabled policy and heading excluded'
    Assert-Equal -Expected 2 -Actual $bibliographySummary.counts.bibliography_reconciliation_rows -Message 'summary counts reconciliation rows'
    Assert-Equal -Expected $bibliographyFixtureHashBefore.ToLowerInvariant() -Actual $bibliographySummary.bibliography_reconciliation.input_sha256_before -Message 'summary includes bibliography input before hash'
    Assert-Equal -Expected $bibliographyFixtureHashBefore.ToLowerInvariant() -Actual $bibliographySummary.bibliography_reconciliation.input_sha256_after -Message 'summary includes bibliography input after hash'

    $bibliographyVariantFirst = @($citationVariants | Where-Object footnote_number -eq '1')[0].normalized_variant
    $bibliographyVariantSecond = @($citationVariants | Where-Object footnote_number -eq '2')[0].normalized_variant
    Assert-Equal -Expected $bibliographyVariantFirst -Actual $bibliographyVariantSecond -Message 'page-variant normalization makes citation variants consistent across pages'

    $bibliographyAmbiguousRows = @(
        foreach ($row in $bibliographyRows) {
            [pscustomobject][ordered]@{
                source_id = $row.source_id
                source_type = $row.source_type
                language = $row.language
                author = $row.author
                short_title = $row.short_title
                aliases = $row.aliases
                bibliography_entry = $row.bibliography_entry
                bibliography_aliases = $row.aliases
            }
        }
    ) + [pscustomobject][ordered]@{
        source_id = 'SRC-A-ALT'
        source_type = 'fictional_book'
        language = 'en'
        author = 'Alternate Author'
        short_title = 'Alternate Orchard'
        aliases = 'Never Appears in Footnotes'
        bibliography_entry = 'Alternate Author. Alternate Orchard. Fictional Press, 2046.'
        bibliography_aliases = 'Aster Vale'
    }
    Set-Content -LiteralPath $bibliographyAmbiguousPath -Value ($bibliographyAmbiguousRows | ConvertTo-Csv -NoTypeInformation) -Encoding utf8
    & $auditScript -InputDocx $fixturePath -BibliographyCsv $bibliographyAmbiguousPath -PolicyJson $policyWithBibliographyPath -BibliographyDocx $bibliographyDocxPath -OutputDirectory $outputBibliographyAmbiguous | Out-Null
    $bibliographyAmbiguousIssues = @(Import-Csv -LiteralPath (Join-Path $outputBibliographyAmbiguous 'issues.csv') -Encoding utf8)
    Assert-Equal -Expected 1 -Actual @($bibliographyAmbiguousIssues | Where-Object issue_type -eq 'document_bibliography_multiple_matches').Count -Message 'ambiguous bibliography paragraph is reported'
    Assert-Equal -Expected 0 -Actual @($bibliographyAmbiguousIssues | Where-Object issue_type -eq 'registry_missing_from_document_bibliography' | Where-Object source_id -eq 'SRC-A').Count -Message 'ambiguous observation does not become a false absence finding'

    & $auditScript -InputDocx $fixturePath -BibliographyCsv $bibliographyPath -PolicyJson $policyWithBibliographyPath -BibliographyDocx $bibliographyDocxPathNoMarker -OutputDirectory $outputBibliographyMissing | Out-Null
    $bibliographyMissingSummary = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $outputBibliographyMissing 'summary.json') | ConvertFrom-Json
    $bibliographyMissingIssues = @(Import-Csv -LiteralPath (Join-Path $outputBibliographyMissing 'issues.csv') -Encoding utf8)
    Assert-Equal -Expected 'marker_not_found' -Actual $bibliographyMissingSummary.bibliography_reconciliation.status -Message 'missing marker sets marker_not_found status'
    Assert-Equal -Expected 1 -Actual @($bibliographyMissingIssues | Where-Object issue_type -eq 'document_bibliography_not_found').Count -Message 'missing marker emits document_bibliography_not_found'
    Assert-Equal -Expected 0 -Actual @($bibliographyMissingIssues | Where-Object issue_type -eq 'registry_missing_from_document_bibliography').Count -Message 'missing marker does not assert that cited sources are absent'

    & $auditScript -InputDocx $fixturePath -BibliographyCsv $bibliographyPath -PolicyJson $policyPath -OutputDirectory $outputNoPolicyReconciliation | Out-Null
    $noPolicyReconciliation = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $outputNoPolicyReconciliation 'summary.json') | ConvertFrom-Json
    $noPolicyReconciliationRows = @(Import-Csv -LiteralPath (Join-Path $outputNoPolicyReconciliation 'bibliography-reconciliation.csv') -Encoding utf8)
    $noPolicyReconciliationIssues = @(Import-Csv -LiteralPath (Join-Path $outputNoPolicyReconciliation 'issues.csv') -Encoding utf8)
    Assert-Equal -Expected 'not_observed' -Actual $noPolicyReconciliation.bibliography_reconciliation.status -Message 'policy without bibliography block stays not_observed'
    Assert-Equal -Expected 0 -Actual $noPolicyReconciliationRows.Count -Message 'policy without bibliography recon writes header-only bibliography reconciliation'
    Assert-Equal -Expected 0 -Actual @($noPolicyReconciliationIssues | Where-Object issue_type -eq 'document_bibliography_not_found').Count -Message 'not observed mode emits no document bibliography marker issue'

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

    $policyPathConfigured = Join-Path $testRoot 'policy-configured.json'
    $configuredBibliographyPath = Join-Path $testRoot 'bibliography-configured.csv'
    $configuredOutputOne = Join-Path $testRoot 'audit-configured-one'
    $configuredOutputTwo = Join-Path $testRoot 'audit-configured-two'
    $configuredFixturePath = Join-Path $testRoot 'configured-footnotes.docx'

    & $fixtureGenerator -OutputPath $configuredFixturePath | Out-Null
    $configuredBibliographyRows = @(
        [pscustomobject][ordered]@{
            source_id = 'SRC-A'
            source_type = 'book'
            language = 'en'
            author = 'Aster Vale'
            short_title = 'Clockwork Orchards'
            aliases = 'Aster Vale|Clockwork Orchards'
            bibliography_entry = 'Aster Vale. Clockwork Orchards. Imaginary Press, 2042.'
            title = 'Clockwork Orchards'
            translator = 'Mina Vale'
            publication_place = 'Fable Port'
            publisher = 'Imaginary Press'
            year = '2042'
            journal = ''
            volume = ''
            issue = ''
        },
        [pscustomobject][ordered]@{
            source_id = 'SRC-C'
            source_type = 'translated_book'
            language = 'en'
            author = 'Cedar Vane'
            short_title = 'Ember Journal'
            aliases = 'Cedar Vane|Ember Journal'
            bibliography_entry = 'Cedar Vane. Ember Journal.'
            title = 'Ember Journal'
            translator = ''
            publication_place = ''
            publisher = 'Northlight'
            year = ''
            journal = 'Night Ledger'
            volume = '4'
            issue = '2'
        },
        [pscustomobject][ordered]@{
            source_id = 'SRC-D'
            source_type = 'article'
            language = 'en'
            author = 'Mira Vale'
            short_title = 'Paper Bridges'
            aliases = 'Mira Vale|Paper Bridges'
            bibliography_entry = 'Mira Vale, "Paper Bridges", Imaginary Review 4, 2044.'
            title = 'Paper Bridges'
            translator = ''
            publication_place = 'Tokyo'
            publisher = 'Imaginary Review'
            year = '2044'
            journal = 'Imaginary Review'
            volume = '4'
            issue = '2'
        },
        [pscustomobject][ordered]@{
            source_id = 'SRC-E'
            source_type = 'mystery'
            language = 'en'
            author = 'Rin Vale'
            short_title = 'Ghost Index'
            aliases = 'Ghost Index'
            bibliography_entry = 'Rin Vale. Ghost Index.'
            title = 'Ghost Index'
            translator = ''
            publication_place = ''
            publisher = 'Ghostline'
            year = '2039'
            journal = ''
            volume = ''
            issue = ''
        },
        [pscustomobject][ordered]@{
            source_id = 'SRC-F'
            source_type = 'article'
            language = 'en'
            author = 'Sora Kuro'
            short_title = 'Quiet Tower'
            aliases = 'Quiet Tower|Sora Kuro'
            bibliography_entry = 'Sora Kuro, "Quiet Tower", Eastward Press, 2040.'
            title = 'Quiet Tower'
            translator = ''
            publication_place = 'Tokyo'
            publisher = 'Eastward Press'
            year = '2040'
            journal = 'Novella Log'
            volume = '7'
            issue = '3'
        }
    )
    Set-Content -LiteralPath $configuredBibliographyPath -Value ($configuredBibliographyRows | ConvertTo-Csv -NoTypeInformation) -Encoding utf8

    $configuredPolicy = [pscustomobject][ordered]@{
        policy_name = 'configured-synthetic-policy'
        subsequent_citation = 'short_form_review'
        consecutive_same_source = 'short-form'
        new_sources = 'human_review_required'
        japanese_note_terminal_mark = '.'
        foreign_note_terminal_mark = '.'
        source_type_policies = [ordered]@{
            default = [ordered]@{
                first_use_required_fields = @('author', 'short_title')
                subsequent_use_required_fields = @('year')
                consecutive_same_source = 'ibid'
            }
            book = [ordered]@{
                first_use_required_fields = @('title', 'year')
                subsequent_use_required_fields = @('year')
                consecutive_same_source = 'ibid'
            }
            article = [ordered]@{
                first_use_required_fields = @('short_title', 'year')
                subsequent_use_required_fields = @('journal', 'volume')
                consecutive_same_source = 'short-form'
            }
            translated_book = [ordered]@{
                first_use_required_fields = @('translator', 'title', 'year')
                subsequent_use_required_fields = @('issue', 'journal')
                consecutive_same_source = 'ibid'
            }
        }
    }
    Set-Content -LiteralPath $policyPathConfigured -Value ($configuredPolicy | ConvertTo-Json -Depth 8) -Encoding utf8

    $configuredHashBefore = (Get-FileHash -LiteralPath $configuredFixturePath -Algorithm SHA256).Hash
    & $auditScript -InputDocx $configuredFixturePath -BibliographyCsv $configuredBibliographyPath -PolicyJson $policyPathConfigured -OutputDirectory $configuredOutputOne | Out-Null
    & $auditScript -InputDocx $configuredFixturePath -BibliographyCsv $configuredBibliographyPath -PolicyJson $policyPathConfigured -OutputDirectory $configuredOutputTwo | Out-Null

    $configuredCitations = @(Import-Csv -LiteralPath (Join-Path $configuredOutputOne 'citations.csv') -Encoding utf8)
    $configuredVariants = @(Import-Csv -LiteralPath (Join-Path $configuredOutputOne 'citation-variants.csv') -Encoding utf8)
    $configuredIssues = @(Import-Csv -LiteralPath (Join-Path $configuredOutputOne 'issues.csv') -Encoding utf8)
    $configuredSummary = Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $configuredOutputOne 'summary.json') | ConvertFrom-Json

    Assert-Equal -Expected 1 -Actual @($configuredIssues | Where-Object issue_type -eq 'source_type_policy_missing').Count -Message 'unknown source_type emits policy-missing issue once'
    Assert-Equal -Expected 4 -Actual @($configuredIssues | Where-Object issue_type -eq 'citation_required_component_missing').Count -Message 'configured policy flags required component mismatches'
    Assert-Equal -Expected 2 -Actual @($configuredIssues | Where-Object issue_type -eq 'bibliography_metadata_missing').Count -Message 'configured policy flags blank required metadata'
    Assert-Equal -Expected 0 -Actual @($configuredIssues | Where-Object issue_type -eq 'terminal_mark_mismatch').Count -Message 'configured policy flags no terminal mark mismatch for fixture citations'
    Assert-True -Condition ($configuredSummary.counts.citation_variants -gt 0) -Message 'citation variant report is emitted'
    Assert-Equal -Expected 5 -Actual $configuredVariants.Count -Message 'citation-variants rows are emitted for each explicit citation'
    Assert-Equal -Expected 0 -Actual @($configuredIssues | Where-Object issue_type -eq 'citation_variant' | Where-Object footnote_number -eq '7').Count -Message 'page-only locator differences do not create citation variants'
    Assert-Equal -Expected 'evaluated' -Actual ($configuredVariants | Where-Object source_id -eq 'SRC-A' | Where-Object citation_classification -eq 'repeat')[0].comparison_status -Message 'explicit repeat citations are evaluated under configured policy'
    Assert-Equal -Expected 'evaluated' -Actual (@($configuredCitations | Where-Object citation_classification -eq 'repeat' | Where-Object source_id -eq 'SRC-A')[0].comparison_status) -Message 'repeat citations retain an explicit comparison status'
    Assert-Equal -Expected '2' -Actual (($configuredVariants | Where-Object citation_classification -eq 'repeat' | Sort-Object {[int]$_.reference_footnote_number}, {[int]$_.footnote_number} | Select-Object -First 1).reference_footnote_number) -Message 'reference footnote selection is deterministic for repeat source/class'
    Assert-Equal -Expected 'true' -Actual (([string](@($configuredCitations | Where-Object source_id -eq 'SRC-A' | Where-Object citation_classification -eq 'repeat')[0].ibid_rewrite_candidate)).ToLowerInvariant()) -Message 'configured source policy permits an ibid rewrite candidate for adjacent repeats'
    Assert-Equal -Expected 0 -Actual @($configuredCitations | Where-Object citation_classification -eq 'ibid_candidate').Count -Message 'legacy ibid_candidate classification is not emitted'

    Assert-Equal -Expected $configuredHashBefore -Actual (Get-FileHash -LiteralPath $configuredFixturePath -Algorithm SHA256).Hash -Message 'configured policy run preserves configured fixture input hash'

    foreach ($duplicateCase in @(
        [pscustomobject]@{ Path = $duplicateDocumentPath; Label = 'duplicate document part' },
        [pscustomobject]@{ Path = $duplicateFootnotesPath; Label = 'duplicate footnotes part' }
    )) {
        $duplicateRejected = $false
        try {
            & $auditScript -InputDocx $duplicateCase.Path -BibliographyCsv $bibliographyPath -PolicyJson $policyPath -OutputDirectory (Join-Path $testRoot ("reject-{0}" -f ($duplicateCase.Label -replace ' ', '-'))) | Out-Null
        }
        catch {
            $duplicateRejected = $_.Exception.Message -like '*duplicate required part*'
        }
        Assert-True -Condition $duplicateRejected -Message "$($duplicateCase.Label) is rejected"
    }

    $duplicateBibliographyRejected = $false
    try {
        & $auditScript -InputDocx $fixturePath -BibliographyCsv $bibliographyPath -PolicyJson $policyWithBibliographyPath -BibliographyDocx $duplicateBibliographyPath -OutputDirectory (Join-Path $testRoot 'reject-duplicate-bibliography') | Out-Null
    }
    catch {
        $duplicateBibliographyRejected = $_.Exception.Message -like '*duplicate required part*'
    }
    Assert-True -Condition $duplicateBibliographyRejected -Message 'duplicate required part in bibliography DOCX is rejected'

    $excessiveEntriesRejected = $false
    try {
        & $auditScript -InputDocx $excessiveEntriesPath -BibliographyCsv $bibliographyPath -PolicyJson $policyPath -OutputDirectory (Join-Path $testRoot 'reject-excessive-entries') | Out-Null
    }
    catch {
        $excessiveEntriesRejected = $_.Exception.Message -like '*entry count*'
    }
    Assert-True -Condition $excessiveEntriesRejected -Message 'DOCX with excessive ZIP entry count is rejected before archive expansion'

    $nonDocxRejected = $false
    try {
        & $auditScript -InputDocx (Join-Path $testRoot 'not-a-docx.txt') -BibliographyCsv $bibliographyPath -PolicyJson $policyPath -OutputDirectory (Join-Path $testRoot 'invalid-extension-output') | Out-Null
    }
    catch {
        $nonDocxRejected = $_.Exception.Message -like '*.docx extension*'
    }
    Assert-True -Condition $nonDocxRejected -Message 'non-DOCX input is rejected'

    Write-Output "PASS: $($summary.counts.footnotes) footnotes, $($summary.counts.citations) citations, $($summary.counts.issues) issues; deterministic outputs and unchanged input hash in all runs."
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
