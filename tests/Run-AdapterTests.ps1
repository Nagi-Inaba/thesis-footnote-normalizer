[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw "Assertion failed: $Message"
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$codexSkill = Join-Path $repoRoot 'codex\skills\normalize-thesis-footnotes\SKILL.md'
$claudeSkill = Join-Path $repoRoot 'claude\skills\normalize-thesis-footnotes\SKILL.md'
$openAiMetadata = Join-Path $repoRoot 'codex\skills\normalize-thesis-footnotes\agents\openai.yaml'
$codexAgent = Join-Path $repoRoot 'codex\agents\footnote-normalization-reviewer.toml'
$claudeAgent = Join-Path $repoRoot 'claude\agents\footnote-normalization-reviewer.md'

foreach ($path in @($codexSkill, $claudeSkill, $openAiMetadata, $codexAgent, $claudeAgent)) {
    Assert-True -Condition (Test-Path -LiteralPath $path -PathType Leaf) -Message "adapter file exists: $path"
    $text = Get-Content -LiteralPath $path -Raw -Encoding utf8
    Assert-True -Condition (-not $text.Contains([char]0xFFFD)) -Message "adapter is valid UTF-8 without replacement characters: $path"
}

foreach ($skill in @($codexSkill, $claudeSkill)) {
    $text = Get-Content -LiteralPath $skill -Raw -Encoding utf8
    Assert-True -Condition ($text -match '(?ms)\A---\s*\r?\nname:\s*normalize-thesis-footnotes\s*\r?\ndescription:\s*.+?\r?\n---') -Message "skill frontmatter includes name and description: $skill"
    Assert-True -Condition ($text -like '*templates/citation-ledger.csv*') -Message "skill references the proposed-change ledger template: $skill"
}

$metadataText = Get-Content -LiteralPath $openAiMetadata -Raw -Encoding utf8
$shortDescriptionMatch = [regex]::Match($metadataText, '(?m)^\s*short_description:\s*"([^"]+)"\s*$')
Assert-True -Condition $shortDescriptionMatch.Success -Message 'OpenAI metadata has short_description'
$shortDescriptionLength = $shortDescriptionMatch.Groups[1].Value.Length
Assert-True -Condition ($shortDescriptionLength -ge 25 -and $shortDescriptionLength -le 64) -Message 'OpenAI short_description length is 25-64 characters'

$codexAgentText = Get-Content -LiteralPath $codexAgent -Raw -Encoding utf8
Assert-True -Condition ($codexAgentText -match '(?m)^name\s*=\s*"footnote-normalization-reviewer"\s*$') -Message 'Codex agent TOML has the expected name'
$claudeAgentText = Get-Content -LiteralPath $claudeAgent -Raw -Encoding utf8
Assert-True -Condition ($claudeAgentText -match '(?ms)\A---\s*\r?\nname:\s*footnote-normalization-reviewer\s*\r?\ndescription:') -Message 'Claude agent has YAML frontmatter'

$ledgerHeader = Get-Content -LiteralPath (Join-Path $repoRoot 'templates\citation-ledger.csv') -First 1 -Encoding utf8
foreach ($column in @('proposed_text', 'governing_rule', 'confidence', 'approval_status')) {
    Assert-True -Condition ($ledgerHeader.Split(',') -contains $column) -Message "ledger contains $column"
}

$installerText = Get-Content -LiteralPath (Join-Path $repoRoot 'install.ps1') -Raw -Encoding utf8
Assert-True -Condition ($installerText -like '*CLAUDE_CONFIG_DIR*') -Message 'installer supports CLAUDE_CONFIG_DIR'
$workflowText = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\validate.yml') -Raw -Encoding utf8
$readmeText = Get-Content -LiteralPath (Join-Path $repoRoot 'README.md') -Raw -Encoding utf8
Assert-True -Condition ($workflowText -like '*runs-on: windows-2025*') -Message 'workflow uses the fixed Windows 2025 runner'
Assert-True -Condition ($readmeText.Contains('windows-2025')) -Message 'README names the fixed Windows 2025 runner'
Assert-True -Condition ($readmeText.Contains('脚注8件')) -Message 'README footnote fixture count matches the audit test'

$publicPaths = @(& git -C $repoRoot ls-files --cached --others --exclude-standard 2>$null)
if ($LASTEXITCODE -eq 0 -and $publicPaths.Count -gt 0) {
    $publicFiles = @($publicPaths | Sort-Object -Unique | ForEach-Object {
        Get-Item -LiteralPath (Join-Path $repoRoot $_)
    })
}
else {
    $excludedPrefixes = @('.git', '.code-review-graph') | ForEach-Object {
        (Join-Path $repoRoot $_).TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
    }
    $publicFiles = @(Get-ChildItem -LiteralPath $repoRoot -Recurse -File | Where-Object {
        $candidate = $_.FullName
        -not @($excludedPrefixes | Where-Object {
            $candidate.StartsWith($_, [System.StringComparison]::OrdinalIgnoreCase)
        }).Count
    })
}
$forbiddenPatterns = @(
    ('C:' + '\Users\' + 'Na' + 'gi'),
    ('One' + 'Drive'),
    ('BEGIN ' + 'PRIVATE KEY'),
    ('gh' + 'p_'),
    ('gh' + 'o_'),
    ('s' + 'k-')
)
foreach ($pattern in $forbiddenPatterns) {
    $hit = @($publicFiles | Select-String -SimpleMatch -Pattern $pattern -Encoding utf8 -ErrorAction SilentlyContinue)
    Assert-True -Condition ($hit.Count -eq 0) -Message "public tree excludes pattern: $pattern"
}

Write-Output 'PASS: adapter metadata, ledger contract, runtime configuration, and public-tree checks.'
