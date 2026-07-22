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
$installer = Join-Path $repoRoot 'install.ps1'
$tempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("thesis-footnote-installer-tests-" + [guid]::NewGuid().ToString('N'))
$previousClaudeConfig = $env:CLAUDE_CONFIG_DIR

try {
    New-Item -ItemType Directory -Path $testRoot | Out-Null

    # A later collision must be discovered before any earlier adapter is installed.
    $preflightCodex = Join-Path $testRoot 'preflight-codex'
    $preflightClaude = Join-Path $testRoot 'preflight-claude'
    $existingClaudeAgent = Join-Path $preflightClaude 'agents\footnote-normalization-reviewer.md'
    New-Item -ItemType Directory -Path (Split-Path -Parent $existingClaudeAgent) -Force | Out-Null
    Set-Content -LiteralPath $existingClaudeAgent -Value 'existing' -Encoding utf8
    $collisionRejected = $false
    try {
        & $installer -Runtime both -CodexConfigDir $preflightCodex -ClaudeConfigDir $preflightClaude -AllowExternalHome | Out-Null
    }
    catch {
        $collisionRejected = $_.Exception.Message -like '*すでに存在*'
    }
    Assert-True -Condition $collisionRejected -Message 'later destination collision is rejected'
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $preflightCodex 'skills\normalize-thesis-footnotes'))) -Message 'preflight failure leaves earlier destinations untouched'

    # Claude Code must honor its official configuration directory variable.
    $claudeConfig = Join-Path $testRoot 'claude-config-env'
    $env:CLAUDE_CONFIG_DIR = $claudeConfig
    & $installer -Runtime claude -AllowExternalHome | Out-Null
    $claudeSkill = Join-Path $claudeConfig 'skills\normalize-thesis-footnotes\SKILL.md'
    $claudeAgent = Join-Path $claudeConfig 'agents\footnote-normalization-reviewer.md'
    Assert-True -Condition (Test-Path -LiteralPath $claudeSkill -PathType Leaf) -Message 'CLAUDE_CONFIG_DIR receives the skill'
    Assert-True -Condition (Test-Path -LiteralPath $claudeAgent -PathType Leaf) -Message 'CLAUDE_CONFIG_DIR receives the agent'

    # Force is an exact replacement, not a merge that preserves stale files.
    $staleFile = Join-Path $claudeConfig 'skills\normalize-thesis-footnotes\stale-from-old-version.txt'
    Set-Content -LiteralPath $staleFile -Value 'stale' -Encoding utf8
    Set-Content -LiteralPath $claudeAgent -Value 'stale agent content' -Encoding utf8
    & $installer -Runtime claude -AllowExternalHome -Force | Out-Null
    Assert-True -Condition (-not (Test-Path -LiteralPath $staleFile)) -Message 'Force removes files that exist only in the old adapter'
    $expectedAgent = Get-Content -LiteralPath (Join-Path $repoRoot 'claude\agents\footnote-normalization-reviewer.md') -Raw -Encoding utf8
    $installedAgent = Get-Content -LiteralPath $claudeAgent -Raw -Encoding utf8
    Assert-True -Condition ($installedAgent -eq $expectedAgent) -Message 'Force replaces the agent exactly'

    # WhatIf must not create destinations and must describe a dry run.
    $whatIfCodex = Join-Path $testRoot 'whatif-codex'
    $whatIfText = (& $installer -Runtime codex -CodexConfigDir $whatIfCodex -AllowExternalHome -WhatIf 2>&1 | Out-String)
    Assert-True -Condition ($whatIfText -like '*アダプターはインストールしていません*') -Message 'WhatIf states that no installation occurred'
    Assert-True -Condition (-not (Test-Path -LiteralPath $whatIfCodex)) -Message 'WhatIf creates no configuration directory'

    Write-Output 'PASS: installer preflight, CLAUDE_CONFIG_DIR, exact Force replacement, and WhatIf behavior.'
}
finally {
    $env:CLAUDE_CONFIG_DIR = $previousClaudeConfig
    $validatedRoot = [System.IO.Path]::GetFullPath($testRoot)
    if (-not $validatedRoot.StartsWith($tempRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
        -not ([System.IO.Path]::GetFileName($validatedRoot)).StartsWith('thesis-footnote-installer-tests-', [System.StringComparison]::Ordinal)) {
        throw "Refusing to clean an unvalidated installer test directory: $validatedRoot"
    }
    if (Test-Path -LiteralPath $validatedRoot) {
        Remove-Item -LiteralPath $validatedRoot -Recurse -Force
    }
}
