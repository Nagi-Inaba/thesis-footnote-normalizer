[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('codex', 'claude', 'both')]
    [string]$Runtime = 'both',

    [switch]$Force,

    [switch]$AllowExternalHome,

    [string]$CodexConfigDir,

    [string]$ClaudeConfigDir
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$repoRoot = Split-Path -Parent $PSCommandPath

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

function Assert-NoReparsePointInTree {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }
    $items = @((Get-Item -LiteralPath $Path -Force))
    if ((Get-Item -LiteralPath $Path -Force).PSIsContainer) {
        $items += @(Get-ChildItem -LiteralPath $Path -Recurse -Force)
    }
    foreach ($item in $items) {
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "$Label contains a reparse point: $($item.FullName)"
        }
    }
}

function Resolve-AdapterHome {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Label is blank."
    }
    if ([string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
        throw 'USERPROFILE is unavailable; use a normal user environment before installation.'
    }

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $userProfile = [System.IO.Path]::GetFullPath($env:USERPROFILE).TrimEnd('\', '/')
    $userPrefix = $userProfile + [System.IO.Path]::DirectorySeparatorChar
    $insideProfile = [string]::Equals($fullPath, $userProfile, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($userPrefix, [System.StringComparison]::OrdinalIgnoreCase)
    if (-not $insideProfile -and -not $AllowExternalHome) {
        throw "$Label is outside USERPROFILE. Rerun with -AllowExternalHome only after reviewing the path: $fullPath"
    }

    Assert-NoReparsePointInPath -Path $fullPath -Label $Label
    return $fullPath
}

function New-AdapterPlanEntry {
    param(
        [Parameter(Mandatory)] [string]$Source,
        [Parameter(Mandatory)] [string]$Destination,
        [Parameter(Mandatory)] [string]$Label
    )

    [pscustomobject][ordered]@{
        Source = [System.IO.Path]::GetFullPath($Source)
        Destination = [System.IO.Path]::GetFullPath($Destination)
        Label = $Label
        Stage = $null
        Backup = $null
        HadExisting = $false
        BackupMoved = $false
        Activated = $false
    }
}

$plan = @()
if ($Runtime -in @('codex', 'both')) {
    $codexRootCandidate = if (-not [string]::IsNullOrWhiteSpace($CodexConfigDir)) {
        $CodexConfigDir
    }
    elseif ($env:CODEX_HOME) {
        $env:CODEX_HOME
    }
    else {
        Join-Path $env:USERPROFILE '.codex'
    }
    $codexRoot = Resolve-AdapterHome -Path $codexRootCandidate -Label 'Codex config directory'
    $plan += New-AdapterPlanEntry `
        -Source (Join-Path $repoRoot 'codex\skills\normalize-thesis-footnotes') `
        -Destination (Join-Path $codexRoot 'skills\normalize-thesis-footnotes') `
        -Label 'Codex skill'
    $plan += New-AdapterPlanEntry `
        -Source (Join-Path $repoRoot 'codex\agents\footnote-normalization-reviewer.toml') `
        -Destination (Join-Path $codexRoot 'agents\footnote-normalization-reviewer.toml') `
        -Label 'Codex agent'
}

if ($Runtime -in @('claude', 'both')) {
    $claudeRootCandidate = if (-not [string]::IsNullOrWhiteSpace($ClaudeConfigDir)) {
        $ClaudeConfigDir
    }
    elseif ($env:CLAUDE_CONFIG_DIR) {
        $env:CLAUDE_CONFIG_DIR
    }
    else {
        Join-Path $env:USERPROFILE '.claude'
    }
    $claudeRoot = Resolve-AdapterHome -Path $claudeRootCandidate -Label 'Claude config directory'
    $plan += New-AdapterPlanEntry `
        -Source (Join-Path $repoRoot 'claude\skills\normalize-thesis-footnotes') `
        -Destination (Join-Path $claudeRoot 'skills\normalize-thesis-footnotes') `
        -Label 'Claude skill'
    $plan += New-AdapterPlanEntry `
        -Source (Join-Path $repoRoot 'claude\agents\footnote-normalization-reviewer.md') `
        -Destination (Join-Path $claudeRoot 'agents\footnote-normalization-reviewer.md') `
        -Label 'Claude agent'
}

# Validate the complete plan before creating or replacing anything.
foreach ($entry in $plan) {
    if (-not (Test-Path -LiteralPath $entry.Source)) {
        throw "Missing package source: $($entry.Source)"
    }
    Assert-NoReparsePointInPath -Path $entry.Destination -Label $entry.Label
    $entry.HadExisting = Test-Path -LiteralPath $entry.Destination
    if ($entry.HadExisting) {
        if (-not $Force) {
            throw "$($entry.Label) already exists at $($entry.Destination). Review it and rerun with -Force only if complete replacement is intended."
        }
        Assert-NoReparsePointInTree -Path $entry.Destination -Label $entry.Label
    }
}

if ($WhatIfPreference) {
    foreach ($entry in $plan) {
        $null = $PSCmdlet.ShouldProcess($entry.Destination, "Install $($entry.Label) by exact replacement")
    }
    Write-Output "Dry run completed; no adapters were installed. Planned selection: $Runtime"
    exit 0
}

try {
    # Prepare every payload first so a copy failure cannot leave a partial installation.
    foreach ($entry in $plan) {
        $parent = Split-Path -Parent $entry.Destination
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
        Assert-NoReparsePointInPath -Path $parent -Label "$($entry.Label) parent"
        $leaf = [System.IO.Path]::GetFileName($entry.Destination)
        $entry.Stage = Join-Path $parent ".$leaf.stage.$([guid]::NewGuid().ToString('N'))"
        $sourceItem = Get-Item -LiteralPath $entry.Source -Force
        if ($sourceItem.PSIsContainer) {
            New-Item -ItemType Directory -Path $entry.Stage | Out-Null
            Get-ChildItem -LiteralPath $entry.Source -Force |
                Copy-Item -Destination $entry.Stage -Recurse -Force
        }
        else {
            Copy-Item -LiteralPath $entry.Source -Destination $entry.Stage
        }
    }

    foreach ($entry in $plan) {
        if ($entry.HadExisting) {
            $parent = Split-Path -Parent $entry.Destination
            $leaf = [System.IO.Path]::GetFileName($entry.Destination)
            $entry.Backup = Join-Path $parent ".$leaf.backup.$([guid]::NewGuid().ToString('N'))"
            Move-Item -LiteralPath $entry.Destination -Destination $entry.Backup
            $entry.BackupMoved = $true
        }
        Move-Item -LiteralPath $entry.Stage -Destination $entry.Destination
        $entry.Activated = $true
        $entry.Stage = $null
    }
}
catch {
    $installError = $_
    foreach ($entry in $plan) {
        if ($entry.Activated -and (Test-Path -LiteralPath $entry.Destination)) {
            Assert-NoReparsePointInTree -Path $entry.Destination -Label "$($entry.Label) rollback target"
            Remove-Item -LiteralPath $entry.Destination -Recurse -Force
        }
        if ($entry.BackupMoved -and (Test-Path -LiteralPath $entry.Backup)) {
            Move-Item -LiteralPath $entry.Backup -Destination $entry.Destination
        }
        if ($entry.Stage -and (Test-Path -LiteralPath $entry.Stage)) {
            Assert-NoReparsePointInTree -Path $entry.Stage -Label "$($entry.Label) staged payload"
            Remove-Item -LiteralPath $entry.Stage -Recurse -Force
        }
    }
    throw $installError
}

foreach ($entry in $plan) {
    if ($entry.BackupMoved -and (Test-Path -LiteralPath $entry.Backup)) {
        Assert-NoReparsePointInTree -Path $entry.Backup -Label "$($entry.Label) replaced backup"
        Remove-Item -LiteralPath $entry.Backup -Recurse -Force
    }
}

Write-Output "Installed adapter selection: $Runtime"
Write-Output 'Keep the cloned repository available because the adapters invoke its PowerShell audit script.'
