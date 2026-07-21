[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$testRunner = Join-Path $repositoryRoot 'tests\Run-Tests.ps1'
$installTestRunner = Join-Path $repositoryRoot 'tests\Run-InstallTests.ps1'
$adapterTestRunner = Join-Path $repositoryRoot 'tests\Run-AdapterTests.ps1'

foreach ($requiredTest in @($testRunner, $installTestRunner, $adapterTestRunner)) {
    if (-not (Test-Path -LiteralPath $requiredTest -PathType Leaf)) {
        throw "Test runner not found: $requiredTest"
    }
}

try {
    & $testRunner
    & $installTestRunner
    & $adapterTestRunner
    exit 0
}
catch {
    Write-Error "Repository tests failed: $($_.Exception.Message)"
    exit 1
}
