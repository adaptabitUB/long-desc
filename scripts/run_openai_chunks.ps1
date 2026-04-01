param(
    [string]$ProjectRoot = (Resolve-Path "$PSScriptRoot\..").Path,
    [string]$ChartsJson = "output/charts.json",
    [string]$GeneratorScript = "src/long_descriptions/generate_alt_text_openai.py",
    [int]$ChunkSize = 50,
    [int]$StartChunk = 1,
    [int]$EndChunk = 0,
    [double]$SleepSeconds = 0.0,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Set-Location $ProjectRoot

if (-not (Test-Path $ChartsJson)) {
    throw "Charts JSON not found: $ChartsJson"
}

if (-not (Test-Path $GeneratorScript)) {
    throw "Generator script not found: $GeneratorScript"
}

$charts = Get-Content $ChartsJson -Raw | ConvertFrom-Json
$ids = @(
    $charts |
    ForEach-Object {
        $id = [string]$_.id
        if ($id -match '^CASE_(\d+)$') {
            [int]$matches[1]
        }
    } |
    Sort-Object -Unique
)

if ($ids.Count -eq 0) {
    throw "No CASE_XXXXX ids found in $ChartsJson"
}

if ($ChunkSize -le 0) {
    throw "ChunkSize must be > 0"
}

$totalChunks = [int][Math]::Ceiling($ids.Count / [double]$ChunkSize)
if ($EndChunk -le 0) {
    $EndChunk = $totalChunks
}

if ($StartChunk -lt 1 -or $StartChunk -gt $totalChunks) {
    throw "StartChunk out of range. Must be between 1 and $totalChunks"
}

if ($EndChunk -lt $StartChunk -or $EndChunk -gt $totalChunks) {
    throw "EndChunk out of range. Must be between $StartChunk and $totalChunks"
}

Write-Host "Project root: $ProjectRoot"
Write-Host "Charts found: $($ids.Count)"
Write-Host "Chunk size: $ChunkSize"
Write-Host "Executing chunks: $StartChunk..$EndChunk (of $totalChunks)"

for ($chunk = $StartChunk; $chunk -le $EndChunk; $chunk++) {
    $startIndex = ($chunk - 1) * $ChunkSize
    $endIndex = [Math]::Min($startIndex + $ChunkSize - 1, $ids.Count - 1)

    $startCase = $ids[$startIndex]
    $endCase = $ids[$endIndex]

    $args = @(
        "run",
        "python",
        $GeneratorScript,
        "--start-case", "$startCase",
        "--end-case", "$endCase",
        "--batch-size", "$ChunkSize"
    )

    if ($SleepSeconds -gt 0) {
        $args += @("--sleep-seconds", "$SleepSeconds")
    }

    Write-Host ""
    Write-Host "[Chunk $chunk/$totalChunks] cases $startCase..$endCase"
    Write-Host "Command: uv $($args -join ' ')"

    if ($DryRun) {
        continue
    }

    & uv @args
    if ($LASTEXITCODE -ne 0) {
        throw "Chunk $chunk failed (cases $startCase..$endCase)."
    }
}

Write-Host ""
Write-Host "Done."
