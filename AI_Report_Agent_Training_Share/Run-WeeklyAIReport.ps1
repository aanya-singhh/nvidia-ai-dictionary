param(
    [string]$ConfigPath = "",
    [string]$InputPath = "",
    [switch]$SkipEmailDraft,
    [switch]$SkipWebsitePublish
)

$packageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptsRoot = Join-Path $packageRoot "scripts"
$dataRoot = Join-Path $packageRoot "data"
$siteRoot = Join-Path $packageRoot "site"
$outputRoot = Join-Path $packageRoot "output"

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $packageRoot "local_config.json"
}
if (-not [System.IO.Path]::IsPathRooted($ConfigPath)) {
    $ConfigPath = Join-Path $packageRoot $ConfigPath
}
if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Create local_config.json from local_config.example.json before running."
}

$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]::IsNullOrWhiteSpace([string]$config.python_path)) {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if (-not $pythonCommand) {
        throw "Python was not found. Install Python 3.11+ and select Add Python to PATH."
    }
    $python = $pythonCommand.Source
} else {
    $python = [string]$config.python_path
}

if ([string]::IsNullOrWhiteSpace($InputPath)) {
    $latestInput = Get-ChildItem -LiteralPath $dataRoot -Filter "weekly_ai_report_input_*.json" |
        Sort-Object Name |
        Select-Object -Last 1
    if (-not $latestInput) {
        throw "No weekly_ai_report_input_YYYY-MM-DD.json file was found in data."
    }
    $InputPath = $latestInput.FullName
} elseif (-not [System.IO.Path]::IsPathRooted($InputPath)) {
    $InputPath = Join-Path $packageRoot $InputPath
}
$InputPath = [System.IO.Path]::GetFullPath($InputPath)
if (-not (Test-Path -LiteralPath $InputPath)) {
    throw "Weekly input not found: $InputPath"
}

$inputData = Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8 | ConvertFrom-Json
$weekEndIso = [string]$inputData.date_range.end
if ([string]::IsNullOrWhiteSpace($weekEndIso)) {
    throw "Weekly input is missing date_range.end."
}

$modelsCsv = Join-Path $dataRoot "ai-models.csv"
$tracker = Join-Path $dataRoot "emerging_ai_tracker.json"
$bestForMemory = Join-Path $dataRoot "best_for_memory.json"
$indexHtml = Join-Path $siteRoot "index.html"

if (-not [bool]$config.approve_public_model_search) {
    throw (
        "Public model search is not approved. Review README-FIRST.md, then set " +
        "approve_public_model_search to true in local_config.json if authorized. " +
        "The search sends public model names and makers from ai-models.csv to Bing."
    )
}

New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null

& $python (Join-Path $scriptsRoot "build_weekly_model_research_plan.py") `
    --models-csv $modelsCsv `
    --week-end $weekEndIso `
    --output-dir $dataRoot
if ($LASTEXITCODE -ne 0) {
    throw "Research-plan generation failed."
}

& $python (Join-Path $scriptsRoot "run_weekly_watchlist_inquiry.py") `
    --models-csv $modelsCsv `
    --week-end $weekEndIso `
    --output-dir $dataRoot
if ($LASTEXITCODE -ne 0) {
    throw "All-target public discovery pass failed."
}

$inputData = Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8 | ConvertFrom-Json
$modelCount = (Import-Csv -LiteralPath $modelsCsv | Measure-Object).Count
$planName = "weekly_ai_model_research_plan_$weekEndIso.md"
$ledgerName = "weekly_ai_supplemental_research_ledger_$weekEndIso.json"
if (-not ($inputData.model_watchlist.PSObject.Properties.Name -contains "research_ledger")) {
    $inputData.model_watchlist | Add-Member -NotePropertyName "research_ledger" -NotePropertyValue $ledgerName
}
$inputData.model_watchlist.source_csv = "ai-models.csv"
$inputData.model_watchlist.research_plan = $planName
$inputData.model_watchlist.research_ledger = $ledgerName
$inputData.model_watchlist.model_count = $modelCount
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText(
    $InputPath,
    (($inputData | ConvertTo-Json -Depth 100) + [Environment]::NewLine),
    $utf8NoBom
)

$generatorArgs = @(
    (Join-Path $scriptsRoot "weekly_ai_report_generator.py"),
    "--input", $InputPath,
    "--tracker", $tracker,
    "--best-for-memory", $bestForMemory,
    "--output-dir", $outputRoot,
    "--draft-style", "ai_learning_report",
    "--subscribe-url", [string]$config.subscribe_url
)
foreach ($recipient in @($config.email_recipients)) {
    if (-not [string]::IsNullOrWhiteSpace([string]$recipient)) {
        $generatorArgs += @("--email-recipient", [string]$recipient)
    }
}
if ($SkipEmailDraft) {
    $generatorArgs += "--skip-email-draft"
}

& $python @generatorArgs
if ($LASTEXITCODE -ne 0) {
    throw "Report generation failed; website publishing was skipped."
}

& $python (Join-Path $scriptsRoot "update_ai_dictionary_site.py") `
    --input $InputPath `
    --html $indexHtml `
    --models-csv $modelsCsv
if ($LASTEXITCODE -ne 0) {
    throw "Website generation failed; GitHub publishing was skipped."
}

& $python (Join-Path $scriptsRoot "reconcile_report_discoveries.py") `
    --input $InputPath `
    --models-csv $modelsCsv
if ($LASTEXITCODE -ne 0) {
    throw "New report discoveries could not be reconciled safely."
}

if ([bool]$config.publish_website -and -not $SkipWebsitePublish) {
    $repositoryUrl = [string]$config.github_repository_url
    if (
        [string]::IsNullOrWhiteSpace($repositoryUrl) -or
        $repositoryUrl.Contains("YOUR-GITHUB-USER")
    ) {
        throw "Set github_repository_url in local_config.json before publishing."
    }
    $branch = [string]$config.github_branch
    if ([string]::IsNullOrWhiteSpace($branch)) {
        $branch = "main"
    }
    $windowsPowerShell = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    & $windowsPowerShell -NoProfile -ExecutionPolicy Bypass `
        -File (Join-Path $scriptsRoot "publish_ai_dictionary_site.ps1") `
        -SourceHtml $indexHtml `
        -RepositoryUrl $repositoryUrl `
        -Branch $branch `
        -NewsletterEndDate $weekEndIso
    if ($LASTEXITCODE -ne 0) {
        throw "GitHub publication failed."
    }
} else {
    Write-Host "Website publication skipped by configuration or command switch."
}

Write-Host "Weekly AI workflow completed."
