param(
    [string]$SourceHtml = "",
    [string]$RepositoryUrl = "https://github.com/aanya-singhh/nvidia-ai-dictionary.git",
    [string]$Branch = "main",
    [string]$NewsletterEndDate = ""
)

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageRoot = Split-Path -Parent $projectRoot
if ([string]::IsNullOrWhiteSpace($SourceHtml)) {
    $SourceHtml = Join-Path $projectRoot "index.html"
}
$checkoutRoot = Join-Path $packageRoot ".github_publish"
$repositoryPath = Join-Path $checkoutRoot "nvidia-ai-dictionary"

if (-not (Test-Path -LiteralPath $SourceHtml)) {
    throw "Website source file not found: $SourceHtml"
}
if (-not (Test-Path -LiteralPath $checkoutRoot)) {
    New-Item -ItemType Directory -Path $checkoutRoot | Out-Null
}
if (-not (Test-Path -LiteralPath (Join-Path $repositoryPath ".git"))) {
    git clone --branch $Branch --single-branch $RepositoryUrl $repositoryPath
    if ($LASTEXITCODE -ne 0) {
        throw "Could not clone the website repository."
    }
}

$safeRepositoryPath = [System.IO.Path]::GetFullPath($repositoryPath).Replace("\", "/")
$gitArgs = @("-c", "safe.directory=$safeRepositoryPath", "-C", $repositoryPath)

git @gitArgs fetch origin $Branch
if ($LASTEXITCODE -ne 0) {
    throw "Could not fetch the website repository."
}
git @gitArgs checkout $Branch
if ($LASTEXITCODE -ne 0) {
    throw "Could not check out website branch $Branch."
}
git @gitArgs pull --ff-only origin $Branch
if ($LASTEXITCODE -ne 0) {
    throw "Website repository is not fast-forwardable; publishing stopped."
}

Copy-Item -LiteralPath $SourceHtml -Destination (Join-Path $repositoryPath "index.html") -Force
git @gitArgs add -- index.html
git @gitArgs diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
    Write-Host "Website already matches GitHub; no commit needed."
    exit 0
}

$configuredName = git @gitArgs config user.name
if ([string]::IsNullOrWhiteSpace($configuredName)) {
    git @gitArgs config user.name "aanya-singhh"
}
$configuredEmail = git @gitArgs config user.email
if ([string]::IsNullOrWhiteSpace($configuredEmail)) {
    git @gitArgs config user.email "aanya-singhh@users.noreply.github.com"
}

$commitDate = if ([string]::IsNullOrWhiteSpace($NewsletterEndDate)) {
    (Get-Date).ToString("yyyy-MM-dd")
} else {
    $NewsletterEndDate
}
git @gitArgs commit -m "Update weekly AI dictionary for $commitDate"
if ($LASTEXITCODE -ne 0) {
    throw "Could not commit the website update."
}
git @gitArgs push origin $Branch
if ($LASTEXITCODE -ne 0) {
    throw "Could not push the website update to GitHub."
}

Write-Host "Website published to GitHub Pages repository."
