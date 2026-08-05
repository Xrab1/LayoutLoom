[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [switch]$SkipPortableArchive
)

$ErrorActionPreference = "Stop"
$projectRoot = (Resolve-Path -LiteralPath (Split-Path -Parent $MyInvocation.MyCommand.Path)).Path
$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"
$releaseRoot = Join-Path $projectRoot "release"
$stageRoot = Join-Path $releaseRoot ".staging"

function Remove-WorkspaceChild {
    param([string]$Path, [string]$ExpectedParent)
    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }
    $parent = [System.IO.Path]::GetFullPath($ExpectedParent).TrimEnd('\') + '\'
    $target = [System.IO.Path]::GetFullPath($Path)
    if (-not $target.StartsWith($parent, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove path outside the release directory: $target"
    }
    Remove-Item -LiteralPath $target -Recurse -Force
}

if (-not (Test-Path -LiteralPath $pythonExe -PathType Leaf)) {
    throw "Project virtual environment is missing. Run install.ps1 first."
}
if (-not $SkipBuild) {
    & (Join-Path $projectRoot "build.ps1")
    if ($LASTEXITCODE -ne 0) {
        throw "LayoutLoom build failed."
    }
}

$bundleDirectory = Join-Path $projectRoot "dist\LayoutLoom"
$exePath = Join-Path $bundleDirectory "LayoutLoom.exe"
if (-not (Test-Path -LiteralPath $exePath -PathType Leaf)) {
    throw "Portable bundle is missing: $exePath"
}
$version = (& $pythonExe -c "import pathlib, tomllib; print(tomllib.loads(pathlib.Path('pyproject.toml').read_text(encoding='utf-8'))['project']['version'])").Trim()
if ($LASTEXITCODE -ne 0 -or -not $version) {
    throw "Cannot read the project version from pyproject.toml."
}

New-Item -ItemType Directory -Path $releaseRoot -Force | Out-Null
Remove-WorkspaceChild -Path $stageRoot -ExpectedParent $releaseRoot
New-Item -ItemType Directory -Path $stageRoot | Out-Null

$sourceStage = Join-Path $stageRoot "LayoutLoom-Source-$version"
New-Item -ItemType Directory -Path $sourceStage | Out-Null
$sourceDirectories = @("docuforge", "tests")
foreach ($name in $sourceDirectories) {
    Copy-Item -LiteralPath (Join-Path $projectRoot $name) -Destination $sourceStage -Recurse -Force
}
$sourceFiles = @(
    ".gitignore",
    "LICENSE",
    "README.md",
    "SOURCE_CODE.md",
    "THIRD_PARTY_NOTICES.md",
    "pyproject.toml",
    "requirements.txt",
    "requirements-dev.txt",
    "launcher.py",
    "install.ps1",
    "run.ps1",
    "build.ps1",
    "release.ps1",
    "prepare_portable_dependencies.ps1",
    "audit_all_features.py"
)
foreach ($name in $sourceFiles) {
    $source = Join-Path $projectRoot $name
    if (Test-Path -LiteralPath $source -PathType Leaf) {
        Copy-Item -LiteralPath $source -Destination (Join-Path $sourceStage $name) -Force
    }
}
$batchLaunchers = @(Get-ChildItem -LiteralPath $projectRoot -File -Filter "*.bat")
if ($batchLaunchers.Count -ne 3) {
    throw "Expected exactly 3 launcher BAT files, found $($batchLaunchers.Count)."
}
foreach ($batchLauncher in $batchLaunchers) {
    Copy-Item -LiteralPath $batchLauncher.FullName -Destination $sourceStage -Force
}
Get-ChildItem -LiteralPath $sourceStage -Recurse -Directory |
    Where-Object { $_.Name -in @("__pycache__", ".pytest_cache", ".mypy_cache") } |
    Sort-Object { $_.FullName.Length } -Descending |
    ForEach-Object { Remove-WorkspaceChild -Path $_.FullName -ExpectedParent $sourceStage }
Get-ChildItem -LiteralPath $sourceStage -Recurse -File |
    Where-Object { $_.Extension -in @(".pyc", ".pyo") } |
    ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force }
$sourceFileCount = @(Get-ChildItem -LiteralPath $sourceStage -Recurse -File).Count
if ($sourceFileCount -lt 50) {
    throw "Source staging is incomplete: only $sourceFileCount files were collected."
}
foreach ($requiredSource in @("docuforge\app.py", "docuforge\registry.py", "tests\test_core.py", "LICENSE", "pyproject.toml")) {
    if (-not (Test-Path -LiteralPath (Join-Path $sourceStage $requiredSource) -PathType Leaf)) {
        throw "Source staging is missing required file: $requiredSource"
    }
}

$versionFile = Join-Path $bundleDirectory "VERSION.txt"
Set-Content -LiteralPath $versionFile -Value $version -Encoding UTF8
$portableZip = Join-Path $releaseRoot "LayoutLoom-Windows-x64-Portable-$version.zip"
$sourceZip = Join-Path $releaseRoot "LayoutLoom-Source-$version.zip"
$archivesToReplace = @($sourceZip)
if (-not $SkipPortableArchive) {
    $archivesToReplace += $portableZip
}
foreach ($archive in $archivesToReplace) {
    if (Test-Path -LiteralPath $archive) {
        $resolvedRelease = [System.IO.Path]::GetFullPath($releaseRoot).TrimEnd('\') + '\'
        $resolvedArchive = [System.IO.Path]::GetFullPath($archive)
        if (-not $resolvedArchive.StartsWith($resolvedRelease, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Unsafe archive target: $resolvedArchive"
        }
        Remove-Item -LiteralPath $resolvedArchive -Force
    }
}
if (-not $SkipPortableArchive) {
    Compress-Archive -LiteralPath $bundleDirectory -DestinationPath $portableZip -CompressionLevel Optimal
} elseif (-not (Test-Path -LiteralPath $portableZip -PathType Leaf)) {
    throw "SkipPortableArchive was requested, but the portable archive does not exist: $portableZip"
}
Compress-Archive -LiteralPath $sourceStage -DestinationPath $sourceZip -CompressionLevel Optimal

$hashLines = foreach ($archive in @($portableZip, $sourceZip)) {
    $hash = Get-FileHash -LiteralPath $archive -Algorithm SHA256
    "{0}  {1}" -f $hash.Hash.ToLowerInvariant(), (Split-Path -Leaf $archive)
}
Set-Content -LiteralPath (Join-Path $releaseRoot "SHA256SUMS.txt") -Value $hashLines -Encoding ASCII

Remove-WorkspaceChild -Path $stageRoot -ExpectedParent $releaseRoot
Write-Host "Open-source release artifacts are ready:"
Write-Host "  $portableZip"
Write-Host "  $sourceZip"
Write-Host "  $(Join-Path $releaseRoot 'SHA256SUMS.txt')"
