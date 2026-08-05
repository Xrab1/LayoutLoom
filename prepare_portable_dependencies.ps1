[CmdletBinding()]
param(
    [string]$FfmpegSource = "",
    [string]$PopplerSource = ""
)

$ErrorActionPreference = "Stop"
$projectRoot = (Resolve-Path -LiteralPath (Split-Path -Parent $MyInvocation.MyCommand.Path)).Path
$thirdPartyRoot = Join-Path $projectRoot "third_party"
$downloads = Join-Path $thirdPartyRoot "downloads"
$workRoot = Join-Path $thirdPartyRoot ".prepare"
New-Item -ItemType Directory -Path $thirdPartyRoot, $downloads, $workRoot -Force | Out-Null

$popplerVersion = "26.02.0-0"
$popplerSha256 = "993E4A94376ED712FAFC7058D724EA0B943D118BBD2305CD9ED55174EB85CDA5"
$popplerUrl = "https://github.com/oschwartz10612/poppler-windows/releases/download/v$popplerVersion/Release-$popplerVersion.zip"
$ffmpegUrl = "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip"

function Get-VerifiedDownload {
    param(
        [string]$Uri,
        [string]$Destination,
        [string]$ExpectedSha256 = ""
    )
    if (-not (Test-Path -LiteralPath $Destination -PathType Leaf)) {
        & curl.exe -L --fail --retry 5 --retry-all-errors --retry-delay 2 --output $Destination $Uri
        if ($LASTEXITCODE -ne 0) {
            throw "Download failed: $Uri"
        }
    }
    if ($ExpectedSha256) {
        $actual = (Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash
        if ($actual -ne $ExpectedSha256) {
            throw "SHA256 mismatch for '$Destination'. Expected $ExpectedSha256, got $actual."
        }
    }
    return (Resolve-Path -LiteralPath $Destination).Path
}

function Resolve-ArchiveOrDirectory {
    param(
        [string]$Source,
        [string]$DefaultUrl,
        [string]$DownloadName,
        [string]$ExtractName,
        [string]$ExpectedSha256 = ""
    )
    $selected = $Source.Trim()
    if (-not $selected) {
        $selected = Get-VerifiedDownload `
            -Uri $DefaultUrl `
            -Destination (Join-Path $downloads $DownloadName) `
            -ExpectedSha256 $ExpectedSha256
    }
    $resolved = (Resolve-Path -LiteralPath $selected).Path
    if (Test-Path -LiteralPath $resolved -PathType Container) {
        return $resolved
    }
    $extractRoot = Join-Path $workRoot $ExtractName
    if (Test-Path -LiteralPath $extractRoot) {
        $resolvedWork = [System.IO.Path]::GetFullPath($workRoot).TrimEnd('\') + '\'
        $resolvedTarget = [System.IO.Path]::GetFullPath($extractRoot)
        if (-not $resolvedTarget.StartsWith($resolvedWork, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Unsafe preparation target: $resolvedTarget"
        }
        Remove-Item -LiteralPath $resolvedTarget -Recurse -Force
    }
    New-Item -ItemType Directory -Path $extractRoot | Out-Null
    Expand-Archive -LiteralPath $resolved -DestinationPath $extractRoot -Force
    return $extractRoot
}

function Copy-EngineRoot {
    param(
        [string]$SourceRoot,
        [string]$ProbeName,
        [string]$Destination
    )
    $probe = Get-ChildItem -LiteralPath $SourceRoot -Recurse -File -Filter $ProbeName |
        Select-Object -First 1
    if (-not $probe) {
        throw "Cannot find $ProbeName under '$SourceRoot'."
    }
    $root = $probe.Directory
    if ($root.Name -ieq "bin" -and $root.Parent.Name -ieq "Library") {
        $root = $root.Parent.Parent
    } elseif ($root.Name -ieq "bin") {
        $root = $root.Parent
    }
    Copy-Item -LiteralPath $root.FullName -Destination $Destination -Recurse -Force
}

$ffmpegTarget = Join-Path $thirdPartyRoot "ffmpeg"
if (-not (Test-Path -LiteralPath (Join-Path $ffmpegTarget "bin\ffmpeg.exe") -PathType Leaf)) {
    $ffmpegRoot = Resolve-ArchiveOrDirectory `
        -Source $FfmpegSource `
        -DefaultUrl $ffmpegUrl `
        -DownloadName "ffmpeg-release-essentials.zip" `
        -ExtractName "ffmpeg"
    $ffmpegExe = Get-ChildItem -LiteralPath $ffmpegRoot -Recurse -File -Filter "ffmpeg.exe" | Select-Object -First 1
    $ffprobeExe = Get-ChildItem -LiteralPath $ffmpegRoot -Recurse -File -Filter "ffprobe.exe" | Select-Object -First 1
    if (-not $ffmpegExe -or -not $ffprobeExe) {
        throw "FFmpeg archive does not contain both ffmpeg.exe and ffprobe.exe."
    }
    New-Item -ItemType Directory -Path (Join-Path $ffmpegTarget "bin") -Force | Out-Null
    Copy-Item -LiteralPath $ffmpegExe.FullName -Destination (Join-Path $ffmpegTarget "bin\ffmpeg.exe") -Force
    Copy-Item -LiteralPath $ffprobeExe.FullName -Destination (Join-Path $ffmpegTarget "bin\ffprobe.exe") -Force
    foreach ($name in @("LICENSE", "README.txt")) {
        $file = Get-ChildItem -LiteralPath $ffmpegRoot -Recurse -File -Filter $name | Select-Object -First 1
        if ($file) {
            Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $ffmpegTarget $name) -Force
        }
    }
}

$popplerTarget = Join-Path $thirdPartyRoot "poppler"
$popplerReady = @(
    (Join-Path $popplerTarget "pdftoppm.exe"),
    (Join-Path $popplerTarget "bin\pdftoppm.exe"),
    (Join-Path $popplerTarget "Library\bin\pdftoppm.exe")
) | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
if (-not $popplerReady) {
    $popplerRoot = Resolve-ArchiveOrDirectory `
        -Source $PopplerSource `
        -DefaultUrl $popplerUrl `
        -DownloadName "Poppler-$popplerVersion.zip" `
        -ExtractName "poppler" `
        -ExpectedSha256 $popplerSha256
    Copy-EngineRoot -SourceRoot $popplerRoot -ProbeName "pdftoppm.exe" -Destination $popplerTarget
}
if (-not (Test-Path -LiteralPath (Join-Path $popplerTarget "COPYING") -PathType Leaf)) {
    Get-VerifiedDownload `
        -Uri "https://gitlab.freedesktop.org/poppler/poppler/-/raw/master/COPYING" `
        -Destination (Join-Path $popplerTarget "COPYING") | Out-Null
}

$required = @(
    (Join-Path $ffmpegTarget "bin\ffmpeg.exe"),
    (Join-Path $ffmpegTarget "bin\ffprobe.exe")
)
$missing = @($required | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
$popplerAfter = @(
    (Join-Path $popplerTarget "pdftoppm.exe"),
    (Join-Path $popplerTarget "bin\pdftoppm.exe"),
    (Join-Path $popplerTarget "Library\bin\pdftoppm.exe")
) | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
if ($missing -or -not $popplerAfter) {
    throw "Portable dependency preparation did not produce a complete engine set."
}

Write-Host "Portable dependencies are ready under: $thirdPartyRoot"
Write-Host "WPS Office, Microsoft Office and LibreOffice are optional local installations and are intentionally not bundled."
