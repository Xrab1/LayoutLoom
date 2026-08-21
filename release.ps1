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

function Stop-FrozenProcessTree {
    param([System.Diagnostics.Process]$Process)
    try {
        if ($Process.HasExited) {
            return
        }
        $taskkill = Join-Path $env:SystemRoot "System32\taskkill.exe"
        if (Test-Path -LiteralPath $taskkill -PathType Leaf) {
            & $taskkill /PID $Process.Id /T /F *> $null
        } else {
            Stop-Process -Id $Process.Id -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Stop-Process -Id $Process.Id -Force -ErrorAction SilentlyContinue
    } finally {
        try {
            $Process.WaitForExit(5000) | Out-Null
        } catch {
        }
    }
}

function Invoke-FrozenSelfTest {
    param(
        [string]$Executable,
        [string]$Description,
        [int]$TimeoutMilliseconds = 60000,
        [hashtable]$EnvironmentOverrides = @{}
    )
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $Executable
    $startInfo.Arguments = "--self-test"
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    foreach ($entry in $EnvironmentOverrides.GetEnumerator()) {
        $startInfo.EnvironmentVariables[$entry.Key] = [string]$entry.Value
    }
    $selfTestErrorFile = $null
    if ($EnvironmentOverrides.ContainsKey("LOCALAPPDATA")) {
        $selfTestErrorFile = Join-Path ([string]$EnvironmentOverrides["LOCALAPPDATA"]) "layoutloom-self-test-error.log"
        if (Test-Path -LiteralPath $selfTestErrorFile -PathType Leaf) {
            Remove-Item -LiteralPath $selfTestErrorFile -Force
        }
        $startInfo.EnvironmentVariables["LAYOUTLOOM_SELF_TEST_ERROR_FILE"] = $selfTestErrorFile
    }
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) {
            throw "$Description could not be started."
        }
        if (-not $process.WaitForExit($TimeoutMilliseconds)) {
            Stop-FrozenProcessTree -Process $process
            $timeoutSeconds = [math]::Ceiling($TimeoutMilliseconds / 1000)
            throw "$Description did not finish within $timeoutSeconds seconds."
        }
        $process.Refresh()
        if ($process.ExitCode -ne 0) {
            $diagnostic = ""
            if ($selfTestErrorFile -and (Test-Path -LiteralPath $selfTestErrorFile -PathType Leaf)) {
                $diagnosticText = (Get-Content -Raw -LiteralPath $selfTestErrorFile).Trim()
                if ($diagnosticText) {
                    $diagnostic = "`n$diagnosticText"
                }
            }
            throw "$Description failed with exit code $($process.ExitCode).$diagnostic"
        }
    } finally {
        $process.Dispose()
    }
}

function Invoke-FrozenCliProbe {
    param(
        [string]$Executable,
        [string]$Arguments,
        [string]$Description,
        [int]$TimeoutMilliseconds = 60000
    )
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $Executable
    $startInfo.Arguments = $Arguments
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $startInfo.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) {
            throw "$Description could not be started."
        }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutMilliseconds)) {
            Stop-FrozenProcessTree -Process $process
            throw "$Description timed out."
        }
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()
        $process.Refresh()
        if ($process.ExitCode -ne 0) {
            throw "$Description failed with exit code $($process.ExitCode).`n$stderr"
        }
        return $stdout.Trim()
    } finally {
        $process.Dispose()
    }
}

function Assert-PortableArchiveRuntime {
    param([string]$ArchivePath)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
    try {
        $entries = @{}
        foreach ($entry in $archive.Entries) {
            $entries[$entry.FullName.Replace('\', '/')] = $entry
        }
        $requiredEntries = @(
            "LayoutLoom/_internal/_tcl_data/init.tcl",
            "LayoutLoom/_internal/_tk_data/tk.tcl",
            "LayoutLoom/_internal/_tkinter.pyd",
            "LayoutLoom/_internal/tcl86t.dll",
            "LayoutLoom/_internal/tk86t.dll",
            "LayoutLoom/_internal/tk_runtime_backup.zip",
            "LayoutLoom/LayoutLoom-CLI.exe",
            "LayoutLoom/AGENT_INTEGRATION.md",
            "LayoutLoom/agent_skill/layoutloom-agent/SKILL.md",
            "LayoutLoom/agent_skill/layoutloom-agent/agents/openai.yaml",
            "LayoutLoom/agent_skill/layoutloom-agent/scripts/layoutloom_agent.py",
            "LayoutLoom/agent_skill/layoutloom-agent/references/protocol.md"
        )
        foreach ($requiredEntry in $requiredEntries) {
            if (-not $entries.ContainsKey($requiredEntry)) {
                throw "The portable archive is missing required runtime entry: $requiredEntry"
            }
        }
        $backupEntry = $entries["LayoutLoom/_internal/tk_runtime_backup.zip"]
        $memory = New-Object System.IO.MemoryStream
        $backupStream = $backupEntry.Open()
        try {
            $backupStream.CopyTo($memory)
        } finally {
            $backupStream.Dispose()
        }
        $memory.Position = 0
        $backupArchive = [System.IO.Compression.ZipArchive]::new(
            $memory,
            [System.IO.Compression.ZipArchiveMode]::Read,
            $false
        )
        try {
            $backupNames = @($backupArchive.Entries | ForEach-Object { $_.FullName.Replace('\', '/') })
            foreach ($requiredBackup in @("_tcl_data/init.tcl", "_tk_data/tk.tcl")) {
                if ($requiredBackup -notin $backupNames) {
                    throw "The Tcl/Tk recovery archive is missing: $requiredBackup"
                }
            }
        } finally {
            $backupArchive.Dispose()
            $memory.Dispose()
        }
    } finally {
        $archive.Dispose()
    }
}

function Assert-SourceArchiveContents {
    param(
        [string]$ArchivePath,
        [string]$Version
    )
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
    try {
        $entries = @{}
        foreach ($entry in $archive.Entries) {
            $entries[$entry.FullName.Replace('\', '/')] = $entry
        }
        $root = "LayoutLoom-Source-$Version"
        $requiredEntries = @(
            "$root/AGENT_INTEGRATION.md",
            "$root/agent_launcher.py",
            "$root/docuforge/agent_api.py",
            "$root/docuforge/cli.py",
            "$root/integrations/codex/layoutloom-agent/SKILL.md",
            "$root/integrations/codex/layoutloom-agent/agents/openai.yaml",
            "$root/integrations/codex/layoutloom-agent/scripts/layoutloom_agent.py",
            "$root/integrations/codex/layoutloom-agent/references/protocol.md",
            "$root/tests/test_agent_api.py",
            "$root/tests/test_agent_cli.py"
        )
        foreach ($requiredEntry in $requiredEntries) {
            if (-not $entries.ContainsKey($requiredEntry)) {
                throw "The source archive is missing required Agent integration entry: $requiredEntry"
            }
        }
    } finally {
        $archive.Dispose()
    }
}

function Test-ExtractedPortableRuntime {
    param(
        [string]$ArchivePath,
        [string]$ValidationDirectory,
        [string]$ExpectedParent
    )
    Remove-WorkspaceChild -Path $ValidationDirectory -ExpectedParent $ExpectedParent
    New-Item -ItemType Directory -Path $ValidationDirectory | Out-Null
    Expand-Archive -LiteralPath $ArchivePath -DestinationPath $ValidationDirectory -Force
    $bundle = Join-Path $ValidationDirectory "LayoutLoom"
    $executable = Join-Path $bundle "LayoutLoom.exe"
    $cliExecutable = Join-Path $bundle "LayoutLoom-CLI.exe"
    $internal = Join-Path $bundle "_internal"
    $primaryTcl = Join-Path $internal "_tcl_data"
    $primaryTk = Join-Path $internal "_tk_data"
    $heldTcl = Join-Path $internal "_tcl_data.primary-validation"
    $heldTk = Join-Path $internal "_tk_data.primary-validation"
    if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) {
        throw "The extracted portable archive does not contain LayoutLoom.exe."
    }
    if (-not (Test-Path -LiteralPath $cliExecutable -PathType Leaf)) {
        throw "The extracted portable archive does not contain LayoutLoom-CLI.exe."
    }
    $savedTkEnvironment = @{}
    foreach ($name in @("TCL_LIBRARY", "TK_LIBRARY", "LOCALAPPDATA")) {
        $savedTkEnvironment[$name] = [System.Environment]::GetEnvironmentVariable($name, "Process")
    }
    [System.Environment]::SetEnvironmentVariable("TCL_LIBRARY", $null, "Process")
    [System.Environment]::SetEnvironmentVariable("TK_LIBRARY", $null, "Process")
    $primaryLocalAppData = Join-Path $ValidationDirectory "primary-local-app-data"
    $recoveryLocalAppData = Join-Path $ValidationDirectory "recovery-local-app-data"
    foreach ($isolatedLocalAppData in @($primaryLocalAppData, $recoveryLocalAppData)) {
        New-Item -ItemType Directory -Path $isolatedLocalAppData -Force | Out-Null
    }
    [System.Environment]::SetEnvironmentVariable("LOCALAPPDATA", $primaryLocalAppData, "Process")
    try {
        Invoke-FrozenSelfTest -Executable $executable -Description "The extracted portable self-test" -EnvironmentOverrides @{ LOCALAPPDATA = $primaryLocalAppData }
        $protocol = Invoke-FrozenCliProbe -Executable $cliExecutable -Arguments "agent protocol" -Description "The extracted Agent CLI protocol probe" | ConvertFrom-Json
        if (
            $protocol.protocol.name -ne "layoutloom-agent" -or
            $protocol.protocol.version -ne "1.0" -or
            $protocol.commands -notcontains "quick-run"
        ) {
            throw "The extracted Agent CLI returned an unexpected protocol descriptor."
        }
        Move-Item -LiteralPath $primaryTcl -Destination $heldTcl
        Move-Item -LiteralPath $primaryTk -Destination $heldTk
        Invoke-FrozenSelfTest -Executable $executable -Description "The extracted Tcl/Tk recovery self-test" -TimeoutMilliseconds 180000 -EnvironmentOverrides @{ LOCALAPPDATA = $recoveryLocalAppData }
    } finally {
        if ((Test-Path -LiteralPath $heldTcl) -and -not (Test-Path -LiteralPath $primaryTcl)) {
            Move-Item -LiteralPath $heldTcl -Destination $primaryTcl
        }
        if ((Test-Path -LiteralPath $heldTk) -and -not (Test-Path -LiteralPath $primaryTk)) {
            Move-Item -LiteralPath $heldTk -Destination $primaryTk
        }
        foreach ($name in @("TCL_LIBRARY", "TK_LIBRARY", "LOCALAPPDATA")) {
            [System.Environment]::SetEnvironmentVariable($name, $savedTkEnvironment[$name], "Process")
        }
    }
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
$cliExePath = Join-Path $bundleDirectory "LayoutLoom-CLI.exe"
if (-not (Test-Path -LiteralPath $exePath -PathType Leaf)) {
    throw "Portable bundle is missing: $exePath"
}
if (-not (Test-Path -LiteralPath $cliExePath -PathType Leaf)) {
    throw "Portable Agent CLI is missing: $cliExePath"
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
$sourceDirectories = @("docuforge", "tests", "packaging_hooks", "integrations")
foreach ($name in $sourceDirectories) {
    Copy-Item -LiteralPath (Join-Path $projectRoot $name) -Destination $sourceStage -Recurse -Force
}
$sourceFiles = @(
    ".gitignore",
    "LICENSE",
    "README.md",
    "AGENT_INTEGRATION.md",
    "CHANGELOG.md",
    "SOURCE_CODE.md",
    "THIRD_PARTY_NOTICES.md",
    "pyproject.toml",
    "requirements.txt",
    "requirements-dev.txt",
    "launcher.py",
    "agent_launcher.py",
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
foreach ($requiredSource in @(
    "docuforge\app.py",
    "docuforge\agent_api.py",
    "docuforge\cli.py",
    "docuforge\registry.py",
    "tests\test_agent_api.py",
    "tests\test_agent_cli.py",
    "tests\test_core.py",
    "integrations\codex\layoutloom-agent\SKILL.md",
    "integrations\codex\layoutloom-agent\agents\openai.yaml",
    "integrations\codex\layoutloom-agent\scripts\layoutloom_agent.py",
    "integrations\codex\layoutloom-agent\references\protocol.md",
    "agent_launcher.py",
    "AGENT_INTEGRATION.md",
    "LICENSE",
    "pyproject.toml"
)) {
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

Assert-PortableArchiveRuntime -ArchivePath $portableZip
Assert-SourceArchiveContents -ArchivePath $sourceZip -Version $version
$portableValidation = Join-Path $stageRoot "portable-runtime-validation"
Test-ExtractedPortableRuntime -ArchivePath $portableZip -ValidationDirectory $portableValidation -ExpectedParent $stageRoot
Remove-WorkspaceChild -Path $portableValidation -ExpectedParent $stageRoot
Write-Host "Portable runtime and source archive Agent integration self-tests passed."

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
