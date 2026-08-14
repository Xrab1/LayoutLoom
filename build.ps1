[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$projectRoot = (Resolve-Path -LiteralPath (Split-Path -Parent $MyInvocation.MyCommand.Path)).Path
$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"
$installScript = Join-Path $projectRoot "install.ps1"
$portableRoot = Join-Path $projectRoot "third_party"
$portableFfmpeg = Join-Path $portableRoot "ffmpeg"
$portablePoppler = Join-Path $portableRoot "poppler"

function Assert-NativeSuccess {
    param([string]$Message)
    if ($LASTEXITCODE -ne 0) {
        throw "$Message (exit code $LASTEXITCODE)."
    }
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

function Test-PythonInterpreter {
    param([string]$Executable)
    if (-not (Test-Path -LiteralPath $Executable)) {
        return $false
    }
    try {
        & $Executable -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" *> $null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

function Test-PythonModule {
    param([string]$ModuleName)
    try {
        & $pythonExe -c "import importlib, sys; importlib.import_module(sys.argv[1])" $ModuleName *> $null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

function New-TkRuntimeBackup {
    param(
        [string]$TclDirectory,
        [string]$TkDirectory,
        [string]$Destination
    )
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $destinationParent = Split-Path -Parent $Destination
    New-Item -ItemType Directory -Path $destinationParent -Force | Out-Null
    if (Test-Path -LiteralPath $Destination -PathType Leaf) {
        Remove-Item -LiteralPath $Destination -Force
    }
    $stream = [System.IO.File]::Open(
        $Destination,
        [System.IO.FileMode]::CreateNew,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None
    )
    $archive = [System.IO.Compression.ZipArchive]::new(
        $stream,
        [System.IO.Compression.ZipArchiveMode]::Create,
        $false
    )
    try {
        foreach ($mapping in @(
            [PSCustomObject]@{ Source = $TclDirectory; Root = "_tcl_data" },
            [PSCustomObject]@{ Source = $TkDirectory; Root = "_tk_data" }
        )) {
            $sourcePrefix = [System.IO.Path]::GetFullPath($mapping.Source).TrimEnd('\') + '\'
            foreach ($file in Get-ChildItem -LiteralPath $mapping.Source -Recurse -File) {
                if (-not $file.FullName.StartsWith($sourcePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                    throw "Unexpected Tcl/Tk runtime path: $($file.FullName)"
                }
                $relative = $file.FullName.Substring($sourcePrefix.Length).Replace('\', '/')
                $entryName = "$($mapping.Root)/$relative"
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                    $archive,
                    $file.FullName,
                    $entryName,
                    [System.IO.Compression.CompressionLevel]::Optimal
                ) | Out-Null
            }
        }
    } finally {
        $archive.Dispose()
        $stream.Dispose()
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

if (-not (Test-PythonInterpreter -Executable $pythonExe)) {
    if (Test-Path -LiteralPath (Join-Path $projectRoot ".venv")) {
        throw "The project virtual environment is invalid. Remove '$projectRoot\.venv' and run '$installScript' again."
    }

    & $installScript
    if (-not (Test-PythonInterpreter -Executable $pythonExe)) {
        throw "The installer did not create a working virtual environment."
    }
}

$requiredPortableFiles = @(
    (Join-Path $portableFfmpeg "bin\ffmpeg.exe"),
    (Join-Path $portableFfmpeg "bin\ffprobe.exe")
)
$popplerPair = @(
    (Join-Path $portablePoppler "pdftoppm.exe"),
    (Join-Path $portablePoppler "pdfinfo.exe")
)
if (-not (($popplerPair | Where-Object { Test-Path -LiteralPath $_ }).Count -eq 2)) {
    $popplerPair = @(
        (Join-Path $portablePoppler "bin\pdftoppm.exe"),
        (Join-Path $portablePoppler "bin\pdfinfo.exe")
    )
}
if (-not (($popplerPair | Where-Object { Test-Path -LiteralPath $_ }).Count -eq 2)) {
    $popplerPair = @(
        (Join-Path $portablePoppler "Library\bin\pdftoppm.exe"),
        (Join-Path $portablePoppler "Library\bin\pdfinfo.exe")
    )
}
$requiredPortableFiles += $popplerPair
$missingPortableFiles = @(
    $requiredPortableFiles | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) }
)
if ($missingPortableFiles) {
    throw (
        "Portable dependencies are incomplete. Run prepare_portable_dependencies.ps1 or provide: " +
        ($missingPortableFiles -join ", ")
    )
}

Push-Location $projectRoot
try {
    & $pythonExe -c "import tkinter, tkinterdnd2, wordninja, PIL, numpy, cv2, pypdf, pdfplumber, pdf2docx, pymupdf, reportlab, docx, openpyxl, pptx, pdf2image, win32com.client, pythoncom, pywintypes, docuforge"
    Assert-NativeSuccess "The pre-build core import check failed"

    & $pythonExe -m pip install "pyinstaller>=6.0"
    Assert-NativeSuccess "Failed to install PyInstaller"

    $tkRuntimeJson = (& $pythonExe -c "import json, pathlib, tkinter; tcl = pathlib.Path(tkinter.Tcl().eval('info library')).resolve(); tk = tcl.parent / ('tk' + str(tkinter.TkVersion)); print(json.dumps({'tcl': str(tcl), 'tk': str(tk)}))").Trim()
    Assert-NativeSuccess "Failed to locate the Python Tcl/Tk runtime"
    $tkRuntime = $tkRuntimeJson | ConvertFrom-Json
    foreach ($requiredTkFile in @(
        (Join-Path $tkRuntime.tcl "init.tcl"),
        (Join-Path $tkRuntime.tk "tk.tcl")
    )) {
        if (-not (Test-Path -LiteralPath $requiredTkFile -PathType Leaf)) {
            throw "The Python Tcl/Tk runtime is incomplete: $requiredTkFile"
        }
    }
    $tkBackupArchive = Join-Path $projectRoot "build\tk_runtime_backup.zip"
    New-TkRuntimeBackup -TclDirectory $tkRuntime.tcl -TkDirectory $tkRuntime.tk -Destination $tkBackupArchive
    if (-not (Test-Path -LiteralPath $tkBackupArchive -PathType Leaf)) {
        throw "Failed to create the Tcl/Tk recovery archive."
    }

    $buildArgs = @(
        "--noconfirm",
        "--clean",
        "--windowed",
        "--onedir",
        "--name", "LayoutLoom",
        "--additional-hooks-dir", (Join-Path $projectRoot "packaging_hooks"),
        "--add-data", ("{0};." -f $tkBackupArchive),
        "--collect-all", "PIL",
        "--collect-all", "cv2",
        "--collect-all", "numpy",
        "--collect-all", "pypdf",
        "--collect-all", "reportlab",
        "--collect-all", "pdfminer",
        "--collect-all", "pdfplumber",
        "--collect-all", "pdf2image",
        "--collect-all", "pdf2docx",
        "--collect-all", "pymupdf",
        "--collect-all", "docx",
        "--collect-all", "openpyxl",
        "--collect-all", "pptx",
        "--collect-all", "tkinterdnd2",
        "--upx-exclude", "libtkdnd*.dll",
        "--hidden-import", "PIL._tkinter_finder",
        "--hidden-import", "fitz"
    )

    $wordNinjaData = Join-Path $projectRoot ".venv\Lib\site-packages\wordninja"
    if (-not (Test-Path -LiteralPath (Join-Path $wordNinjaData "wordninja_words.txt.gz"))) {
        throw "The WordNinja language model is missing from the virtual environment."
    }
    $buildArgs += @("--add-data", ("{0};wordninja" -f $wordNinjaData))

    if (Test-PythonModule -ModuleName "win32com") {
        $buildArgs += @(
            "--collect-submodules", "win32com",
            "--hidden-import", "win32com.client",
            "--hidden-import", "pythoncom",
            "--hidden-import", "pywintypes"
        )
        if (Test-PythonModule -ModuleName "win32comext") {
            $buildArgs += @("--collect-submodules", "win32comext")
        }
        Write-Host "Collecting installed pywin32 COM modules."
    }

    if (Test-PythonModule -ModuleName "realesrgan_ncnn_py") {
        $buildArgs += @("--collect-all", "realesrgan_ncnn_py")
        Write-Host "Bundling Real-ESRGAN NCNN Vulkan models and GPU binding."
    } else {
        Write-Warning "Real-ESRGAN NCNN Vulkan is not installed; the packaged application will use the safe OpenCV fallback."
    }

    $portableRealEsrgan = Join-Path $projectRoot "third_party\realesrgan-ncnn-vulkan"
    if (
        (Test-Path -LiteralPath (Join-Path $portableRealEsrgan "realesrgan-ncnn-vulkan.exe")) -and
        (Test-Path -LiteralPath (Join-Path $portableRealEsrgan "models\realesrgan-x4plus.param")) -and
        (Test-Path -LiteralPath (Join-Path $portableRealEsrgan "models\realesrgan-x4plus.bin"))
    ) {
        $buildArgs += @("--add-data", ("{0};realesrgan-ncnn-vulkan" -f $portableRealEsrgan))
        Write-Host "Bundling the optional official Real-ESRGAN portable executable."
    }

    $entryPoint = Join-Path $projectRoot "launcher.py"
    & $pythonExe -m PyInstaller @buildArgs $entryPoint
    Assert-NativeSuccess "PyInstaller build failed"

    $bundleDirectory = Join-Path $projectRoot "dist\LayoutLoom"
    $exePath = Join-Path $bundleDirectory "LayoutLoom.exe"
    if (-not (Test-Path -LiteralPath $exePath -PathType Leaf)) {
        throw "The build completed without producing '$exePath'."
    }
    $exeInfo = Get-Item -LiteralPath $exePath
    if ($exeInfo.Length -le 0) {
        throw "The generated executable is empty: '$exePath'."
    }
    if (-not (Test-Path -LiteralPath $bundleDirectory -PathType Container)) {
        throw "The one-folder bundle directory is missing: '$bundleDirectory'."
    }
    $bundleInternal = Join-Path $bundleDirectory "_internal"
    foreach ($requiredTkRuntimeFile in @(
        (Join-Path $bundleInternal "_tcl_data\init.tcl"),
        (Join-Path $bundleInternal "_tk_data\tk.tcl"),
        (Join-Path $bundleInternal "_tkinter.pyd"),
        (Join-Path $bundleInternal "tcl86t.dll"),
        (Join-Path $bundleInternal "tk86t.dll"),
        (Join-Path $bundleInternal "tk_runtime_backup.zip")
    )) {
        if (-not (Test-Path -LiteralPath $requiredTkRuntimeFile -PathType Leaf)) {
            throw "The packaged Tcl/Tk runtime is incomplete: $requiredTkRuntimeFile"
        }
    }

    $bundleFfmpeg = Join-Path $bundleDirectory "ffmpeg"
    $bundlePoppler = Join-Path $bundleDirectory "poppler"
    New-Item -ItemType Directory -Path (Join-Path $bundleFfmpeg "bin") -Force | Out-Null
    foreach ($name in @("ffmpeg.exe", "ffprobe.exe")) {
        Copy-Item -LiteralPath (Join-Path $portableFfmpeg "bin\$name") -Destination (Join-Path $bundleFfmpeg "bin\$name") -Force
    }
    foreach ($name in @("LICENSE", "README.txt")) {
        $source = Join-Path $portableFfmpeg $name
        if (Test-Path -LiteralPath $source -PathType Leaf) {
            Copy-Item -LiteralPath $source -Destination (Join-Path $bundleFfmpeg $name) -Force
        }
    }
    Copy-Item -LiteralPath $portablePoppler -Destination $bundlePoppler -Recurse -Force

    foreach ($name in @("LICENSE", "README.md", "SOURCE_CODE.md", "THIRD_PARTY_NOTICES.md")) {
        Copy-Item -LiteralPath (Join-Path $projectRoot $name) -Destination (Join-Path $bundleDirectory $name) -Force
    }
    $thirdPartyLicenses = Join-Path $bundleDirectory "THIRD_PARTY_LICENSES"
    New-Item -ItemType Directory -Path $thirdPartyLicenses -Force | Out-Null
    $licenseCopies = @(
        [PSCustomObject]@{ Source = (Join-Path $portableFfmpeg "LICENSE"); Target = "FFmpeg-GPLv3.txt" },
        [PSCustomObject]@{ Source = (Join-Path $portableFfmpeg "README.txt"); Target = "FFmpeg-build-info.txt" },
        [PSCustomObject]@{ Source = (Join-Path $portablePoppler "COPYING"); Target = "Poppler-GPLv2.txt" },
        [PSCustomObject]@{ Source = (Join-Path $portablePoppler "manifest.json"); Target = "Poppler-build-manifest.json" }
    )
    foreach ($copy in $licenseCopies) {
        if (Test-Path -LiteralPath $copy.Source -PathType Leaf) {
            Copy-Item -LiteralPath $copy.Source -Destination (Join-Path $thirdPartyLicenses $copy.Target) -Force
        }
    }
    $pythonLicenseDirectory = Join-Path $thirdPartyLicenses "Python-packages"
    New-Item -ItemType Directory -Path $pythonLicenseDirectory -Force | Out-Null
    $sitePackages = Join-Path $projectRoot ".venv\Lib\site-packages"
    foreach ($metadataDirectory in Get-ChildItem -LiteralPath $sitePackages -Directory -Filter "*.dist-info") {
        $licenseFiles = @(
            Get-ChildItem -LiteralPath $metadataDirectory.FullName -File |
                Where-Object { $_.Name -match '^(LICENSE|LICENCE|COPYING|NOTICE|AUTHORS)' }
        )
        if (-not $licenseFiles) {
            $licensesSubdirectory = Join-Path $metadataDirectory.FullName "licenses"
            if (Test-Path -LiteralPath $licensesSubdirectory -PathType Container) {
                $licenseFiles = @(Get-ChildItem -LiteralPath $licensesSubdirectory -Recurse -File)
            }
        }
        if ($licenseFiles) {
            $packageLicenseDirectory = Join-Path $pythonLicenseDirectory $metadataDirectory.Name
            New-Item -ItemType Directory -Path $packageLicenseDirectory -Force | Out-Null
            foreach ($licenseFile in $licenseFiles) {
                Copy-Item -LiteralPath $licenseFile.FullName -Destination (Join-Path $packageLicenseDirectory $licenseFile.Name) -Force
            }
        }
    }
    $pythonBase = (& $pythonExe -c "import sys; print(sys.base_prefix)").Trim()
    foreach ($name in @("LICENSE.txt", "LICENSE")) {
        $source = Join-Path $pythonBase $name
        if (Test-Path -LiteralPath $source -PathType Leaf) {
            Copy-Item -LiteralPath $source -Destination (Join-Path $thirdPartyLicenses "Python-runtime-$name") -Force
            break
        }
    }

    $savedTkEnvironment = @{}
    foreach ($name in @("TCL_LIBRARY", "TK_LIBRARY", "LOCALAPPDATA")) {
        $savedTkEnvironment[$name] = [System.Environment]::GetEnvironmentVariable($name, "Process")
    }
    [System.Environment]::SetEnvironmentVariable("TCL_LIBRARY", $null, "Process")
    [System.Environment]::SetEnvironmentVariable("TK_LIBRARY", $null, "Process")
    # Give the two independent gates separate cache roots so diagnostic and
    # recovery state from one frozen launch cannot affect the other.
    $primarySelfTestLocalAppData = Join-Path $projectRoot ("build\tk-primary-self-test-" + [guid]::NewGuid().ToString("N"))
    $recoverySelfTestLocalAppData = Join-Path $projectRoot ("build\tk-recovery-self-test-" + [guid]::NewGuid().ToString("N"))
    foreach ($selfTestLocalAppData in @($primarySelfTestLocalAppData, $recoverySelfTestLocalAppData)) {
        New-Item -ItemType Directory -Path $selfTestLocalAppData -Force | Out-Null
    }
    [System.Environment]::SetEnvironmentVariable("LOCALAPPDATA", $primarySelfTestLocalAppData, "Process")
    $primaryTcl = Join-Path $bundleInternal "_tcl_data"
    $primaryTk = Join-Path $bundleInternal "_tk_data"
    $heldTcl = Join-Path $bundleInternal "_tcl_data.primary-validation"
    $heldTk = Join-Path $bundleInternal "_tk_data.primary-validation"
    try {
        Invoke-FrozenSelfTest -Executable $exePath -Description "The generated executable self-test" -EnvironmentOverrides @{ LOCALAPPDATA = $primarySelfTestLocalAppData }
        Move-Item -LiteralPath $primaryTcl -Destination $heldTcl
        Move-Item -LiteralPath $primaryTk -Destination $heldTk
        Invoke-FrozenSelfTest -Executable $exePath -Description "The Tcl/Tk recovery self-test" -TimeoutMilliseconds 180000 -EnvironmentOverrides @{ LOCALAPPDATA = $recoverySelfTestLocalAppData }
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
        $buildRoot = [System.IO.Path]::GetFullPath((Join-Path $projectRoot "build")).TrimEnd('\') + '\'
        foreach ($selfTestLocalAppData in @($primarySelfTestLocalAppData, $recoverySelfTestLocalAppData)) {
            $cachePath = [System.IO.Path]::GetFullPath($selfTestLocalAppData)
            if ($cachePath.StartsWith($buildRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
                Remove-Item -LiteralPath $cachePath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
    Write-Host "Executable and Tcl/Tk recovery self-tests passed."
} finally {
    Pop-Location
}

Write-Host ("Build complete: {0}" -f $exePath)
Write-Host ("Copy the complete application folder: {0}" -f $bundleDirectory)
