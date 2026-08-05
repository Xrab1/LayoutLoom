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

    $buildArgs = @(
        "--noconfirm",
        "--clean",
        "--windowed",
        "--onedir",
        "--name", "LayoutLoom",
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

    $smokeProcess = Start-Process -FilePath $exePath -ArgumentList "--self-test" -PassThru -WindowStyle Hidden
    if (-not $smokeProcess.WaitForExit(60000)) {
        try {
            Stop-Process -Id $smokeProcess.Id -Force -ErrorAction SilentlyContinue
        } catch {
        }
        throw "The generated executable did not finish its self-test within 60 seconds."
    }
    $smokeProcess.Refresh()
    if ($smokeProcess.ExitCode -ne 0) {
        throw "The generated executable failed its self-test with exit code $($smokeProcess.ExitCode)."
    }
    Write-Host "Executable self-test passed."
} finally {
    Pop-Location
}

Write-Host ("Build complete: {0}" -f $exePath)
Write-Host ("Copy the complete application folder: {0}" -f $bundleDirectory)
