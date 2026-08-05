[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$projectRoot = (Resolve-Path -LiteralPath (Split-Path -Parent $MyInvocation.MyCommand.Path)).Path
$venvDir = Join-Path $projectRoot ".venv"

function Assert-NativeSuccess {
    param([string]$Message)
    if ($LASTEXITCODE -ne 0) {
        throw "$Message (exit code $LASTEXITCODE)."
    }
}

function Test-PythonCandidate {
    param(
        [string]$Executable,
        [string[]]$PrefixArguments = @()
    )
    try {
        & $Executable @PrefixArguments -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" *> $null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

function Resolve-BasePython {
    $candidates = @(
        [PSCustomObject]@{ Name = "py -3"; Command = "py"; Arguments = @("-3") },
        [PSCustomObject]@{ Name = "python"; Command = "python"; Arguments = @() }
    )

    foreach ($candidate in $candidates) {
        $applications = @(
            Get-Command $candidate.Command -CommandType Application -All -ErrorAction SilentlyContinue
        )
        foreach ($application in $applications) {
            $executable = $application.Source
            $valid = Test-PythonCandidate -Executable $executable -PrefixArguments $candidate.Arguments
            if ($valid) {
                return [PSCustomObject]@{
                    Name = $candidate.Name
                    Executable = $executable
                    Arguments = [string[]]$candidate.Arguments
                }
            }
        }
    }

    throw "A working Python 3.10 or newer was not found. Install Python from python.org and enable the Python launcher."
}

$basePython = $null

Push-Location $projectRoot
try {
    if (-not (Test-Path -LiteralPath $venvDir)) {
        $basePython = Resolve-BasePython
        $basePythonExe = [string]$basePython.Executable
        $basePythonArgs = [string[]]$basePython.Arguments
        & $basePythonExe @basePythonArgs -m venv $venvDir
        Assert-NativeSuccess "Failed to create the virtual environment"
    }

    $pythonExe = Join-Path $venvDir "Scripts\python.exe"
    if (-not (Test-Path -LiteralPath $pythonExe)) {
        throw "The virtual environment is incomplete. Remove '$venvDir' and run this installer again."
    }
    if (-not (Test-PythonCandidate -Executable $pythonExe)) {
        throw "The virtual environment does not contain a working Python 3.10 or newer. Remove '$venvDir' and run this installer again."
    }

    & $pythonExe -m pip install --upgrade pip
    Assert-NativeSuccess "Failed to upgrade pip"
    & $pythonExe -m pip install -r (Join-Path $projectRoot "requirements.txt")
    Assert-NativeSuccess "Failed to install LayoutLoom dependencies"
    if ($env:OS -eq "Windows_NT") {
        & $pythonExe -m pip install --no-deps "realesrgan-ncnn-py==2.0.0"
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Real-ESRGAN NCNN Vulkan could not be installed. LayoutLoom will keep the safe OpenCV enhancement fallback."
        } else {
            Write-Host "Installed the portable Real-ESRGAN NCNN Vulkan GPU engine."
        }
    }
    & $pythonExe -m pip install -e $projectRoot
    Assert-NativeSuccess "Failed to install LayoutLoom"

    & $pythonExe -c "import tkinter, tkinterdnd2, PIL, pypdf, pdfplumber, pdf2docx, pymupdf, reportlab, docx, openpyxl, pptx, pdf2image, win32com.client, pythoncom, pywintypes, docuforge"
    Assert-NativeSuccess "The core dependency import check failed"
} finally {
    Pop-Location
}

$runScript = Join-Path $projectRoot "run.ps1"
$pythonSource = if ($null -eq $basePython) { "the existing virtual environment" } else { $basePython.Name }
Write-Host ("Installation complete using {0}." -f $pythonSource)
Write-Host ("Start LayoutLoom with: {0}" -f $runScript)
