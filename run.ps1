$ErrorActionPreference = "Stop"
$projectRoot = (Resolve-Path -LiteralPath (Split-Path -Parent $MyInvocation.MyCommand.Path)).Path
$venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"

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

function Resolve-PythonCommand {
    if ((Test-Path -LiteralPath $venvPython) -and (Test-PythonCandidate -Executable $venvPython)) {
        return [PSCustomObject]@{
            Executable = $venvPython
            Arguments = [string[]]@()
            IsVirtualEnvironment = $true
        }
    }

    $candidates = @(
        [PSCustomObject]@{ Command = "py"; Arguments = @("-3") },
        [PSCustomObject]@{ Command = "python"; Arguments = @() }
    )
    foreach ($candidate in $candidates) {
        $applications = @(
            Get-Command $candidate.Command -CommandType Application -All -ErrorAction SilentlyContinue
        )
        foreach ($application in $applications) {
            if (Test-PythonCandidate -Executable $application.Source -PrefixArguments $candidate.Arguments) {
                return [PSCustomObject]@{
                    Executable = $application.Source
                    Arguments = [string[]]$candidate.Arguments
                    IsVirtualEnvironment = $false
                }
            }
        }
    }

    throw "A working Python 3.10 or newer was not found. Run install.ps1 after installing Python from python.org."
}

$python = Resolve-PythonCommand
$pythonExe = [string]$python.Executable
$pythonArgs = [string[]]$python.Arguments

Push-Location $projectRoot
try {
    & $pythonExe @pythonArgs -c "import tkinter, tkinterdnd2, PIL, pypdf, pdfplumber, pdf2docx, pymupdf, reportlab, docx, openpyxl, pptx, pdf2image, docuforge"
    if ($LASTEXITCODE -ne 0) {
        if ($python.IsVirtualEnvironment) {
            throw "LayoutLoom dependencies are missing from '$venvPython'. Run '$projectRoot\install.ps1' first."
        }
        throw "LayoutLoom dependencies are missing. Run '$projectRoot\install.ps1' first."
    }

    & $pythonExe @pythonArgs -m docuforge
    Assert-NativeSuccess "LayoutLoom exited with an error"
} finally {
    Pop-Location
}
