param(
    [string]$PythonExe = ""
)

$ErrorActionPreference = "Stop"
$projectRoot = $PSScriptRoot
$venvPath = Join-Path $projectRoot ".venv-build"
$distPath = Join-Path $projectRoot "dist"
$workPath = Join-Path $projectRoot "build\pyinstaller\work"
$specPath = Join-Path $projectRoot "build\pyinstaller\spec"

Push-Location $projectRoot
try {
    if ($PythonExe) {
        if (-not (Test-Path -LiteralPath $PythonExe -PathType Leaf)) {
            throw "Python executable not found: $PythonExe"
        }
        $pythonCommand = (Resolve-Path -LiteralPath $PythonExe).Path
        $pythonPrefix = @()
    } elseif (Get-Command py -ErrorAction SilentlyContinue) {
        $pythonCommand = (Get-Command py).Source
        $pythonPrefix = @("-3.12")
    } elseif (Get-Command python -ErrorAction SilentlyContinue) {
        $pythonCommand = (Get-Command python).Source
        $pythonPrefix = @()
    } else {
        throw "Python 3.11+ is required to build. Install Python or pass -PythonExe."
    }

    $pythonVersion = (& $pythonCommand @pythonPrefix --version 2>&1 | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -or $pythonVersion -notmatch "Python (\d+)\.(\d+)") {
        throw "Unable to verify Python at '$pythonCommand'. Install Python 3.11+ and ensure 'py' or 'python' works, or pass -PythonExe with the full path to python.exe. Details: $pythonVersion"
    }
    $versionMajor = [int]$Matches[1]
    $versionMinor = [int]$Matches[2]
    if ($versionMajor -lt 3 -or ($versionMajor -eq 3 -and $versionMinor -lt 11)) {
        throw "Python 3.11 or newer is required; found $pythonVersion"
    }

    if (-not (Test-Path -LiteralPath $venvPath -PathType Container)) {
        & $pythonCommand @pythonPrefix -m venv $venvPath
        if ($LASTEXITCODE -ne 0) { throw "Failed to create the isolated build environment." }
    }

    $venvPython = Join-Path $venvPath "Scripts\python.exe"
    if (-not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
        throw "The build environment is incomplete. Remove only .venv-build and run this script again."
    }

    & $venvPython -m pip install --disable-pip-version-check -r requirements.txt -r requirements-build.txt
    if ($LASTEXITCODE -ne 0) { throw "Failed to install the declared build dependencies." }

    New-Item -ItemType Directory -Path $distPath, $workPath, $specPath -Force | Out-Null
    $entryPoint = Join-Path $projectRoot "run_freight_calculator.py"
    & $venvPython -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --windowed `
        --name FreightCalculator `
        --distpath $distPath `
        --workpath $workPath `
        --specpath $specPath `
        $entryPoint
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller did not complete successfully." }

    $executablePath = Join-Path $distPath "FreightCalculator.exe"
    if (-not (Test-Path -LiteralPath $executablePath -PathType Leaf)) {
        throw "Build completed without producing FreightCalculator.exe."
    }
    $executable = Get-Item -LiteralPath $executablePath
    if ($executable.Length -lt 1MB) {
        throw "The generated executable is unexpectedly small ($($executable.Length) bytes)."
    }
    Write-Host "Portable executable created: $executablePath"
    Write-Host ("Size: {0:N1} MB" -f ($executable.Length / 1MB))
} finally {
    Pop-Location
}
