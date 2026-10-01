#Stop immediately when a build step fails so an incomplete executable is not reported as finished.
$ErrorActionPreference = "Stop"

#Run every relative path from the folder containing this script.
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $projectRoot

#Prefer the Python Launcher so the script can select a supported Python version explicitly.
$pythonLauncher = Get-Command py -ErrorAction SilentlyContinue
if ($pythonLauncher) {
    & $pythonLauncher.Source -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"
    if ($LASTEXITCODE -eq 0) {
        $pythonCommand = $pythonLauncher.Source
        $pythonPrefix = @("-3")
    }
}

#If the launcher is unavailable, look for a normal python command instead.
if (-not $pythonCommand) {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        & $python.Source -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"
        if ($LASTEXITCODE -eq 0) {
            $pythonCommand = $python.Source
            $pythonPrefix = @()
        }
    }
}

#Stop with setup guidance instead of failing later with a confusing build error.
if (-not $pythonCommand) {
    throw "Python 3.10 or newer is required to build the executable. Install Python and run this script again. The finished executable does not require Python."
}

#Create a project-local build environment so global Python packages are not changed.
$environment = Join-Path $projectRoot ".build-venv"
$environmentPython = Join-Path $environment "Scripts\python.exe"
if (-not (Test-Path -LiteralPath $environmentPython)) {
    & $pythonCommand @pythonPrefix -m venv $environment
    if ($LASTEXITCODE -ne 0) { throw "Could not create the build environment." }
}

#Check Tk support up front because the app's desktop window depends on Tcl/Tk.
& $environmentPython -c "import tkinter; tkinter.Tcl().eval('info patchlevel')"
if ($LASTEXITCODE -ne 0) {
    throw "This Python installation does not include working Tkinter/Tcl-Tk support. Repair Python with the Tcl/Tk feature enabled, then run this script again."
}

#Install build tools and app libraries inside the isolated build environment.
& $environmentPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "Could not prepare pip in the build environment." }

& $environmentPython -m pip install -r (Join-Path $projectRoot "requirements.txt") pyinstaller
if ($LASTEXITCODE -ne 0) { throw "Could not install the app's build dependencies." }

#Keep the executable, temporary work files, and PyInstaller spec together under build.
$outputRoot = Join-Path $projectRoot "build"
$distPath = Join-Path $outputRoot "dist"
$workPath = Join-Path $outputRoot "work"
$specPath = Join-Path $outputRoot "spec"

#Package Python, Tk, and the app into one GUI executable for Windows users.
& $environmentPython -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --icon (Join-Path $projectRoot "app_icon.ico") `
    --add-data "$(Join-Path $projectRoot 'app_icon.ico');." `
    --add-data "$(Join-Path $projectRoot 'install_update.ps1');." `
    --hidden-import=tkinter `
    --hidden-import=tkinter.ttk `
    --hidden-import=tkinter.filedialog `
    --hidden-import=tkinter.messagebox `
    --hidden-import=_tkinter `
    --name FantasyTradeCalculator `
    --distpath $distPath `
    --workpath $workPath `
    --specpath $specPath `
    (Join-Path $projectRoot "sleeper_exporter.py")

if ($LASTEXITCODE -ne 0) { throw "PyInstaller could not build the executable." }

Write-Host "Build complete: $distPath\FantasyTradeCalculator.exe"
