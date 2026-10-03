param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath
)

$ErrorActionPreference = "Stop"

function Show-UpdateMessage([string]$Message, [string]$Title, [int]$Buttons = 0x10) {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($Message, $Title, $Buttons) | Out-Null
    }
    catch {
        #The app already closed, so keep the helper failure visible in its own console if needed.
        Write-Error $Message
    }
}

$stagingDirectory = $null
$backupPath = $null

try {
    $config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
    $processId = [int]$config.process_id
    $executable = [System.IO.Path]::GetFullPath([string]$config.executable)
    $archive = [System.IO.Path]::GetFullPath([string]$config.archive)
    $helperScript = [System.IO.Path]::GetFullPath([string]$config.helper_script)

    if ([System.IO.Path]::GetFileName($executable) -ne "FantasyTradeCalculator.exe") {
        throw "The current app file does not have the expected name."
    }
    if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) {
        throw "The current app file could not be found."
    }

    $runningApp = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($runningApp) {
        Wait-Process -Id $processId -Timeout 120 -ErrorAction Stop
    }

    $stagingDirectory = Join-Path $env:TEMP ("FantasyTradeCalculator-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $stagingDirectory | Out-Null
    Expand-Archive -LiteralPath $archive -DestinationPath $stagingDirectory -Force

    $topLevelFiles = @(Get-ChildItem -LiteralPath $stagingDirectory -File)
    $topLevelDirectories = @(Get-ChildItem -LiteralPath $stagingDirectory -Directory)
    if ($topLevelFiles.Count -ne 1 -or $topLevelFiles[0].Name -ne "FantasyTradeCalculator.exe" -or $topLevelDirectories.Count -ne 0) {
        throw "The update archive did not contain exactly the expected app file."
    }

    $newExecutable = $topLevelFiles[0].FullName
    if ($topLevelFiles[0].Length -lt 1000000) {
        throw "The downloaded app file is unexpectedly small."
    }

    $replacementPath = "$executable.update-$processId"
    $backupPath = "$executable.backup-$processId"
    Copy-Item -LiteralPath $newExecutable -Destination $replacementPath
    Move-Item -LiteralPath $executable -Destination $backupPath
    try {
        Move-Item -LiteralPath $replacementPath -Destination $executable
        #Launch the replacement as a fresh one-file PyInstaller instance instead of reusing
        #the old app's temporary extraction directory, which is removed as that app exits.
        $previousResetEnvironment = $env:PYINSTALLER_RESET_ENVIRONMENT
        try {
            $env:PYINSTALLER_RESET_ENVIRONMENT = "1"
            Start-Process -FilePath $executable -WorkingDirectory ([System.IO.Path]::GetDirectoryName($executable))
        }
        finally {
            if ($null -eq $previousResetEnvironment) {
                Remove-Item Env:PYINSTALLER_RESET_ENVIRONMENT -ErrorAction SilentlyContinue
            }
            else {
                $env:PYINSTALLER_RESET_ENVIRONMENT = $previousResetEnvironment
            }
        }
    }
    catch {
        if (Test-Path -LiteralPath $executable -PathType Leaf) {
            Remove-Item -LiteralPath $executable -Force
        }
        Move-Item -LiteralPath $backupPath -Destination $executable -Force
        $backupPath = $null
        throw
    }

    if ($backupPath -and (Test-Path -LiteralPath $backupPath)) {
        Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
        $backupPath = $null
    }
    Remove-Item -LiteralPath $archive -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $ConfigPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    Remove-Item -LiteralPath $helperScript -Force -ErrorAction SilentlyContinue
}
catch {
    $message = "The update could not be installed. Your existing app should still be available.`n`n$($_.Exception.Message)"
    if ($backupPath -and (Test-Path -LiteralPath $backupPath) -and -not (Test-Path -LiteralPath $executable)) {
        Move-Item -LiteralPath $backupPath -Destination $executable -Force -ErrorAction SilentlyContinue
    }
    if ($stagingDirectory -and (Test-Path -LiteralPath $stagingDirectory)) {
        Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
    Show-UpdateMessage $message "Fantasy Trade Calculator Update Failed"
}
