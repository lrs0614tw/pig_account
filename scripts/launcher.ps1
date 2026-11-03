param([ValidateSet('install', 'run')][string]$Action = 'run')

$ErrorActionPreference = 'Stop'
$packageRoot = Split-Path -Parent $PSScriptRoot
$transcribing = $false
$resultCode = 1

function Invoke-Python($Candidate, $Arguments) {
    # Route native stdout/stderr through the host so the transcript records both.
    $ErrorActionPreference = 'Continue'
    & $Candidate.Exe @($Candidate.Args) @Arguments 2>&1 | ForEach-Object { Write-Host $_ }
    return $LASTEXITCODE
}

function Test-Python($Candidate, [switch]$Packages) {
    $snippet = 'import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)'
    if ($Packages) {
        $snippet = 'import sys, tkinter, openpyxl, docx, win32com.client; tkinter.Tcl(); sys.exit(0 if sys.version_info >= (3,10) else 1)'
    }
    try {
        & $Candidate.Exe @($Candidate.Args) -c $snippet 2>&1 | Out-Null
        return $LASTEXITCODE -eq 0
    } catch { return $false }
}

try {
    Set-Location -LiteralPath $packageRoot
    $logFolder = Join-Path $packageRoot 'logs'
    New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
    $logFile = Join-Path $logFolder ("{0}-{1}.log" -f $Action, (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    Start-Transcript -LiteralPath $logFile | Out-Null
    $transcribing = $true
    Write-Host "Log: $logFile"
    $venvPath = Join-Path $packageRoot '.venv\Scripts\python.exe'
    $venv = @{ Exe = $venvPath; Args = @() }
    $candidates = @()
    foreach ($name in @('py.exe', 'python.exe')) {
        $commands = @(Get-Command $name -CommandType Application -ErrorAction SilentlyContinue)
        foreach ($command in $commands) {
            $arguments = @()
            if ($name -eq 'py.exe') { $arguments = @('-3') }
            $candidates += @{ Exe = $command.Source; Args = $arguments }
        }
    }
    if ($Action -eq 'install') {
        if (-not (Test-Python $venv)) {
            $basePython = $candidates | Where-Object { Test-Python $_ } | Select-Object -First 1
            if (-not $basePython) { throw 'Python 3.10+ was not found. Install Python with Add to PATH enabled, then retry.' }
            Write-Host 'Creating the local Python environment...'
            $code = Invoke-Python $basePython @('-m', 'venv', (Join-Path $packageRoot '.venv'))
            if ($code -ne 0) { throw 'Could not create .venv. Check folder permissions and Python installation.' }
        }
        Write-Host 'Installing dependencies...'
        $code = Invoke-Python $venv @('-m', 'pip', 'install', '--disable-pip-version-check', '-r', (Join-Path $packageRoot 'requirements.txt'))
        if ($code -ne 0) { throw 'Dependency installation failed. Check the network and the error above.' }
        $code = Invoke-Python $venv @('-m', 'pip', 'check')
        if ($code -ne 0 -or -not (Test-Python $venv -Packages)) { throw 'The Python environment is incomplete.' }
        Write-Host 'Installation complete. You can now run the application BAT.'
    } else {
        $selected = $null
        if (Test-Python $venv -Packages) { $selected = $venv }
        if (-not $selected) {
            $selected = $candidates | Where-Object { Test-Python $_ -Packages } | Select-Object -First 1
        }
        if (-not $selected) { throw 'No working Python environment found. Run the installation BAT first.' }
        # Discover the entry point without depending on the console's Chinese code page.
        $entryPoints = @(Get-ChildItem -LiteralPath $packageRoot -Directory | ForEach-Object {
            $entry = Join-Path $_.FullName 'pig_accounting.py'
            if (Test-Path -LiteralPath $entry -PathType Leaf) { Get-Item -LiteralPath $entry }
        })
        if ($entryPoints.Count -ne 1) { throw 'Application files are missing or duplicated. Extract the entire downloaded folder first.' }
        Write-Host "Python: $($selected.Exe)"
        Set-Location -LiteralPath $entryPoints[0].DirectoryName
        $code = Invoke-Python $selected @($entryPoints[0].FullName)
        if ($code -ne 0) { throw "Application exited with code $code. See the error above." }
    }
    $resultCode = 0
} catch {
    Write-Host ''
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host 'Please send the log file to the maintainer if the problem continues.'
} finally {
    if ($transcribing) { Stop-Transcript | Out-Null }
}
exit $resultCode
