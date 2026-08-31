<#
.SYNOPSIS
    Creates a shortcut which launches Proxyshop from this repository.

.DESCRIPTION
    Points the shortcut at the virtual environment's pythonw.exe rather than wrapping
    `poetry run`, which avoids a console window on every launch, skips Poetry's startup
    cost, and gives Windows a real executable to pin to the taskbar.

    The shortcut's working directory is set to the repository root. This is required:
    paths for `templates`, `art`, `out`, and `fonts` are all resolved relative to the
    working directory, so launching from anywhere else would read and write the wrong
    folders.

.PARAMETER Name
    Name of the shortcut. Defaults to 'Proxyshop'.

.PARAMETER PythonPath
    Full path to the environment's python.exe, skipping automatic detection. Use this
    if Poetry cannot be found or its output cannot be interpreted. Get the value with:
    poetry run python -c "import sys; print(sys.executable)"

.PARAMETER Console
    Use python.exe instead of pythonw.exe, so the app runs with a console window
    attached. Use this when the app fails to start and you need to see the error.

.PARAMETER Desktop
    Also place a copy of the shortcut on the desktop.

.EXAMPLE
    .\tools\create_shortcut.ps1
    Creates a Start Menu shortcut. Search for it in Start, then right-click it and
    choose 'Pin to taskbar'.

.EXAMPLE
    .\tools\create_shortcut.ps1 -Console -Name 'Proxyshop (Debug)'
    Creates a second shortcut which keeps a console window open for troubleshooting.
#>
[CmdletBinding()]
param(
    [string] $Name = 'Proxyshop',
    [string] $PythonPath,
    [switch] $Console,
    [switch] $Desktop
)

$ErrorActionPreference = 'Stop'

# Resolve the repository root from this script's location, so the script works
# no matter which directory it is invoked from
$Repo = Split-Path -Parent $PSScriptRoot
$Entry = Join-Path $Repo 'main.py'
if (-not (Test-Path $Entry)) {
    throw "Could not find main.py at '$Entry'. Is this script still in the repository's tools folder?"
}

# Ask the environment's own interpreter where it lives. Reading the path out of
# `poetry env info` is unreliable, because Poetry mixes advisory lines into its output,
# such as when the active Python is unsupported and it selects a different one.
if ($PythonPath) {
    if (-not [System.IO.File]::Exists($PythonPath)) {
        throw "No interpreter found at '$PythonPath'."
    }
    $Output = @($PythonPath)
} else {
    $Output = $null
}
Push-Location $Repo
try {
    # A native command writing to stderr must not abort the script here
    if (-not $Output) {
        $Previous, $ErrorActionPreference = $ErrorActionPreference, 'Continue'
        $Output = & poetry run python -c "import sys; print(sys.executable)" 2>&1
        $ErrorActionPreference = $Previous
    }
} catch {
    throw "Could not run Poetry. Install it first, see the setup guide in README.md."
} finally {
    Pop-Location
}

# Advisory lines are discarded by keeping only output which is a real file.
# `File.Exists` is used rather than `Test-Path`, which throws on a value containing
# characters illegal in a path. Poetry's advisory text contains them, for example the
# '<' and '>' in a version constraint such as (>=3.10,<3.13).
$Python = $Output |
    ForEach-Object { $_.ToString().Trim() } |
    Where-Object { $_.EndsWith('.exe') -and [System.IO.File]::Exists($_) } |
    Select-Object -Last 1

if (-not $Python) {
    throw ("Could not determine the environment's Python interpreter.`n" +
           "Run 'poetry install' first, or pass the interpreter directly with -PythonPath.`n" +
           "Poetry returned:`n" + ($Output | Out-String))
}

# pythonw.exe runs without a console, matching how the packaged app is built.
# It sits beside python.exe, so the interpreter's own folder is used rather than
# assuming the environment's layout.
$Interpreter = if ($Console) { 'python.exe' } else { 'pythonw.exe' }
$Target = Join-Path (Split-Path -Parent $Python) $Interpreter
if (-not (Test-Path $Target)) {
    throw "Could not find $Interpreter beside '$Python'. Try recreating the environment with 'poetry install'."
}

# Build the shortcut
$Icon = Join-Path $Repo 'src\img\favicon.ico'
$Shell = New-Object -ComObject WScript.Shell
$Paths = @(Join-Path $env:AppData "Microsoft\Windows\Start Menu\Programs\$Name.lnk")
if ($Desktop) {
    $Paths += Join-Path ([Environment]::GetFolderPath('Desktop')) "$Name.lnk"
}

foreach ($Path in $Paths) {
    $Link = $Shell.CreateShortcut($Path)
    $Link.TargetPath = $Target
    $Link.Arguments = 'main.py'

    # Required, the app resolves its data folders relative to the working directory
    $Link.WorkingDirectory = $Repo
    $Link.Description = "Proxyshop, launched from $Repo"
    if (Test-Path $Icon) {
        $Link.IconLocation = $Icon
    }
    $Link.Save()
    Write-Host "Created $Path"
}

Write-Host ''
Write-Host "Open Start, search for '$Name', then right-click it and choose 'Pin to taskbar'."
if (-not $Console) {
    Write-Host "If nothing happens when it launches, rerun this script with -Console to see the error."
}
