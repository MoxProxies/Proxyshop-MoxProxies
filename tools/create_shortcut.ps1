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

# Locate the virtual environment Poetry created for this project
Push-Location $Repo
try {
    $Venv = (& poetry env info --path 2>$null | Out-String).Trim()
} catch {
    throw "Could not run Poetry. Install it first, see the setup guide in README.md."
} finally {
    Pop-Location
}
if (-not $Venv -or -not (Test-Path $Venv)) {
    throw "Poetry reported no virtual environment for this project. Run 'poetry install' first."
}

# pythonw.exe runs without a console, matching how the packaged app is built
$Interpreter = if ($Console) { 'python.exe' } else { 'pythonw.exe' }
$Target = Join-Path $Venv "Scripts\$Interpreter"
if (-not (Test-Path $Target)) {
    throw "Could not find $Interpreter at '$Target'. Try recreating the environment with 'poetry install'."
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
