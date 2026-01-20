<# 
    PowerShell version of the CIMS setup script.

    Behaviour mirrors the Bash script:
      - Creates/uses a virtual environment
      - Installs/updates dependencies from requirements.txt
      - Optionally launches Jupyter Lab with ./scenarios/Reference.ipynb

    Usage (rough equivalents to Bash):
      ./setup.ps1
      ./setup.ps1 -VenvName myenv
      ./setup.ps1 -NoJupyter
      ./setup.ps1 -NoUpdate
      ./setup.ps1 -Help
      ./setup.ps1 -Version
#>

[CmdletBinding()]
param(
    [Alias('n')]
    [string]$VenvName = "cims-env",

    # Equivalent to Bash --no-jupyter (default is on)
    [switch]$NoJupyter,

    # Equivalent to Bash --no-update (default is on)
    [switch]$NoUpdate,

    [switch]$Help,
    [switch]$Version
)

# --- Help / Version ---

if ($Help) {
    Write-Host "This script sets up a CIMS virtual environment, installs any required dependencies, and launches a modeling notebook in Jupyter Lab."
    Write-Host "Usage: setup.ps1 [-VenvName <name>] [-NoJupyter] [-NoUpdate] [-Help] [-Version]"
    Write-Host "  -VenvName   Specify the name of the virtual environment (default: 'cims-env')"
    Write-Host "  -NoJupyter  Do not launch Jupyter Lab (on by default)"
    Write-Host "  -NoUpdate   Do not update Python dependencies in existing virtual environment (on by default)"
    Write-Host "  -Help       Prints help"
    Write-Host "  -Version    Prints version"
    return
}

if ($Version) {
    Write-Host "setup.ps1 1.0"
    return
}

# --- Strict mode / error handling ---

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Configuration ---

$MinPythonMajor = 3
$MinPythonMinor = 9

# Default behaviour: jupyter ON, update ON
$LaunchJupyter = -not $NoJupyter.IsPresent
$UpdateDeps    = -not $NoUpdate.IsPresent

# --- Color helpers (rough analogue of print_color) ---

$NoColor = $false
function Write-Color {
    param(
        [string]$Message,
        [ValidateSet('Red','Green','Yellow','Cyan','None')]
        [string]$Color = 'None',
        [switch]$NoNewline
    )
    if ($NoColor) {
        if ($NoNewline) {
            Write-Host -NoNewline $Message
        } else {
            Write-Host $Message
        }
    } else {
        if ($Color -eq 'None') {
            if ($NoNewline) {
                Write-Host -NoNewline $Message
            } else {
                Write-Host $Message
            }
        } else {
            if ($NoNewline) {
                Write-Host -ForegroundColor $Color -NoNewline $Message
            } else {
                Write-Host -ForegroundColor $Color $Message
            }
        }
    }
}

# --- Python version check ---

function Test-PythonVersion {
    param(
        [string]$PythonPath
    )
    try {
        & $PythonPath -c "import sys; assert sys.version_info >= ($MinPythonMajor, $MinPythonMinor)" 2>$null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

# --- Discover Python interpreters (similar logic to Bash script) ---

function Get-PythonCandidates {
    $candidates = @()

    # 1. Prefer well-known commands if present (python3, python, py)
    foreach ($cmd in @('python3', 'python', 'py')) {
        try {
            $gc = Get-Command $cmd -ErrorAction SilentlyContinue
            if ($gc -and $gc.Source -and (Test-Path $gc.Source)) {
                $candidates += $gc.Source
            }
        } catch { }
    }

    # 2. Scan PATH for python3.* executables
    $pathDirs = $env:PATH -split ';'
    foreach ($dir in $pathDirs) {
        if (-not (Test-Path $dir)) { continue }
        try {
            $items = Get-ChildItem -Path $dir -Filter 'python3*' -File -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                $candidates += $item.FullName
            }
        } catch { }
    }

    # 3. Windows-style common Python locations (equivalent to /c/... in Bash)
    $winPaths = @(
        'C:\Python3*',
        'C:\Program Files\Python3*',
        (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python3*')
    )
    foreach ($pattern in $winPaths) {
        $dirs = Get-ChildItem -Path $pattern -Directory -ErrorAction SilentlyContinue
        foreach ($d in $dirs) {
            $exe = Join-Path $d.FullName 'python.exe'
            if (Test-Path $exe) {
                $candidates += $exe
            }
        }
    }

    # Deduplicate by full path
    $unique = $candidates | Sort-Object -Unique

    # Filter by minimum Python version
    $valid = @()
    foreach ($exe in $unique) {
        if (Test-PythonVersion -PythonPath $exe) {
            $valid += $exe
        }
    }

    if (-not $valid -or $valid.Count -eq 0) {
        return @()
    }

    # Sort by actual Python version, descending
    $withVersion = foreach ($exe in $valid) {
        try {
            $versionText = (& $exe --version 2>&1).Split()[1]
            [PSCustomObject]@{
                Path    = $exe
                Version = [version]$versionText
            }
        } catch {
            # If we can't parse version, give it a very low version
            [PSCustomObject]@{
                Path    = $exe
                Version = [version]'0.0.0'
            }
        }
    }

    $withVersion | Sort-Object Version -Descending
}

# --- Create virtual environment (includes interactive Python selection) ---

function New-CimsVirtualEnv {
    param(
        [string]$Name
    )

    $pyList = Get-PythonCandidates

    if (-not $pyList -or $pyList.Count -eq 0) {
        $min = "{0}.{1}" -f $MinPythonMajor, $MinPythonMinor
        Write-Color "No Python interpreter >= $min found." Red
        Write-Color "Please install Python >= $min and rerun this script." Red
        exit 1
    }

    # Build menu
    $menu = @()
    $index = 1
    foreach ($entry in $pyList) {
        if ($index -eq 1) {
            $menu += [PSCustomObject]@{
                Index   = $index
                Label   = "{0} (v{1}) - RECOMMENDED" -f $entry.Path, $entry.Version
                Path    = $entry.Path
                Version = $entry.Version
            }
        } else {
            $menu += [PSCustomObject]@{
                Index   = $index
                Label   = "{0} (v{1})" -f $entry.Path, $entry.Version
                Path    = $entry.Path
                Version = $entry.Version
            }
        }
        $index++
    }

    $cancelIndex = $menu.Count + 1

    # Print menu
    Write-Color ("Select an installed version of Python to use in your virtual environment:") Yellow
    foreach ($m in $menu) {
        Write-Host ("  [{0}] {1}" -f $m.Index, $m.Label)
    }
    Write-Host ("  [{0}] I'll install another Python version" -f $cancelIndex)

    # Read selection
    while ($true) {
        $choice = Read-Host "Enter selection number"
        if ([string]::IsNullOrWhiteSpace($choice)) {
            Write-Color "Invalid selection." Red
            continue
        }

        if (-not [int]::TryParse($choice, [ref]$null)) {
            Write-Color "Invalid selection." Red
            continue
        }

        $choiceInt = [int]$choice

        if ($choiceInt -eq $cancelIndex) {
            $min = "{0}.{1}" -f $MinPythonMajor, $MinPythonMinor
            Write-Color "Please install Python >= $min and rerun this script." Red
            exit 0
        }

        $selected = $menu | Where-Object { $_.Index -eq $choiceInt }
        if ($null -ne $selected) {
            Write-Color ("Selected {0}" -f $selected.Path) Green
            $selectedPython = $selected.Path

            Write-Color ("Building {0}..." -f $Name) None
            & $selectedPython -m venv $Name
            Write-Color "DONE" Green
            return
        }

        Write-Color "Invalid selection." Red
    }
}

# --- Step 1: Setup virtual environment ---

Write-Color ("Checking for {0} virtual environment..." -f $VenvName) None
if (Test-Path $VenvName) {
    $createNewEnv = $false
    Write-Color "FOUND" Green
} else {
    $createNewEnv = $true
    Write-Color "NOT FOUND" None
    Write-Color ("Creating {0} virtual environment..." -f $VenvName) Yellow
    New-CimsVirtualEnv -Name $VenvName
}

# --- Step 2: Activate virtual environment ---

Write-Color ("Activating {0} virtual environment..." -f $VenvName) None

# PowerShell-specific activation script
$activateScript = Join-Path $VenvName "Scripts\Activate.ps1"
if (-not (Test-Path $activateScript)) {
    # Fallback: cmd-style activate script (if user runs from a mixed environment)
    $activateScript = Join-Path $VenvName "Scripts\activate"
}

if (-not (Test-Path $activateScript)) {
    Write-Color ("Activation script not found in {0}\Scripts." -f $VenvName) Red
    exit 1
}

. $activateScript  # dot-source so it alters current session
Write-Color "DONE" Green

# --- Step 3: Install/update dependencies ---

$reqFile = "requirements.txt"
if (-not (Test-Path $reqFile)) {
    $hasReqs = $false
    Write-Color "requirements.txt not found in current directory; skipping dependency installation." Yellow
} else {
    $hasReqs = $true
}

if ($createNewEnv -and $hasReqs) {
    Write-Color "Installing dependencies..." Cyan
    & pip install -q --upgrade pip --disable-pip-version-check
    & pip install -q -r $reqFile --disable-pip-version-check
    Write-Color "DONE" Green
}
elseif (-not $createNewEnv -and $UpdateDeps -and $hasReqs) {
    Write-Color "Updating dependencies..." Cyan
    & pip install -q --upgrade pip --disable-pip-version-check
    & pip install -q -r $reqFile --disable-pip-version-check
    Write-Color "DONE" Green
}
elseif ($hasReqs) {
    Write-Color "Skipping dependency installation (use -NoUpdate:$false to force)." Yellow
}

# --- Step 4: Launch JupyterLab (if enabled) ---

if ($LaunchJupyter) {
    $jupyterPath = (Get-Command jupyter -ErrorAction SilentlyContinue)?.Source
    if (-not $jupyterPath) {
        Write-Color "Jupyter is not installed in this virtual environment." Red
        Write-Color "Add 'jupyterlab' to requirements.txt and re-run with -NoUpdate:$false to install it." Yellow
    } else {
        $notebook = "./scenarios/Reference.ipynb"
        if (Test-Path $notebook) {
            Write-Color "Launching JupyterLab...Use Ctrl+C to exit" None
            & jupyter lab --log-level=40 --notebook-dir=./ $notebook
        } else {
            Write-Color "Notebook $notebook not found. Launching JupyterLab in repository root instead." Yellow
            Write-Color "Launching JupyterLab...Use Ctrl+C to exit" None
            & jupyter lab --log-level=40 --notebook-dir=./
        }
        Write-Color "Closing Jupyter Lab..." None
        Write-Color "DONE" Green
    }
}

# --- Deactivate virtual environment ---

Write-Color ("Deactivating {0} virtual environment..." -f $VenvName) None
if (Get-Command deactivate -ErrorAction SilentlyContinue) {
    deactivate
}
Write-Color "DONE" Green

exit 0
