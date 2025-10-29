# PowerShell launcher for CIMS virtual environment and Jupyter Lab
# Usage: .\launch_cims.ps1 [-VenvName "cims-env"] [-NoJupyter] [-NoUpdate]

param(
    [string]$VenvName = "cims-env",
    [switch]$NoJupyter,
    [switch]$NoUpdate,
    [switch]$Help
)

# Configuration
$MIN_PYTHON_VERSION = @(3, 9)
$LaunchJupyter = -not $NoJupyter
$UpdateDeps = -not $NoUpdate

# Show help
if ($Help) {
    Write-Host @"
Usage: .\launch_cims.ps1 [-VenvName <name>] [-NoJupyter] [-NoUpdate]
  -VenvName      Name of virtual environment (default: cims-env)
  -NoJupyter     Skip launching Jupyter Lab
  -NoUpdate      Skip updating dependencies
  -Help          Show this help message

Examples:
  .\launch_cims.ps1
  .\launch_cims.ps1 -VenvName my-env
  .\launch_cims.ps1 -NoJupyter
  .\launch_cims.ps1 -VenvName my-env -NoUpdate
"@
    exit 0
}

# Color output functions
function Write-Msg { Write-Host $args -ForegroundColor Cyan }
function Write-Success { Write-Host $args -ForegroundColor Green }
function Write-Err { Write-Host $args -ForegroundColor Red }
function Write-Warn { Write-Host $args -ForegroundColor Yellow }

# Check Python version
function Test-PythonVersion {
    param([string]$PythonPath)
    
    try {
        $version = & $PythonPath -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>$null
        if ($version) {
            $major, $minor = $version.Split('.')
            return ([int]$major -gt $MIN_PYTHON_VERSION[0]) -or `
                   (([int]$major -eq $MIN_PYTHON_VERSION[0]) -and ([int]$minor -ge $MIN_PYTHON_VERSION[1]))
        }
    } catch {
        return $false
    }
    return $false
}

# Find all Python installations
function Find-PythonVersions {
    $pythonVersions = @()
    
    # Check common Python commands
    $pythonCommands = @("python", "python3", "py")
    
    foreach ($cmd in $pythonCommands) {
        try {
            $path = (Get-Command $cmd -ErrorAction SilentlyContinue).Source
            if ($path -and (Test-PythonVersion $path)) {
                $pythonVersions += $path
            }
        } catch {
            # Command not found, skip
        }
    }
    
    # Check py launcher with different versions
    if (Get-Command py -ErrorAction SilentlyContinue) {
        for ($minor = 9; $minor -le 15; $minor++) {
            try {
                $path = py -3.$minor -c "import sys; print(sys.executable)" 2>$null
                if ($path -and (Test-Path $path) -and (Test-PythonVersion $path)) {
                    $pythonVersions += $path
                }
            } catch {
                # Version not available
            }
        }
    }
    
    # Check common installation directories
    $commonPaths = @(
        "$env:LOCALAPPDATA\Programs\Python\Python3*",
        "C:\Python3*",
        "C:\Program Files\Python3*"
    )
    
    foreach ($pattern in $commonPaths) {
        $dirs = Get-ChildItem -Path $pattern -Directory -ErrorAction SilentlyContinue
        foreach ($dir in $dirs) {
            $pythonExe = Join-Path $dir.FullName "python.exe"
            if ((Test-Path $pythonExe) -and (Test-PythonVersion $pythonExe)) {
                $pythonVersions += $pythonExe
            }
        }
    }
    
    # Remove duplicates and return
    return $pythonVersions | Sort-Object -Unique
}

# Get Python version string
function Get-PythonVersionString {
    param([string]$PythonPath)
    
    try {
        $version = & $PythonPath --version 2>&1
        return $version -replace "Python ", ""
    } catch {
        return "Unknown"
    }
}

# Create virtual environment
function New-VirtualEnvironment {
    $pythonOptions = Find-PythonVersions
    
    if ($pythonOptions.Count -eq 0) {
        Write-Err "No Python >= $($MIN_PYTHON_VERSION[0]).$($MIN_PYTHON_VERSION[1]) found."
        Write-Err "Please install Python from https://www.python.org/downloads/"
        exit 1
    }
    
    # Create menu
    Write-Warn "`nSelect Python version for virtual environment:"
    
    $menu = @()
    for ($i = 0; $i -lt $pythonOptions.Count; $i++) {
        $version = Get-PythonVersionString $pythonOptions[$i]
        $path = $pythonOptions[$i]
        if ($i -eq 0) {
            $menu += "[$($i+1)] $path (v$version) - RECOMMENDED"
        } else {
            $menu += "[$($i+1)] $path (v$version)"
        }
    }
    $menu += "[0] Cancel - I'll install another version"
    
    foreach ($item in $menu) {
        Write-Host $item
    }
    
    # Get selection
    do {
        $selection = Read-Host "`nEnter selection number"
        $selectionNum = [int]$selection
    } while ($selectionNum -lt 0 -or $selectionNum -gt $pythonOptions.Count)
    
    if ($selectionNum -eq 0) {
        Write-Err "Installation cancelled."
        exit 0
    }
    
    $selectedPython = $pythonOptions[$selectionNum - 1]
    $selectedVersion = Get-PythonVersionString $selectedPython
    Write-Success "Selected Python $selectedVersion"
    
    # Create virtual environment
    Write-Msg "Creating virtual environment: $VenvName"
    & $selectedPython -m venv $VenvName
    
    if ($LASTEXITCODE -ne 0) {
        Write-Err "Failed to create virtual environment"
        exit 1
    }
    
    Write-Success "Virtual environment created"
    return $true
}

# Main script
Write-Host "`n=== CIMS Environment Launcher ===" -ForegroundColor Cyan
Write-Host ""

# Check if virtual environment exists
$installDeps = $false
if (Test-Path $VenvName) {
    Write-Msg "Virtual environment found: $VenvName"
    $installDeps = $UpdateDeps
} else {
    Write-Msg "Virtual environment not found"
    $installDeps = New-VirtualEnvironment
}

# Determine activation script path
$activateScript = Join-Path $VenvName "Scripts\Activate.ps1"

if (-not (Test-Path $activateScript)) {
    Write-Err "Virtual environment activation script not found: $activateScript"
    exit 1
}

# Activate virtual environment
Write-Msg "Activating virtual environment..."
try {
    & $activateScript
    Write-Success "Activated"
} catch {
    Write-Err "Failed to activate virtual environment: $_"
    exit 1
}

# Install/update dependencies
if ($installDeps) {
    if (Test-Path "requirements.txt") {
        Write-Msg "Installing dependencies..."
        python -m pip install --quiet --upgrade pip
        python -m pip install --quiet -r requirements.txt
        
        if ($LASTEXITCODE -eq 0) {
            Write-Success "Dependencies installed"
        } else {
            Write-Warn "Warning: Some dependencies may have failed to install"
        }
    } else {
        Write-Warn "requirements.txt not found, skipping dependency installation"
    }
}

# Launch Jupyter Lab
if ($LaunchJupyter) {
    # Check if Jupyter is installed
    $jupyterInstalled = Get-Command jupyter -ErrorAction SilentlyContinue
    
    if (-not $jupyterInstalled) {
        Write-Err "Jupyter not found. Please add 'jupyterlab' to requirements.txt"
        deactivate
        exit 1
    }
    
    # Check if notebook exists
    $notebookPath = ".\scenarios\Reference.ipynb"
    if (Test-Path $notebookPath) {
        Write-Msg "Launching Jupyter Lab... (Press Ctrl+C to exit)"
        jupyter lab --log-level=40 --notebook-dir=.\ $notebookPath
    } else {
        Write-Msg "Launching Jupyter Lab... (Press Ctrl+C to exit)"
        jupyter lab --log-level=40 --notebook-dir=.\
    }
}

# Deactivate
Write-Msg "Deactivating virtual environment..."
deactivate
Write-Success "Done"
Write-Host ""
