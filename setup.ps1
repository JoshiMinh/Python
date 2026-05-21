$ErrorActionPreference = "Stop"

Write-Host "Setting up Virtual Environment..." -ForegroundColor Green
python -m venv .venv

# Set up environment variables to use the venv pip
$env:VIRTUAL_ENV = "$PWD\.venv"
$env:Path = "$PWD\.venv\Scripts;" + $env:Path

Write-Host "Upgrading pip..." -ForegroundColor Green
python -m pip install --upgrade pip

# Automatically find all subdirectories (excluding hidden folders like .venv)
$directories = Get-ChildItem -Directory | Where-Object { $_.Name -notmatch "^\." }

foreach ($dir in $directories) {
    $reqFile = Join-Path $dir.FullName "requirements.txt"
    if (Test-Path $reqFile) {
        Write-Host "Installing requirements for $($dir.Name)..." -ForegroundColor Green
        python -m pip install -r $reqFile
    }
}

Write-Host "Setup complete!" -ForegroundColor Green
