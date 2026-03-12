param(
    [switch]$Reload = $true,
    [string]$BindHost = "127.0.0.1",
    [int]$Port = 8000,
    [switch]$Loop
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Push-Location $repoRoot

$venvPath = Join-Path $repoRoot ".venv"
$pythonExe = Join-Path $venvPath "Scripts\python.exe"
$activatePs1 = Join-Path $venvPath "Scripts\Activate.ps1"

if (-not (Test-Path $pythonExe)) {
    Write-Host "Creating virtual environment in .venv..."
    python -m venv $venvPath
}

Write-Host "Activating virtual environment..."
. $activatePs1

Write-Host "Installing dependencies..."
python -m pip install --upgrade pip
pip install -r (Join-Path $repoRoot "requirements.txt")

# Optional: set Supabase env vars in your shell before running this script.
# Supported:
# - SUPABASE_URL / SUPABASE_ANON_KEY
# - VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY

$env:API_HOST = $BindHost
$env:API_PORT = "$Port"
if ($Reload) {
    $env:API_RELOAD = "true"
}

if ($env:API_HOST -match "^https?://") {
    $env:API_HOST = ($env:API_HOST -replace "^https?://", "")
}
if ($env:API_HOST -match ":\\d+$") {
    $env:API_HOST = ($env:API_HOST -replace ":\\d+$", "")
}

Write-Host "Starting API on http://$($env:API_HOST)`:$Port (docs at /docs)"

try {
    do {
        if ($Reload) {
            python -m uvicorn main:app --host $env:API_HOST --port $Port --reload
        } else {
            python -m uvicorn main:app --host $env:API_HOST --port $Port
        }

        if ($Loop) {
            Write-Host "Uvicorn exited. Restarting in 2 seconds..."
            Start-Sleep -Seconds 2
        }
    } while ($Loop)
} finally {
    Pop-Location
}
