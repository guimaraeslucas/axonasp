# build_caddy_linux.ps1
# Script to cross-compile Caddy with AxonASP module for Linux (amd64) from Windows

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "   AxonASP Caddy Linux (amd64) Builder       " -ForegroundColor White
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

$ParentDir = (Get-Item "..").FullName

$XcaddyCmd = Get-Command xcaddy -ErrorAction SilentlyContinue
$XcaddyExec = "xcaddy"

if (-not $XcaddyCmd) {
    $GoPath = & go env GOPATH 2>$null
    if ($GoPath) {
        $Candidate = Join-Path $GoPath "bin\xcaddy.exe"
        if (Test-Path $Candidate) {
            $XcaddyExec = $Candidate
            $XcaddyCmd = $true
        }
    }
}

if (-not $XcaddyCmd) {
    Write-Host "xcaddy not found. Downloading/installing xcaddy..." -ForegroundColor Yellow
    & go install github.com/caddyserver/xcaddy/cmd/xcaddy@latest
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to install xcaddy. Ensure Go is installed."
        exit 1
    }
    
    $GoPath = & go env GOPATH 2>$null
    if ($GoPath) {
        $Candidate = Join-Path $GoPath "bin\xcaddy.exe"
        if (Test-Path $Candidate) {
            $XcaddyExec = $Candidate
        }
    }
}

$env:GOOS = "linux"
$env:GOARCH = "amd64"

Write-Host "Building Caddy binary for Linux (amd64) -> caddy-linux-amd64..." -ForegroundColor Yellow
& $XcaddyExec build --output caddy-linux-amd64 --with g3pix.com.br/axonasp/caddy=. --replace "g3pix.com.br/axonasp=$ParentDir" --replace "github.com/google/cel-go=github.com/google/cel-go@v0.20.1"

Remove-Item Env:\GOOS -ErrorAction SilentlyContinue
Remove-Item Env:\GOARCH -ErrorAction SilentlyContinue

if ($LASTEXITCODE -eq 0) {
    Write-Host "[SUCCESS] caddy-linux-amd64 created successfully!" -ForegroundColor Green
} else {
    Write-Error "Compilation failed."
}
