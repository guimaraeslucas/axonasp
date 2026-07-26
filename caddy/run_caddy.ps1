# run_caddy.ps1
# Script to build and start Caddy with AxonASP module on Windows

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "   AxonASP Caddy Server Launcher             " -ForegroundColor White
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

# Build Caddy with the AxonASP module if it doesn't exist
if (-not (Test-Path "caddy.exe")) {
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
            Write-Error "Failed to install xcaddy. Ensure Go is installed and configured correctly."
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

    $ParentDir = (Get-Item "..").FullName
    Write-Host "Building custom Caddy executable with AxonASP module via xcaddy..." -ForegroundColor Yellow
    & $XcaddyExec build --with g3pix.com.br/axonasp/caddy=. --replace "g3pix.com.br/axonasp=$ParentDir" --replace "github.com/google/cel-go=github.com/google/cel-go@v0.20.1"
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to compile custom Caddy server using xcaddy."
        exit 1
    }
    Write-Host "Caddy built successfully." -ForegroundColor Green
}

Write-Host "Starting Caddy on ports 8080 and 8081..." -ForegroundColor Cyan
$CaddyProcess = Start-Process -FilePath ".\caddy.exe" -ArgumentList "run", "--config", ".\Caddyfile" -NoNewWindow -PassThru -ErrorAction SilentlyContinue

if ($CaddyProcess) {
    Write-Host "Server running. Waiting to initialize..." -ForegroundColor Green
    Start-Sleep -Seconds 2
    
    Write-Host "Opening Site 1 (Port 8080)..." -ForegroundColor Yellow
    Start-Process "http://localhost:8080/default.asp"

    Write-Host "Opening Site 2 (Port 8081)..." -ForegroundColor Yellow
    Start-Process "http://localhost:8081/default.asp"

    Write-Host "Opening Site 3 (Port 8082)..." -ForegroundColor Yellow
    Start-Process "http://localhost:8082/default.asp"

    Write-Host ""
    Write-Host "Press any key to stop the server..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

    Write-Host "Stopping Caddy..." -ForegroundColor Cyan
    Stop-Process -Id $CaddyProcess.Id -Force
    Write-Host "Caddy stopped." -ForegroundColor Green
} else {
    Write-Error "Failed to start Caddy."
}
