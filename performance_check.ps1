# ================= CONFIGURATION =================
$url = 'http://localhost:4050/ebook.asp'
$SimultaneousUsers = 50  # How many "users" accessing at the same time
$RestTimeMs = 1          # Rest time between requests from the same user (ms)
# =================================================

# Creates a synchronized hash (Thread-safe) to share data between threads
$syncHash = [hashtable]::Synchronized(@{})
$syncHash.SuccessCount = 0
$syncHash.ErrorCount = 0
$syncHash.Running = $true
$syncHash.StartTime = Get-Date

# The code block that each "user" (thread) will execute
$ScriptBlock = {
    param($url, $syncHash, $restTime)
    
    while ($syncHash.Running) {
        try {
            # Make the request without parsing HTML (faster)
            $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -Method Get -ErrorAction Stop
            
            # Increment success counter (Thread-safe)
            $syncHash.SuccessCount++
        }
        catch {
            # Increment error counter
            $syncHash.ErrorCount++
        }
        # Small pause to respect the configuration (can be 0)
        Start-Sleep -Milliseconds $restTime
    }
}

# Configure the Runspace Pool (The Threads)
Write-Host "Starting $SimultaneousUsers threads to stress test $url..." -ForegroundColor Cyan
$RunspacePool = [runspacefactory]::CreateRunspacePool(1, $SimultaneousUsers)
$RunspacePool.Open()
$Jobs = @()

# Launch the "users"
for ($i = 0; $i -lt $SimultaneousUsers; $i++) {
    $PSInstance = [powershell]::Create()
    $PSInstance.RunspacePool = $RunspacePool
    $PSInstance.AddScript($ScriptBlock)
    $PSInstance.AddArgument($url)
    $PSInstance.AddArgument($syncHash)
    $PSInstance.AddArgument($RestTimeMs)
    
    # Start the thread asynchronously
    $Jobs += $PSInstance.BeginInvoke()
}

# ================= MONITORING LOOP =================
try {
    Clear-Host
    while ($true) {
        $elapsed = (Get-Date) - $syncHash.StartTime
        $totalReq = $syncHash.SuccessCount + $syncHash.ErrorCount
        
        # Avoid division by zero in the first milliseconds
        if ($elapsed.TotalSeconds -gt 0) {
            $rps = [math]::Round($totalReq / $elapsed.TotalSeconds, 2)
        } else { $rps = 0 }

        # Update the dashboard
        [Console]::SetCursorPosition(0,0)
        Write-Host "=== STRESS TEST IN PROGRESS ===" -ForegroundColor Cyan
        Write-Host "Target: $url"
        Write-Host "Threads (Users): $SimultaneousUsers"
        Write-Host "Elapsed Time:    $($elapsed.ToString('hh\:mm\:ss'))"
        Write-Host "--------------------------------------"
        Write-Host "Successes (200 OK):  $($syncHash.SuccessCount)" -ForegroundColor Green
        Write-Host "Failures / Errors:   $($syncHash.ErrorCount)" -ForegroundColor Red
        Write-Host "Total Requests:      $totalReq"
        Write-Host "--------------------------------------"
        Write-Host "Speed:            $rps Req/second" -ForegroundColor Yellow
        Write-Host "--------------------------------------"
        Write-Host "Press CTRL+C to stop..."
        
        Start-Sleep -Milliseconds 500
    }
}
finally {
    # Cleanup on close
    Write-Host "`nStopping threads..." -ForegroundColor Yellow
    $syncHash.Running = $false
    $RunspacePool.Close()
    $RunspacePool.Dispose()
    Write-Host "Test finished."
}