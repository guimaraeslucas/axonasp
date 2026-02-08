# ================= CONFIGURAÇÕES =================
$url = 'http://localhost:4050/ebook.asp'
$SimultaneousUsers = 50  # Quantos "usuários" acessando ao mesmo tempo
$RestTimeMs = 1          # Tempo de descanso entre requisições de um mesmo usuário (ms)
# =================================================

# Cria um hash sincronizado (Thread-safe) para compartilhar dados entre as threads
$syncHash = [hashtable]::Synchronized(@{})
$syncHash.SuccessCount = 0
$syncHash.ErrorCount = 0
$syncHash.Running = $true
$syncHash.StartTime = Get-Date

# O bloco de código que cada "usuário" (thread) vai executar
$ScriptBlock = {
    param($url, $syncHash, $restTime)
    
    while ($syncHash.Running) {
        try {
            # Faz a requisição sem analisar o HTML (mais rápido)
            $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -Method Get -ErrorAction Stop
            
            # Incrementa contador de sucesso (Thread-safe)
            $syncHash.SuccessCount++
        }
        catch {
            # Incrementa contador de erro
            $syncHash.ErrorCount++
        }
        # Pequena pausa para respeitar a configuração (pode ser 0)
        Start-Sleep -Milliseconds $restTime
    }
}

# Configura o Pool de Runspaces (As Threads)
Write-Host "Iniciando $SimultaneousUsers threads para atacar $url..." -ForegroundColor Cyan
$RunspacePool = [runspacefactory]::CreateRunspacePool(1, $SimultaneousUsers)
$RunspacePool.Open()
$Jobs = @()

# Lança os "usuários"
for ($i = 0; $i -lt $SimultaneousUsers; $i++) {
    $PSInstance = [powershell]::Create()
    $PSInstance.RunspacePool = $RunspacePool
    $PSInstance.AddScript($ScriptBlock)
    $PSInstance.AddArgument($url)
    $PSInstance.AddArgument($syncHash)
    $PSInstance.AddArgument($RestTimeMs)
    
    # Inicia a thread de forma assíncrona
    $Jobs += $PSInstance.BeginInvoke()
}

# ================= LOOP DE MONITORAMENTO =================
try {
    Clear-Host
    while ($true) {
        $elapsed = (Get-Date) - $syncHash.StartTime
        $totalReq = $syncHash.SuccessCount + $syncHash.ErrorCount
        
        # Evita divisão por zero nos primeiros milissegundos
        if ($elapsed.TotalSeconds -gt 0) {
            $rps = [math]::Round($totalReq / $elapsed.TotalSeconds, 2)
        } else { $rps = 0 }

        # Atualiza o painel
        [Console]::SetCursorPosition(0,0)
        Write-Host "=== TESTE DE ESTRESSE EM ANDAMENTO ===" -ForegroundColor Cyan
        Write-Host "Alvo: $url"
        Write-Host "Threads (Usuários): $SimultaneousUsers"
        Write-Host "Tempo Decorrido:    $($elapsed.ToString('hh\:mm\:ss'))"
        Write-Host "--------------------------------------"
        Write-Host "Sucessos (200 OK):  $($syncHash.SuccessCount)" -ForegroundColor Green
        Write-Host "Falhas / Erros:     $($syncHash.ErrorCount)" -ForegroundColor Red
        Write-Host "Total Requisições:  $totalReq"
        Write-Host "--------------------------------------"
        Write-Host "Velocidade:         $rps Req/segundo" -ForegroundColor Yellow
        Write-Host "--------------------------------------"
        Write-Host "Pressione CTRL+C para parar..."
        
        Start-Sleep -Milliseconds 500
    }
}
finally {
    # Limpeza ao fechar
    Write-Host "`nParando threads..." -ForegroundColor Yellow
    $syncHash.Running = $false
    $RunspacePool.Close()
    $RunspacePool.Dispose()
    Write-Host "Teste finalizado."
}