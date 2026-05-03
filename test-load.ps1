# Força o carregamento da biblioteca HTTP no PowerShell
Add-Type -AssemblyName System.Net.Http

$url = "http://localhost:8801/tests/?file=test_g3zstd.asp"
#$url = "http://localhost:8801/manual/"
$totalRequests = 2000 # Ajuste o número de requisições simultâneas aqui

Write-Host "Iniciando disparo de $totalRequests requisições simultâneas para $url..." -ForegroundColor Cyan

# Instancia o HttpClient
$client = New-Object System.Net.Http.HttpClient

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Dispara todas as requisições e salva as Tasks nativamente em um array do PowerShell
$tasks = foreach ($i in 1..$totalRequests) {
    $client.GetAsync($url)
}

# Trava a execução até que TODAS as requisições tenham retornado
try {
    [System.Threading.Tasks.Task]::WaitAll($tasks)
}
catch {
    Write-Host "Algumas requisições falharam por timeout ou recusa de conexão." -ForegroundColor Yellow
}

$stopwatch.Stop()

# Contadores para os resultados
$successCount = 0
$errorCount = 0

# Avalia o status HTTP de cada tarefa
foreach ($task in $tasks) {
    if ($task.Status -eq 'RanToCompletion') {
        if ($task.Result.IsSuccessStatusCode) {
            $successCount++
        }
        else {
            $errorCount++
        }
        $task.Result.Dispose()
    }
    else {
        $errorCount++
    }
}

$client.Dispose()

Write-Host "`n--- Resultado do Teste ---" -ForegroundColor White
Write-Host "Tempo total: $($stopwatch.ElapsedMilliseconds) ms" -ForegroundColor White
Write-Host "Requisições com Sucesso (HTTP 200): $successCount" -ForegroundColor Green

if ($errorCount -gt 0) {
    Write-Host "Requisições com Erro/Falha: $errorCount" -ForegroundColor Red
}
else {
    Write-Host "Requisições com Erro/Falha: $errorCount" -ForegroundColor Green
}