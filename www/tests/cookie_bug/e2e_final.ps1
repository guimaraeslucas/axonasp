$ErrorActionPreference = "Stop"
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
Write-Host "--- REQUEST 1 (lang=pt-BR) ---"
$res1 = Invoke-WebRequest -Uri "http://localhost:8801/g3pix/sitemap.asp?lang=pt-BR" -WebSession $session -UseBasicParsing
Write-Host "Set-Cookie Header: " $res1.Headers['Set-Cookie']
Write-Host "--- REQUEST 2 (no lang param) ---"
$res2 = Invoke-WebRequest -Uri "http://localhost:8801/g3pix/sitemap.asp" -WebSession $session -UseBasicParsing
Write-Host "Set-Cookie Header: " $res2.Headers['Set-Cookie']

$html = $res2.Content
if ($html -match "Mapa do site" -or $html -match "Página inicial") {
    Write-Host "SUCCESS: Language persisted as pt-BR!"
} else {
    Write-Host "FAILED: Language reverted!"
}