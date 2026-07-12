Write-Host "Iniciando a preparação do documento..." -ForegroundColor Cyan

# 1. Cria o cabeçalho do Pandoc
$yamlHeader = @"
output-file: axonasp_manual.docx
toc: false
toc-depth: 3
input-files:
"@
$yamlHeader | Out-File -FilePath "pandoc.yaml" -Encoding utf8

$menuLines = Get-Content -Path "menu.md"
$headerDir = ".pandoc_headers"

# 2. Prepara uma pasta temporária para os cabeçalhos das sessões
if (Test-Path $headerDir) { Remove-Item -Path $headerDir -Recurse -Force }
New-Item -ItemType Directory -Path $headerDir | Out-Null

$headerCounter = 1

# 3. Lê o menu linha por linha e monta a estrutura
foreach ($line in $menuLines) {
    # Ignora linhas em branco
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    # Caso A: É um link para um arquivo de documentação
    if ($line -match '\[.*?\]\((.*?\.md)\)') {
        $filePath = $matches[1]
        
        # Proteção contra o "Erro 1": Só adiciona se o arquivo existir no disco
        if (Test-Path $filePath) {
            "  - $filePath" | Add-Content -Path "pandoc.yaml" -Encoding utf8
        }
        else {
            Write-Host "AVISO: Arquivo não encontrado e será ignorado: $filePath" -ForegroundColor Yellow
        }
    }
    # Caso B: É um item de lista sem link (Atua como um Título de Sessão)
    elseif ($line -match '^(\s*)\*\s+([^\[\]]+)$') {
        $indentSpaces = $matches[1].Length
        $titleText = $matches[2].Trim()
        
        # Calcula o nível do título baseado na indentação (0 espaços = H1, 4 espaços = H2...)
        $levelCount = [math]::Floor($indentSpaces / 4) + 1
        $headingPrefix = "#" * $levelCount
        
        # Cria um arquivo temporário apenas com esse título
        $tempFile = "$headerDir\header_$headerCounter.md"
        "$headingPrefix $titleText" | Out-File -FilePath $tempFile -Encoding utf8
        "  - $tempFile" | Add-Content -Path "pandoc.yaml" -Encoding utf8
        $headerCounter++
    }
    # Caso C: É um título Markdown puro (ex: # AxonASP Documentation)
    elseif ($line -match '^#+\s+(.*)') {
        $tempFile = "$headerDir\header_$headerCounter.md"
        $line | Out-File -FilePath $tempFile -Encoding utf8
        "  - $tempFile" | Add-Content -Path "pandoc.yaml" -Encoding utf8
        $headerCounter++
    }
}

Write-Host "Estrutura criada! Iniciando Pandoc..." -ForegroundColor Cyan

# 4. Executa o Pandoc e faz a limpeza
try {
    .\pandoc.exe -d pandoc.yaml
    Write-Host "Sucesso! A documentação foi gerada." -ForegroundColor Green
}
catch {
    Write-Host "Erro ao executar o Pandoc. Verifique as dependências." -ForegroundColor Red
}
finally {
    # Opcional: Apaga a pasta temporária depois de usar
    if (Test-Path $headerDir) { Remove-Item -Path $headerDir -Recurse -Force }
}