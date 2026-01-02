# Define o diretório atual como base
$currentDir = $PSScriptRoot
$pandocPath = Join-Path $currentDir "pandoc.exe"

# 1. Pergunta o nome do arquivo ao usuário
$fileName = Read-Host "Digite o nome do arquivo Markdown (ou pressione Enter para usar 'teste-word-conver.md')"

# 2. Define o padrão se a entrada for vazia
if ([string]::IsNullOrWhiteSpace($fileName)) {
    $fileName = "teste-word-conver.md"
}

# 3. Garante que o arquivo tenha a extensão .md para a busca
if (-not $fileName.EndsWith(".md")) {
    $fileName = $fileName + ".md"
}

$inputFile = Join-Path $currentDir $fileName

# 4. Verifica se o arquivo de entrada existe
if (-not (Test-Path $inputFile)) {
    Write-Host "Erro: O arquivo '$fileName' não foi encontrado na pasta." -ForegroundColor Red
    Pause
    exit
}

# 5. Define o nome do arquivo de saída (trocando .md por .docx)
$outputFile = $inputFile.Replace(".md", ".docx")

# 6. Executa o Pandoc
Write-Host "Convertendo '$fileName' para Word..." -ForegroundColor Cyan

try {
    & $pandocPath $inputFile -o $outputFile --reference-doc="custom-reference.docx"
    Write-Host "Sucesso! Arquivo gerado: $(Split-Path $outputFile -Leaf)" -ForegroundColor Green
}
catch {
    Write-Host "Ocorreu um erro durante a conversão." -ForegroundColor Red
    $PSItem.Exception.Message
}

Pause