# Script PowerShell para converter o livro completo para DOCX usando Pandoc
# Uso: .\converter-para-docx.ps1

# Configurar encoding UTF-8 para suportar acentuação
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['*:Encoding'] = 'utf8'
chcp 65001 | Out-Null

# Definir o diretório de trabalho
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

# Verificar primeiro se existe pandoc.exe na pasta raiz
$pandocLocal = Join-Path $scriptDir "pandoc.exe"
$pandocExecutavel = $null

if (Test-Path $pandocLocal) {
    $pandocExecutavel = $pandocLocal
    Write-Host "Usando Pandoc local: $pandocLocal" -ForegroundColor Green
} else {
    # Fallback: verificar se o Pandoc está instalado no sistema
    $pandocPath = Get-Command pandoc -ErrorAction SilentlyContinue
    if (-not $pandocPath) {
        Write-Host "ERRO: Pandoc não encontrado." -ForegroundColor Red
        Write-Host "  - Procurando por: pandoc.exe na pasta raiz" -ForegroundColor Yellow
        Write-Host "  - Procurando por: Pandoc instalado no sistema" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Soluções:" -ForegroundColor Yellow
        Write-Host "  1. Coloque pandoc.exe na pasta raiz do projeto" -ForegroundColor Yellow
        Write-Host "  2. Ou instale o Pandoc: https://pandoc.org/installing.html" -ForegroundColor Yellow
        exit 1
    }
    $pandocExecutavel = $pandocPath.Source
    Write-Host "Usando Pandoc do sistema: $pandocExecutavel" -ForegroundColor Cyan
}

# Nome do arquivo de saída
$outputFile = "Compendio_Futhark_Antigo.docx"

# Lista de arquivos na ordem correta (pasta livro)
#$arquivos = @(
#    "compendio\sumario.md",
#    "compendio\prefacio.md",
#    "compendio\prologo.md",
#    "compendio\capitulo-01.md",
#    "compendio\capitulo-02.md",
#    "compendio\capitulo-03.md",
#    "compendio\capitulo-04.md",
#    "compendio\capitulo-05.md",
#    "compendio\capitulo-06.md",
#    "compendio\capitulo-07.md",
#    "compendio\capitulo-08.md",
#    "compendio\capitulo-09.md",
#
#    "compendio\bibliografia.md",
#
#    "compendio\apendice-a.md",
#    "compendio\apendice-b.md",
#    "compendio\apendice-c.md",
#    "compendio\apendice-d.md",
#    "compendio\apendice-e.md",
#    "compendio\apendice-f.md",
#
#    "compendio\indice-remissivo.md"
#)

#Teste de arquivos
$arquivos = @(
    "compendio\sumario.md",
    "compendio\prefacio.md",
    "compendio\prologo.md",
    "compendio\capitulo-01.md",
    "compendio\capitulo-02.md"
)

# Verificar se todos os arquivos existem
$arquivosFaltando = @()
foreach ($arquivo in $arquivos) {
    if (-not (Test-Path $arquivo)) {
        $arquivosFaltando += $arquivo
    }
}

if ($arquivosFaltando.Count -gt 0) {
    Write-Host "AVISO: Os seguintes arquivos não foram encontrados:" -ForegroundColor Yellow
    foreach ($arquivo in $arquivosFaltando) {
        Write-Host "  - $arquivo" -ForegroundColor Yellow
    }
    Write-Host ""
    $continuar = Read-Host "Deseja continuar mesmo assim? (S/N)"
    if ($continuar -ne "S" -and $continuar -ne "s") {
        exit 1
    }
}

# Construir argumentos do Pandoc
$arquivosExistentes = $arquivos | Where-Object { Test-Path $_ }
$argumentosPandoc = @()

# Perguntar ao usuário se deseja usar o template custom-reference.docx ANTES de construir argumentos
$usarTemplate = $true
#$usarTemplate = $false
#if (Test-Path "custom-reference.docx") {
#    Write-Host ""
#    Write-Host "Arquivo custom-reference.docx encontrado!" -ForegroundColor Green
#    $resposta = Read-Host "Deseja usar o template custom-reference.docx para formatação? (S/N - padrão: S)"
#    if ($resposta -eq "" -or $resposta -eq "S" -or $resposta -eq "s") {
#        $usarTemplate = $true
#        Write-Host "Usando arquivo de referência: custom-reference.docx" -ForegroundColor Cyan
#        Write-Host ""
#        Write-Host "NOTA: Certifique-se de que os estilos no template estao nomeados corretamente:" -ForegroundColor Yellow
#        Write-Host "  - Title (ou Titulo) para # titulos" -ForegroundColor Gray
#        Write-Host "  - Heading 1 (ou Titulo 1) para ## titulos" -ForegroundColor Gray
#        Write-Host "  - Heading 2 (ou Titulo 2) para ### titulos" -ForegroundColor Gray
#        Write-Host "  - Heading 3 (ou Titulo 3) para #### titulos" -ForegroundColor Gray
#        Write-Host "  - Heading 4 (ou Titulo 4) para ##### titulos" -ForegroundColor Gray
#        Write-Host "  - Subtitle (ou Subtitulo) para subtitulos" -ForegroundColor Gray
#        Write-Host "  - Normal para paragrafos" -ForegroundColor Gray
#        Write-Host ""
#    } else {
#        Write-Host "Usando formatação padrão do Pandoc (sem template)." -ForegroundColor Yellow
#    }
#} else {
#    Write-Host ""
#    Write-Host "AVISO: Arquivo custom-reference.docx não encontrado!" -ForegroundColor Yellow
#    Write-Host "  As fontes personalizadas (Acadian Runes, Caveat) NÃO serão aplicadas." -ForegroundColor Yellow
#    Write-Host "  Para usar as fontes configuradas, crie o template seguindo:" -ForegroundColor Yellow
#    Write-Host "  GUIA_TEMPLATE_WORD.md" -ForegroundColor Cyan
#    Write-Host ""
#    Write-Host "Usando formatação padrão do Pandoc (sem template)." -ForegroundColor Yellow
#}

# Verificar se existe arquivo de defaults
if (Test-Path "pandoc-docx.yaml") {
    $argumentosPandoc += "--defaults=pandoc-docx.yaml"
    Write-Host "Usando configurações de: pandoc-docx.yaml" -ForegroundColor Cyan
} else {
    # Fallback: usar metadata.yaml como variáveis
    $argumentosPandoc += "--metadata-file=metadata.yaml"
    $argumentosPandoc += "--toc"
    $argumentosPandoc += "--toc-depth=3"
    $argumentosPandoc += " --no-number-sections"
    $argumentosPandoc += "--standalone"
    Write-Host "Usando configurações de: metadata.yaml (modo fallback)" -ForegroundColor Yellow
}

# Adicionar arquivos de entrada
$argumentosPandoc += $arquivosExistentes

# Adicionar arquivo de saída
$argumentosPandoc += "-o"
$argumentosPandoc += $outputFile

# IMPORTANTE: --reference-doc deve vir DEPOIS de --defaults e DEPOIS de -o
# Mas antes dos arquivos de entrada não funciona bem, então vamos adicionar no final
if ($usarTemplate) {
    # Adicionar reference-doc no final para garantir que seja aplicado
    $argumentosPandoc += "--reference-doc=custom-reference.docx"
}

# Construir string do comando para exibição
$comandoExibicao = "`"$pandocExecutavel`" " + ($argumentosPandoc -join " ")

Write-Host ""
Write-Host "Executando conversão..." -ForegroundColor Cyan
Write-Host "Comando: $comandoExibicao" -ForegroundColor Gray
Write-Host ""

# Executar o comando usando operador de chamada &
& $pandocExecutavel $argumentosPandoc

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "SUCESSO! Arquivo gerado: $outputFile" -ForegroundColor Green
    Write-Host "Localização: $(Resolve-Path $outputFile)" -ForegroundColor Green
    if ($usarTemplate) {
        Write-Host ""
        Write-Host "NOTA: O documento foi gerado usando o template custom-reference.docx." -ForegroundColor Cyan
        Write-Host "      Você pode abrir o arquivo no Word e ajustar estilos conforme necessário." -ForegroundColor Cyan
    }
} else {
    Write-Host ""
    Write-Host "ERRO na conversão. Código de saída: $LASTEXITCODE" -ForegroundColor Red
    Write-Host ""
    Write-Host "Dicas para resolver problemas:" -ForegroundColor Yellow
    Write-Host "  - Verifique se todos os arquivos Markdown existem na pasta livro\" -ForegroundColor Yellow
    Write-Host "  - Verifique a versão do Pandoc: & `"$pandocExecutavel`" --version" -ForegroundColor Yellow
    Write-Host "  - Tente executar o comando manualmente para ver erros detalhados" -ForegroundColor Yellow
    exit 1
}

