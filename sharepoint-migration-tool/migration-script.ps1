<#
.SYNOPSIS
  Migra pastas de bibliotecas do SharePoint Online (origem) para sites de destino, arquivo por arquivo.

.DESCRIPTION
  Le uma planilha CSV (NomePasta, SourceUrl, TargetUrl). Para cada linha: lista os arquivos da pasta de origem (recursivo),
  e para cada arquivo baixa para uma pasta temporaria, recria a estrutura de pastas no destino, envia e apaga o temporario.
  Cada arquivo gera uma linha no log (OK, ERRO com o motivo). Se o script for interrompido, rode de novo: o que ja deu OK no log
  e pulado. Antes de migrar de verdade, use -SimularApenas: lista o que seria migrado, sem baixar nem enviar nada.

.EXAMPLE
  .\migration-script.ps1 -ClientId 00000000-0000-0000-0000-000000000000 -SimularApenas
  .\migration-script.ps1 -ClientId 00000000-0000-0000-0000-000000000000 -BatchSize 5 -StartIndex 10
#>
[CmdletBinding()]
param(
    [string]$CsvPath = 'C:\Temp\Migracao\MATCH_FINAL_OK.csv',
    [string]$TempPath = 'C:\Temp\MigrationTemp',
    [string]$LogPath = 'C:\Temp\Migracao\log_migracao.csv',
    [string]$ClientId = $env:SPMIG_CLIENT_ID,
    [int]$BatchSize = 1,
    [int]$StartIndex = 0,
    [string[]]$Bibliotecas = @('Documentos Compartilhados', 'Shared Documents'),
    [int]$Tentativas = 3,
    [switch]$SimularApenas
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Migracao.psm1') -Force

if (-not $ClientId -or $ClientId -eq 'SEU_CLIENT_ID_AQUI') {
    throw 'Informe o ID do aplicativo do Entra ID com -ClientId (ou na variavel de ambiente SPMIG_CLIENT_ID).'
}
if ($BatchSize -lt 1 -or $StartIndex -lt 0) { throw 'BatchSize precisa ser 1 ou mais e StartIndex nao pode ser negativo.' }

$linhas = Import-PlanilhaMigracao -Caminho $CsvPath
$lote = @($linhas | Select-Object -Skip $StartIndex -First $BatchSize)
if ($lote.Count -eq 0) { Write-Host "Nada a migrar: a planilha tem $($linhas.Count) linha(s) e StartIndex e $StartIndex." -ForegroundColor Yellow; return }

foreach ($d in $TempPath, (Split-Path -Parent $LogPath)) { if ($d -and -not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d | Out-Null } }
$feitos = Get-ConcluidosDoLog -Caminho $LogPath
$total = @{ OK = 0; ERRO = 0; PULADO = 0 }

foreach ($linha in $lote) {
    Write-Host "`n=====================================" -ForegroundColor Cyan
    Write-Host "MIGRANDO: $($linha.NomePasta)" -ForegroundColor Yellow
    try {
        $origem = Split-UrlOrigem -Url $linha.SourceUrl -Bibliotecas $Bibliotecas
        $src = Connect-PnPOnline -Url $origem.SiteUrl -ClientId $ClientId -Interactive -ReturnConnection
        $itens = @(Get-PnPFolderItem -FolderSiteRelativeUrl $origem.PastaRelativa -ItemType File -Recursive -Connection $src)
        Write-Host "  $($itens.Count) arquivo(s) na origem."

        if ($SimularApenas) {
            foreach ($i in $itens | Select-Object -First 20) { Write-Host "  [simulacao] $(Get-CaminhoRelativo -ServerRelativeUrl $i.ServerRelativeUrl -Bibliotecas $Bibliotecas)" }
            if ($itens.Count -gt 20) { Write-Host "  [simulacao] ... e mais $($itens.Count - 20)" }
            continue
        }

        $dst = Connect-PnPOnline -Url $linha.TargetUrl -ClientId $ClientId -Interactive -ReturnConnection
        $biblioteca = Get-PnPList -Connection $dst | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden } | Select-Object -First 1
        if (-not $biblioteca) { throw 'O site de destino nao tem biblioteca de documentos visivel.' }
        $raiz = $biblioteca.RootFolder.ServerRelativeUrl
        $pastasOk = @{}

        foreach ($item in $itens) {
            $rel = Get-CaminhoRelativo -ServerRelativeUrl $item.ServerRelativeUrl -Bibliotecas $Bibliotecas
            if ($feitos.Contains("$($linha.NomePasta)|$rel")) { $total.PULADO++; continue }
            $tempNome = [guid]::NewGuid().ToString() + '_' + $item.Name
            $tempArq = Join-Path $TempPath $tempNome
            try {
                Invoke-ComRetry -Tentativas $Tentativas -Acao {
                    Get-PnPFile -Url $item.ServerRelativeUrl -Path $TempPath -FileName $tempNome -AsFile -Force -Connection $src
                }

                # recria a arvore de pastas no destino (cada pasta so e conferida uma vez)
                $pasta = Get-PastaDoCaminho -CaminhoRelativo $rel
                $destino = $raiz
                if ($pasta) {
                    $atual = $raiz
                    foreach ($parte in $pasta -split '/') {
                        $pai = $atual
                        $atual = "$atual/$parte"
                        if (-not $pastasOk.ContainsKey($atual)) {
                            if (-not (Get-PnPFolder -Url $atual -ErrorAction SilentlyContinue -Connection $dst)) {
                                Add-PnPFolder -Name $parte -Folder $pai -ErrorAction Stop -Connection $dst | Out-Null
                            }
                            $pastasOk[$atual] = $true
                        }
                    }
                    $destino = "$raiz/$pasta"
                }

                Invoke-ComRetry -Tentativas $Tentativas -Acao {
                    Add-PnPFile -Path $tempArq -Folder $destino -NewFileName $item.Name -Connection $dst | Out-Null
                }
                Write-LogMigracao -Caminho $LogPath -Pasta $linha.NomePasta -Arquivo $rel -Status OK
                $total.OK++
            }
            catch {
                Write-LogMigracao -Caminho $LogPath -Pasta $linha.NomePasta -Arquivo $rel -Status ERRO -Mensagem $_.Exception.Message
                Write-Host "  ERRO: $rel ($($_.Exception.Message))" -ForegroundColor Red
                $total.ERRO++
            }
            finally {
                if (Test-Path -LiteralPath $tempArq) { Remove-Item -LiteralPath $tempArq -Force }
            }
        }
    }
    catch {
        Write-LogMigracao -Caminho $LogPath -Pasta $linha.NomePasta -Arquivo '(pasta inteira)' -Status ERRO -Mensagem $_.Exception.Message
        Write-Host "Erro geral na pasta $($linha.NomePasta): $($_.Exception.Message)" -ForegroundColor Red
        $total.ERRO++
    }
}

Write-Host "`n=== FINALIZADO: $($total.OK) ok, $($total.ERRO) com erro, $($total.PULADO) ja migrados (pulados) ===" -ForegroundColor Green
if ($total.ERRO -gt 0) { Write-Host "Veja os erros em $LogPath (Status = ERRO) e rode de novo: so o que falhou e refeito." -ForegroundColor Yellow }
