# Funcoes da migracao de SharePoint que nao dependem de conexao (por isso podem ser testadas sozinhas).
# O script migration-script.ps1 usa este modulo. Arquivo so ASCII, para o PowerShell 5.1 nao trocar os caracteres.
Set-StrictMode -Version 2.0

$script:BibliotecasPadrao = @('Documentos Compartilhados', 'Shared Documents')

# Separa a URL de uma pasta da origem em: URL do site, biblioteca e pasta relativa (biblioteca + subpastas).
# Ex.: https://x.sharepoint.com/sites/a/Documentos Compartilhados/RH/2024
#  ->  SiteUrl = https://x.sharepoint.com/sites/a ; PastaRelativa = Documentos Compartilhados/RH/2024
function Split-UrlOrigem {
    param([Parameter(Mandatory)] [string]$Url, [string[]]$Bibliotecas = $script:BibliotecasPadrao)
    $u = [uri]::UnescapeDataString($Url.Trim()).TrimEnd('/')
    foreach ($b in $Bibliotecas) {
        $marca = '/' + $b
        $i = $u.IndexOf($marca, [StringComparison]::OrdinalIgnoreCase)
        if ($i -lt 0) { continue }
        $resto = $u.Substring($i + $marca.Length)
        if ($resto -ne '' -and -not $resto.StartsWith('/')) { continue }   # nao confunde "Documentos Compartilhados2"
        return [pscustomobject]@{ SiteUrl = $u.Substring(0, $i); Biblioteca = $b; PastaRelativa = ($b + $resto) }
    }
    throw "A URL de origem nao tem uma biblioteca conhecida ($($Bibliotecas -join ', ')): $Url"
}

# Caminho de um arquivo dentro da biblioteca, a partir da URL relativa ao servidor.
# Ex.: /sites/a/Documentos Compartilhados/RH/2024/folha.xlsx -> RH/2024/folha.xlsx
function Get-CaminhoRelativo {
    param([Parameter(Mandatory)] [string]$ServerRelativeUrl, [string[]]$Bibliotecas = $script:BibliotecasPadrao)
    foreach ($b in $Bibliotecas) {
        $m = [regex]::Match($ServerRelativeUrl, '(?i)^.*?/' + [regex]::Escape($b) + '/(.+)$')
        if ($m.Success) { return $m.Groups[1].Value }
    }
    throw "Nao achei a biblioteca na URL do arquivo: $ServerRelativeUrl"
}

# Pasta de um caminho relativo ("RH/2024/folha.xlsx" -> "RH/2024"; sem pasta -> "").
function Get-PastaDoCaminho {
    param([Parameter(Mandatory)] [string]$CaminhoRelativo)
    $i = $CaminhoRelativo.LastIndexOf('/')
    if ($i -lt 0) { return '' }
    return $CaminhoRelativo.Substring(0, $i)
}

# Valida uma linha da planilha. Devolve a lista de problemas (vazia quando esta tudo certo).
function Test-LinhaMigracao {
    param([Parameter(Mandatory)] $Linha)
    $erros = New-Object System.Collections.Generic.List[string]
    $nomes = @($Linha.PSObject.Properties.Name)
    foreach ($c in 'NomePasta', 'SourceUrl', 'TargetUrl') {
        if ($nomes -notcontains $c) { $erros.Add("coluna ausente: $c"); continue }
        if (-not ([string]$Linha.$c).Trim()) { $erros.Add("valor vazio: $c") }
    }
    foreach ($c in 'SourceUrl', 'TargetUrl') {
        if ($nomes -contains $c -and ([string]$Linha.$c).Trim() -and ([string]$Linha.$c).Trim() -notmatch '^https://') { $erros.Add("$c precisa comecar com https://") }
    }
    return , @($erros)
}

# Le a planilha e confere todas as linhas antes de comecar (melhor falhar agora do que no meio da migracao).
function Import-PlanilhaMigracao {
    param([Parameter(Mandatory)] [string]$Caminho)
    if (-not (Test-Path -LiteralPath $Caminho)) { throw "Planilha nao encontrada: $Caminho" }
    $linhas = @(Import-Csv -LiteralPath $Caminho)
    if ($linhas.Count -eq 0) { throw "A planilha esta vazia: $Caminho" }
    $n = 1
    foreach ($l in $linhas) {
        $n++
        $problemas = Test-LinhaMigracao -Linha $l
        if ($problemas.Count -gt 0) { throw "Planilha, linha ${n}: $($problemas -join '; ')" }
    }
    return , $linhas
}

# Grava uma linha no log CSV (cria o arquivo com cabecalho). Os valores saem entre aspas, entao nomes com virgula nao quebram o CSV.
function Write-LogMigracao {
    param([Parameter(Mandatory)] [string]$Caminho, [Parameter(Mandatory)] [string]$Pasta, [Parameter(Mandatory)] [string]$Arquivo,
          [Parameter(Mandatory)] [ValidateSet('OK', 'ERRO', 'PULADO')] [string]$Status, [string]$Mensagem = '')
    $obj = [pscustomobject]@{ DataHora = (Get-Date).ToString('s'); Pasta = $Pasta; Arquivo = $Arquivo; Status = $Status; Mensagem = ($Mensagem -replace '\s+', ' ').Trim() }
    $csv = @($obj | ConvertTo-Csv -NoTypeInformation)
    if (Test-Path -LiteralPath $Caminho) { Add-Content -LiteralPath $Caminho -Value $csv[1] -Encoding UTF8 }
    else { Set-Content -LiteralPath $Caminho -Value $csv -Encoding UTF8 }
}

# Arquivos que ja foram migrados com sucesso, para continuar de onde parou. Chave: "Pasta|Arquivo".
# Entende o log novo (com cabecalho) e o formato antigo do script (sem cabecalho: pasta,caminho,OK).
function Get-ConcluidosDoLog {
    param([Parameter(Mandatory)] [string]$Caminho)
    $feitos = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    if (-not (Test-Path -LiteralPath $Caminho)) { return , $feitos }
    $linhas = @(Get-Content -LiteralPath $Caminho)
    if ($linhas.Count -eq 0) { return , $feitos }
    if ($linhas[0] -like '"DataHora"*') {
        foreach ($r in @(Import-Csv -LiteralPath $Caminho)) { if ($r.Status -eq 'OK') { [void]$feitos.Add("$($r.Pasta)|$($r.Arquivo)") } }
    } else {
        foreach ($l in $linhas) {
            $p = $l -split ','
            if ($p.Count -ge 3 -and $p[-1].Trim() -eq 'OK') { [void]$feitos.Add($p[0] + '|' + (($p[1..($p.Count - 2)]) -join ',')) }
        }
    }
    return , $feitos
}

# Executa uma acao com novas tentativas (rede instavel, limite do SharePoint). Espera 1x, 2x, 3x... o tempo base.
function Invoke-ComRetry {
    param([Parameter(Mandatory)] [scriptblock]$Acao, [int]$Tentativas = 3, [int]$EsperaSeg = 5, [scriptblock]$Dormir = { param($s) Start-Sleep -Seconds $s })
    for ($i = 1; $i -le $Tentativas; $i++) {
        try { return (& $Acao) }
        catch {
            if ($i -ge $Tentativas) { throw }
            & $Dormir ($EsperaSeg * $i)
        }
    }
}

Export-ModuleMember -Function Split-UrlOrigem, Get-CaminhoRelativo, Get-PastaDoCaminho, Test-LinhaMigracao, Import-PlanilhaMigracao, Write-LogMigracao, Get-ConcluidosDoLog, Invoke-ComRetry
