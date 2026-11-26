# Testes dos scripts de Microsoft 365. Rodar:  Invoke-Pester .\Testes
# Nao precisam de conta do Microsoft 365: as funcoes puras sao testadas direto e o script de migracao roda com o PnP simulado (mocks).
$raiz = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $raiz 'sharepoint-migration-tool\Migracao.psm1') -Force
Import-Module (Join-Path $raiz 'user-provisioning\Provisionamento.psm1') -Force

Describe 'Migracao: URLs e caminhos' {
    It 'separa o site e a pasta relativa' {
        $r = Split-UrlOrigem 'https://contoso.sharepoint.com/sites/matriz/Documentos Compartilhados/RH/2024'
        $r.SiteUrl | Should Be 'https://contoso.sharepoint.com/sites/matriz'
        $r.PastaRelativa | Should Be 'Documentos Compartilhados/RH/2024'
    }
    It 'aceita a biblioteca em ingles, URL codificada e barra no fim' {
        $r = Split-UrlOrigem 'https://contoso.sharepoint.com/sites/a/Shared%20Documents/Fin/'
        $r.SiteUrl | Should Be 'https://contoso.sharepoint.com/sites/a'
        $r.PastaRelativa | Should Be 'Shared Documents/Fin'
    }
    It 'pasta na raiz da biblioteca' {
        (Split-UrlOrigem 'https://x.sharepoint.com/sites/a/Documentos Compartilhados').PastaRelativa | Should Be 'Documentos Compartilhados'
    }
    It 'nao confunde nome parecido e recusa URL sem biblioteca' {
        { Split-UrlOrigem 'https://x.sharepoint.com/sites/a/Documentos Compartilhados2/RH' } | Should Throw 'biblioteca conhecida'
        { Split-UrlOrigem 'https://x.sharepoint.com/sites/a/Lists/Tarefas' } | Should Throw
    }
    It 'caminho do arquivo dentro da biblioteca' {
        Get-CaminhoRelativo '/sites/a/Documentos Compartilhados/RH/2024/folha.xlsx' | Should Be 'RH/2024/folha.xlsx'
        Get-CaminhoRelativo '/sites/a/Shared Documents/x.txt' | Should Be 'x.txt'
        { Get-CaminhoRelativo '/sites/a/Outra/x.txt' } | Should Throw
    }
    It 'pasta de um caminho relativo' {
        Get-PastaDoCaminho 'RH/2024/folha.xlsx' | Should Be 'RH/2024'
        Get-PastaDoCaminho 'folha.xlsx' | Should Be ''
    }
}

Describe 'Migracao: planilha' {
    function Nova-Planilha($texto) { $p = Join-Path $TestDrive ([guid]::NewGuid().ToString('N') + '.csv'); Set-Content -Path $p -Value $texto -Encoding UTF8; return $p }
    It 'aceita o exemplo do repositorio' {
        $l = Import-PlanilhaMigracao (Join-Path $raiz 'sharepoint-migration-tool\exemplo-match.csv')
        @($l).Count | Should Be 2
        $l[0].NomePasta | Should Be 'Financeiro'
    }
    It 'recusa planilha ausente, vazia, sem coluna ou com URL ruim, dizendo a linha' {
        { Import-PlanilhaMigracao (Join-Path $TestDrive 'nao-existe.csv') } | Should Throw 'nao encontrada'
        { Import-PlanilhaMigracao (Nova-Planilha 'NomePasta,SourceUrl,TargetUrl') } | Should Throw 'vazia'
        { Import-PlanilhaMigracao (Nova-Planilha "NomePasta,SourceUrl`nA,https://x/y") } | Should Throw 'coluna ausente: TargetUrl'
        { Import-PlanilhaMigracao (Nova-Planilha "NomePasta,SourceUrl,TargetUrl`nA,https://x/y,https://z`nB,http://x,https://z") } | Should Throw 'linha 3'
        { Import-PlanilhaMigracao (Nova-Planilha "NomePasta,SourceUrl,TargetUrl`n,https://x/y,https://z") } | Should Throw 'valor vazio: NomePasta'
    }
}

Describe 'Migracao: log e retomada' {
    It 'cria o log com cabecalho e protege nomes com virgula e aspas' {
        $log = Join-Path $TestDrive 'log1.csv'
        Write-LogMigracao -Caminho $log -Pasta 'RH' -Arquivo 'a, b "c".xlsx' -Status OK
        Write-LogMigracao -Caminho $log -Pasta 'RH' -Arquivo 'd.txt' -Status ERRO -Mensagem "falhou`r`nde novo"
        $l = @(Import-Csv $log)
        $l.Count | Should Be 2
        $l[0].Arquivo | Should Be 'a, b "c".xlsx'
        $l[1].Status | Should Be 'ERRO'
        $l[1].Mensagem | Should Be 'falhou de novo'
        (Get-Content $log)[0] | Should Be '"DataHora","Pasta","Arquivo","Status","Mensagem"'
    }
    It 'so conta como concluido o que deu OK' {
        $log = Join-Path $TestDrive 'log2.csv'
        Write-LogMigracao -Caminho $log -Pasta 'RH' -Arquivo 'a, b.xlsx' -Status OK
        Write-LogMigracao -Caminho $log -Pasta 'RH' -Arquivo 'x.txt' -Status ERRO -Mensagem 'x'
        $f = Get-ConcluidosDoLog $log
        $f.Contains('RH|a, b.xlsx') | Should Be $true
        $f.Contains('rh|A, B.XLSX') | Should Be $true     # sem diferenciar maiusculas
        $f.Contains('RH|x.txt') | Should Be $false
    }
    It 'entende o log antigo, sem cabecalho, e arquivo ausente' {
        $log = Join-Path $TestDrive 'antigo.csv'
        Set-Content $log "RH,2024/folha.xlsx,OK`nRH,2024/outro, com virgula.txt,OK`nRH,erro.txt,ERRO"
        $f = Get-ConcluidosDoLog $log
        $f.Contains('RH|2024/folha.xlsx') | Should Be $true
        $f.Contains('RH|2024/outro, com virgula.txt') | Should Be $true
        $f.Contains('RH|erro.txt') | Should Be $false
        (Get-ConcluidosDoLog (Join-Path $TestDrive 'nao-existe.csv')).Count | Should Be 0
    }
}

Describe 'Migracao: novas tentativas' {
    It 'tenta de novo e devolve o resultado quando funciona' {
        $global:t = 0
        $r = Invoke-ComRetry -Tentativas 3 -EsperaSeg 2 -Dormir { param($s) $global:esperas += $s } -Acao { $global:t++; if ($global:t -lt 3) { throw 'rede' }; 'ok' }
        $r | Should Be 'ok'
        $global:t | Should Be 3
    }
    It 'depois da ultima tentativa devolve o erro original, com esperas crescentes' {
        $global:t = 0; $global:esperas = @()
        { Invoke-ComRetry -Tentativas 3 -EsperaSeg 2 -Dormir { param($s) $global:esperas += $s } -Acao { $global:t++; throw 'sempre falha' } } | Should Throw 'sempre falha'
        $global:t | Should Be 3
        ($global:esperas -join ',') | Should Be '2,4'
    }
}

Describe 'Provisionamento' {
    It 'valida o login (UPN)' {
        foreach ($ok in 'ana@contoso.com', 'ana.souza@sub.contoso.com.br', "o'brien@contoso.com") { (Test-Upn $ok) | Should Be $true }
        foreach ($ruim in '', 'ana', 'ana@', '@contoso.com', 'ana@contoso', 'ana souza@contoso.com', 'a@b..com') { (Test-Upn $ruim) | Should Be $false }
    }
    It 'monta o apelido de e-mail' {
        Get-ApelidoEmail 'ana.souza@contoso.com' | Should Be 'ana.souza'
        { Get-ApelidoEmail '###@contoso.com' } | Should Throw
    }
    It 'gera senha com o tamanho pedido, os 4 tipos de caractere e sem repetir' {
        $senhas = 1..30 | ForEach-Object { New-SenhaInicial }
        foreach ($s in $senhas) {
            $s.Length | Should Be 16
            ($s -cmatch '[A-Z]') | Should Be $true
            ($s -cmatch '[a-z]') | Should Be $true
            ($s -match '\d') | Should Be $true
            ($s -match '[^A-Za-z0-9]') | Should Be $true
            ($s -cmatch '[0O1lI]') | Should Be $false
        }
        (@($senhas | Select-Object -Unique)).Count | Should Be 30
        (New-SenhaInicial -Tamanho 24).Length | Should Be 24
        { New-SenhaInicial -Tamanho 8 } | Should Throw
    }
    It 'converte a SecureString de volta para texto' {
        ConvertFrom-SenhaSegura (ConvertTo-SecureString 'Abc!123xyz' -AsPlainText -Force) | Should Be 'Abc!123xyz'
    }
}

Describe 'Script de migracao com o PnP simulado' {
    # funcoes de mentira com os mesmos parametros, so para o Pester poder simular (o modulo PnP nao e necessario)
    function Connect-PnPOnline { param($Url, $ClientId, [switch]$Interactive, [switch]$ReturnConnection) }
    function Get-PnPFolderItem { param($FolderSiteRelativeUrl, $ItemType, [switch]$Recursive, $Connection) }
    function Get-PnPFile { param($Url, $Path, $FileName, [switch]$AsFile, [switch]$Force, $Connection) }
    function Get-PnPList { param($Connection) }
    function Get-PnPFolder { param($Url, $Connection) }
    function Add-PnPFolder { param($Name, $Folder, $Connection) }
    function Add-PnPFile { param($Path, $Folder, $NewFileName, $Connection) }

    $script = Join-Path $raiz 'sharepoint-migration-tool\migration-script.ps1'
    function Nova-Execucao {
        $d = Join-Path $TestDrive ([guid]::NewGuid().ToString('N')); New-Item -ItemType Directory $d | Out-Null
        $csv = Join-Path $d 'm.csv'
        Set-Content $csv "NomePasta,SourceUrl,TargetUrl`nRH,https://c.sharepoint.com/sites/a/Documentos Compartilhados/RH,https://c.sharepoint.com/sites/rh" -Encoding UTF8
        return @{ Csv = $csv; Temp = (Join-Path $d 'tmp'); Log = (Join-Path $d 'log.csv') }
    }
    function Montar-Mocks {
        Mock Connect-PnPOnline { 'conexao' }
        Mock Get-PnPFolderItem {
            @([pscustomobject]@{ ServerRelativeUrl = '/sites/a/Documentos Compartilhados/RH/2024/a.txt'; Name = 'a.txt' },
              [pscustomobject]@{ ServerRelativeUrl = '/sites/a/Documentos Compartilhados/RH/b.txt'; Name = 'b.txt' })
        }
        Mock Get-PnPFile { [IO.File]::WriteAllText((Join-Path $Path $FileName), 'conteudo') }
        Mock Get-PnPList { @([pscustomobject]@{ BaseTemplate = 101; Hidden = $false; RootFolder = [pscustomobject]@{ ServerRelativeUrl = '/sites/rh/Shared Documents' } }) }
        Mock Get-PnPFolder { $null }
        Mock Add-PnPFolder { }
        Mock Add-PnPFile { }
    }

    It 'migra os arquivos, recria as pastas, loga OK e apaga os temporarios' {
        Montar-Mocks
        $e = Nova-Execucao
        & $script -CsvPath $e.Csv -TempPath $e.Temp -LogPath $e.Log -ClientId 'id-teste' -BatchSize 1 | Out-Null
        Assert-MockCalled Add-PnPFile -Exactly 2 -Scope It
        Assert-MockCalled Add-PnPFile -Exactly 1 -Scope It -ParameterFilter { $Folder -eq '/sites/rh/Shared Documents/RH/2024' -and $NewFileName -eq 'a.txt' }
        Assert-MockCalled Add-PnPFile -Exactly 1 -Scope It -ParameterFilter { $Folder -eq '/sites/rh/Shared Documents/RH' -and $NewFileName -eq 'b.txt' }
        Assert-MockCalled Add-PnPFolder -Exactly 2 -Scope It     # RH e RH/2024, cada uma so uma vez
        Assert-MockCalled Get-PnPFile -Exactly 1 -Scope It -ParameterFilter { $Url -eq '/sites/a/Documentos Compartilhados/RH/2024/a.txt' }
        $l = @(Import-Csv $e.Log)
        @($l | Where-Object Status -eq 'OK').Count | Should Be 2
        $l[0].Arquivo | Should Be 'RH/2024/a.txt'
        @(Get-ChildItem $e.Temp -File).Count | Should Be 0
    }
    It 'rodar de novo pula o que ja deu OK (retomada)' {
        Montar-Mocks
        $e = Nova-Execucao
        & $script -CsvPath $e.Csv -TempPath $e.Temp -LogPath $e.Log -ClientId 'x' | Out-Null
        & $script -CsvPath $e.Csv -TempPath $e.Temp -LogPath $e.Log -ClientId 'x' | Out-Null
        Assert-MockCalled Add-PnPFile -Exactly 2 -Scope It
        Assert-MockCalled Get-PnPFile -Exactly 2 -Scope It
    }
    It 'erro em um arquivo vira linha ERRO com o motivo e nao para os outros' {
        Montar-Mocks
        Mock Add-PnPFile { if ($NewFileName -eq 'a.txt') { throw 'quota excedida' } }
        $e = Nova-Execucao
        & $script -CsvPath $e.Csv -TempPath $e.Temp -LogPath $e.Log -ClientId 'x' -Tentativas 1 | Out-Null
        $l = @(Import-Csv $e.Log)
        @($l | Where-Object { $_.Status -eq 'ERRO' -and $_.Arquivo -eq 'RH/2024/a.txt' -and $_.Mensagem -like '*quota excedida*' }).Count | Should Be 1
        @($l | Where-Object { $_.Status -eq 'OK' -and $_.Arquivo -eq 'RH/b.txt' }).Count | Should Be 1
        @(Get-ChildItem $e.Temp -File).Count | Should Be 0       # o temporario do arquivo com erro tambem some
    }
    It 'falha de rede passageira e refeita sem virar erro' {
        Montar-Mocks
        $global:falhou = $false
        Mock Get-PnPFile { if (-not $global:falhou) { $global:falhou = $true; throw 'timeout' }; [IO.File]::WriteAllText((Join-Path $Path $FileName), 'x') }
        $e = Nova-Execucao
        & $script -CsvPath $e.Csv -TempPath $e.Temp -LogPath $e.Log -ClientId 'x' -Tentativas 2 | Out-Null
        @(Import-Csv $e.Log | Where-Object Status -eq 'OK').Count | Should Be 2
    }
    It '-SimularApenas lista mas nao baixa, nao cria pasta e nao envia' {
        Montar-Mocks
        $e = Nova-Execucao
        & $script -CsvPath $e.Csv -TempPath $e.Temp -LogPath $e.Log -ClientId 'x' -SimularApenas | Out-Null
        Assert-MockCalled Get-PnPFolderItem -Exactly 1 -Scope It
        Assert-MockCalled Get-PnPFile -Exactly 0 -Scope It
        Assert-MockCalled Add-PnPFile -Exactly 0 -Scope It
        Assert-MockCalled Add-PnPFolder -Exactly 0 -Scope It
        Test-Path $e.Log | Should Be $false
    }
    It 'sem ClientId ou com o texto de exemplo nao conecta em nada' {
        Montar-Mocks
        $e = Nova-Execucao
        { & $script -CsvPath $e.Csv -TempPath $e.Temp -LogPath $e.Log -ClientId '' } | Should Throw 'ClientId'
        { & $script -CsvPath $e.Csv -TempPath $e.Temp -LogPath $e.Log -ClientId 'SEU_CLIENT_ID_AQUI' } | Should Throw 'ClientId'
        Assert-MockCalled Connect-PnPOnline -Exactly 0 -Scope It
    }
    It 'erro ao ler a pasta de origem vira ERRO da pasta inteira' {
        Montar-Mocks
        Mock Get-PnPFolderItem { throw 'acesso negado' }
        $e = Nova-Execucao
        & $script -CsvPath $e.Csv -TempPath $e.Temp -LogPath $e.Log -ClientId 'x' | Out-Null
        $l = @(Import-Csv $e.Log)
        $l[0].Arquivo | Should Be '(pasta inteira)'
        $l[0].Mensagem | Should Be 'acesso negado'
    }
}

Describe 'Arquivos do repositorio' {
    It 'os scripts e modulos tem sintaxe valida e so caracteres ASCII (PowerShell 5.1)' {
        $arquivos = Get-ChildItem $raiz -Recurse -Include *.ps1, *.psm1 -File | Where-Object { $_.FullName -notmatch '\\Testes\\' }
        @($arquivos).Count | Should BeGreaterThan 4
        foreach ($f in $arquivos) {
            $erros = $null; $tk = $null
            [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tk, [ref]$erros)
            @($erros).Count | Should Be 0
            (@([System.IO.File]::ReadAllBytes($f.FullName) | Where-Object { $_ -gt 127 }).Count) | Should Be 0
        }
    }
    It 'o .gitignore bloqueia a planilha e o log reais da migracao' {
        (Get-Content (Join-Path $raiz '.gitignore') -Raw) | Should Match 'MATCH_FINAL_OK\.csv'
        (Get-Content (Join-Path $raiz '.gitignore') -Raw) | Should Match 'log_migracao\.csv'
    }
}
