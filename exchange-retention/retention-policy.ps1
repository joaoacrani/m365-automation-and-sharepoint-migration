<#
.SYNOPSIS
  Aplica uma politica de retencao a uma ou mais caixas de e-mail do Exchange Online e roda o Managed Folder Assistant.

.DESCRIPTION
  Confere se a politica e a caixa existem antes de mexer, trata cada caixa separadamente (uma falha nao para as outras),
  mostra um resumo no fim e sempre desconecta do Exchange. Use -WhatIf para ver o que seria feito sem alterar nada.

.EXAMPLE
  .\retention-policy.ps1 -Mailbox ana@contoso.com -PoliticaRetencao "Retencao 5 anos" -WhatIf
  'ana@contoso.com', 'bia@contoso.com' | .\retention-policy.ps1 -PoliticaRetencao "Retencao 5 anos"
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, ValueFromPipeline)] [string[]]$Mailbox,
    [Parameter(Mandatory)] [string]$PoliticaRetencao,
    [switch]$NaoExecutarAssistente
)

begin {
    $ErrorActionPreference = 'Stop'
    $caixas = New-Object System.Collections.Generic.List[string]
}
process { foreach ($m in $Mailbox) { if ($m.Trim()) { $caixas.Add($m.Trim()) } } }
end {
    Connect-ExchangeOnline -ShowBanner:$false
    $ok = 0; $falhas = 0
    try {
        $politica = Get-RetentionPolicy -Identity $PoliticaRetencao   # erro claro se a politica nao existe
        foreach ($m in $caixas) {
            try {
                $caixa = Get-Mailbox -Identity $m
                if ($PSCmdlet.ShouldProcess($caixa.PrimarySmtpAddress, "Aplicar a politica '$($politica.Name)'")) {
                    Set-Mailbox -Identity $caixa.Identity -RetentionPolicy $politica.Name
                    if (-not $NaoExecutarAssistente) { Start-ManagedFolderAssistant -Identity $caixa.Identity }
                    Write-Host "OK: $($caixa.PrimarySmtpAddress)" -ForegroundColor Green
                    $ok++
                }
            }
            catch {
                Write-Host "FALHOU: $m ($($_.Exception.Message))" -ForegroundColor Red
                $falhas++
            }
        }
    }
    finally {
        Disconnect-ExchangeOnline -Confirm:$false
    }
    Write-Host "`nResumo: $ok caixa(s) atualizada(s), $falhas com falha." -ForegroundColor Cyan
    if ($falhas -gt 0) { exit 1 }
}
