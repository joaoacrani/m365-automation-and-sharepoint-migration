<#
.SYNOPSIS
  Cria um usuario no Microsoft 365 (Microsoft Graph) com senha inicial e troca obrigatoria no primeiro acesso.

.DESCRIPTION
  Substitui a versao antiga, que usava o modulo MSOnline (descontinuado pela Microsoft). Valida o login, confere se o usuario
  ainda nao existe e le a senha como SecureString. Sem -SenhaInicial, gera uma senha aleatoria forte e mostra uma unica vez na tela.
  Use -WhatIf para ver o que seria criado sem criar.

.EXAMPLE
  .\user-provisioning.ps1 -Nome "Ana Souza" -Email ana.souza@contoso.com -WhatIf
  .\user-provisioning.ps1 -Nome "Ana Souza" -Email ana.souza@contoso.com -UsageLocation BR
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)] [string]$Nome,
    [Parameter(Mandatory)] [string]$Email,
    [string]$UsageLocation = 'BR',
    [securestring]$SenhaInicial
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Provisionamento.psm1') -Force

if (-not $Nome.Trim()) { throw 'Informe o nome completo.' }
if (-not (Test-Upn -Upn $Email)) { throw "E-mail/login invalido: $Email (use o formato usuario@dominio.com)." }
if ($UsageLocation -notmatch '^[A-Za-z]{2}$') { throw 'UsageLocation precisa ser o codigo do pais com 2 letras, por exemplo BR.' }

$senhaGerada = $false
if (-not $SenhaInicial) {
    $SenhaInicial = ConvertTo-SecureString (New-SenhaInicial) -AsPlainText -Force
    $senhaGerada = $true
}

Connect-MgGraph -Scopes 'User.ReadWrite.All' -NoWelcome
try {
    if (Get-MgUser -UserId $Email -ErrorAction SilentlyContinue) { throw "O usuario $Email ja existe." }

    if ($PSCmdlet.ShouldProcess($Email, "Criar o usuario '$($Nome.Trim())'")) {
        $perfil = @{ Password = (ConvertFrom-SenhaSegura -Senha $SenhaInicial); ForceChangePasswordNextSignIn = $true }
        New-MgUser -DisplayName $Nome.Trim() -UserPrincipalName $Email -MailNickname (Get-ApelidoEmail -Upn $Email) `
            -AccountEnabled -UsageLocation $UsageLocation.ToUpper() -PasswordProfile $perfil | Out-Null
        Write-Host "Usuario criado: $Email" -ForegroundColor Green
        if ($senhaGerada) {
            Write-Host "Senha inicial (aparece so agora, troca obrigatoria no primeiro acesso): $(ConvertFrom-SenhaSegura -Senha $SenhaInicial)" -ForegroundColor Yellow
        }
    }
}
finally {
    Disconnect-MgGraph | Out-Null
}
