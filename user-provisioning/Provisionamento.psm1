# Funcoes do provisionamento de usuarios que nao dependem de conexao (testaveis sozinhas). Arquivo so ASCII.
Set-StrictMode -Version 2.0

# Nome de login (UPN) no formato usuario@dominio.tld
function Test-Upn {
    param([string]$Upn)
    return [bool]($Upn -match "^[A-Za-z0-9._%+'-]{1,64}@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$")
}

# Apelido de e-mail (MailNickname) a partir do UPN: a parte antes do @, so com caracteres aceitos.
function Get-ApelidoEmail {
    param([Parameter(Mandatory)] [string]$Upn)
    $a = ($Upn.Split('@')[0]) -replace "[^A-Za-z0-9._-]", ''
    if (-not $a) { throw "Nao foi possivel montar o apelido de e-mail a partir de: $Upn" }
    return $a
}

# Senha inicial aleatoria (gerador criptografico), com maiuscula, minuscula, numero e simbolo. Sem caracteres parecidos (0/O, 1/l/I).
function New-SenhaInicial {
    param([ValidateRange(12, 64)] [int]$Tamanho = 16)
    $grupos = @('ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnpqrstuvwxyz', '23456789', '!@#$%&*?-_=+')
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $sorteio = {
            param([int]$max)
            $b = New-Object byte[] 4
            do { $rng.GetBytes($b); $n = [BitConverter]::ToUInt32($b, 0) } while ($n -ge ([uint32]::MaxValue - ([uint32]::MaxValue % $max)))
            return [int]($n % $max)
        }
        $chars = New-Object System.Collections.Generic.List[char]
        foreach ($g in $grupos) { $chars.Add($g[(& $sorteio $g.Length)]) }                 # pelo menos um de cada grupo
        $todos = ($grupos -join '')
        while ($chars.Count -lt $Tamanho) { $chars.Add($todos[(& $sorteio $todos.Length)]) }
        for ($i = $chars.Count - 1; $i -gt 0; $i--) {                                      # embaralha (Fisher-Yates)
            $j = & $sorteio ($i + 1)
            $t = $chars[$i]; $chars[$i] = $chars[$j]; $chars[$j] = $t
        }
        return (-join $chars)
    } finally { $rng.Dispose() }
}

# Texto puro de uma SecureString, so pelo tempo de usar (a memoria temporaria e apagada em seguida).
function ConvertFrom-SenhaSegura {
    param([Parameter(Mandatory)] [securestring]$Senha)
    $p = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Senha)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($p) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($p) }
}

Export-ModuleMember -Function Test-Upn, Get-ApelidoEmail, New-SenhaInicial, ConvertFrom-SenhaSegura
