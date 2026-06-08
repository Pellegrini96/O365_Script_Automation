#Requires -Version 5.1
<#
.SYNOPSIS
    Ferramenta de Suporte Office 365 - YEB
.DESCRIPTION
    Menu interativo para gerenciamento completo de usuarios e recursos do Microsoft 365.
.NOTES
    Versao      : 4.0.0
    Autor       : YEB Suporte TI
    Criado em   : 2026
    Requer      : PowerShell 5.1+, Microsoft Graph SDK v2+, ExchangeOnlineManagement v3+
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =============================================================
#  CONFIGURACAO GLOBAL
# =============================================================
$Script:CONFIG = @{
    Versao             = '4.0.0'
    SenhaPadrao        = 'Yeb9034##'
    GrupoGeralEmail    = 'yebgeral@yeb.com.br'
    PaisDefault        = 'BR'
    LicensaPartNames   = @(
        'O365_BUSINESS_ESSENTIALS',
        'O365_BUSINESS_PREMIUM',
        'SPB',
        'SMB_BUSINESS_PREMIUM',
        'MICROSOFT_BUSINESS_CENTER'
    )
    ModulosNecessarios = @(
        @{ Nome = 'Microsoft.Graph';          MinVersion = '2.0.0' }
        @{ Nome = 'ExchangeOnlineManagement'; MinVersion = '3.0.0' }
        @{ Nome = 'Microsoft.Online.SharePoint.PowerShell'; MinVersion = '16.0.0' }
    )
    SubModulosGraph    = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Reports'
    )
}

$Script:CacheSkus   = $null
$Script:OperadorLog = $env:USERNAME

# =============================================================
#  UTILIDADES DE UI
# =============================================================
function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  +========================================================+" -ForegroundColor Cyan
    Write-Host "  |                                                        |" -ForegroundColor Cyan
    Write-Host "  |        YEB - FERRAMENTA DE SUPORTE OFFICE 365         |" -ForegroundColor Cyan
    Write-Host ('  |                    Versao ' + $Script:CONFIG.Versao + '                      |') -ForegroundColor Cyan
    Write-Host "  |                                                        |" -ForegroundColor Cyan
    Write-Host "  +========================================================+" -ForegroundColor Cyan
    Write-Host ""
}

function Show-Titulo {
    param([string]$Titulo, [string]$Cor = 'Cyan')
    Write-Host "  [ $Titulo ]" -ForegroundColor $Cor
    Show-Separador
    Write-Host ""
}

function Show-Separador {
    Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Mensagem,
        [ValidateSet('INFO','AVISO','ERRO','SUCESSO','AUDITORIA')][string]$Nivel = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    switch ($Nivel) {
        'INFO'      { $cor = 'Gray';    $tag = '[INFO]     ' }
        'AVISO'     { $cor = 'Yellow';  $tag = '[AVISO]    ' }
        'ERRO'      { $cor = 'Red';     $tag = '[ERRO]     ' }
        'SUCESSO'   { $cor = 'Green';   $tag = '[OK]       ' }
        'AUDITORIA' { $cor = 'Magenta'; $tag = '[AUDITORIA]' }
    }
    Write-Host "  $tag $Mensagem" -ForegroundColor $cor
    $logDir = Join-Path $PSScriptRoot 'logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    ("{0} [{1}] [{2}] {3}" -f $timestamp, $Nivel, $Script:OperadorLog, $Mensagem) |
        Out-File -FilePath (Join-Path $logDir 'o365suporte.log') -Append -Encoding UTF8
    if ($Nivel -eq 'AUDITORIA') {
        ("{0} [AUDITORIA] [{1}] {2}" -f $timestamp, $Script:OperadorLog, $Mensagem) |
            Out-File -FilePath (Join-Path $logDir 'auditoria.log') -Append -Encoding UTF8
    }
}

function PauseMenu {
    Write-Host ""
    Write-Host "  Pressione ENTER para voltar ao menu..." -ForegroundColor DarkGray -NoNewline
    Read-Host | Out-Null
}

function Read-InputObrigatorio {
    param([string]$Prompt)
    do {
        Write-Host "  $Prompt" -ForegroundColor White -NoNewline
        $valor = Read-Host
        if ([string]::IsNullOrWhiteSpace($valor)) {
            Write-Host "  >> Campo obrigatorio." -ForegroundColor Yellow
        }
    } while ([string]::IsNullOrWhiteSpace($valor))
    return $valor.Trim()
}

function Test-FormatoEmail {
    param([string]$Email)
    return $Email -match '^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$'
}

function Read-EmailObrigatorio {
    param([string]$Prompt)
    do {
        $email = Read-InputObrigatorio $Prompt
        if (-not (Test-FormatoEmail $email)) {
            Write-Host "  >> Formato de e-mail invalido." -ForegroundColor Yellow
        }
    } while (-not (Test-FormatoEmail $email))
    return $email.Trim().ToLower()
}

function New-SenhaAleatoria {
    $min  = 'abcdefghijkmnpqrstuvwxyz'
    $mai  = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    $num  = '23456789'
    $esp  = '!@#$%&*'
    $all  = $min + $mai + $num + $esp
    $rng  = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $b    = New-Object byte[] 16
    $rng.GetBytes($b)
    $s    = ''
    $s   += $mai[$b[0] % $mai.Length]
    $s   += $num[$b[1] % $num.Length]
    $s   += $esp[$b[2] % $esp.Length]
    for ($i = 3; $i -lt 12; $i++) { $s += $all[$b[$i] % $all.Length] }
    $arr = $s.ToCharArray()
    for ($i = $arr.Length - 1; $i -gt 0; $i--) {
        $j = $b[$i % $b.Length] % ($i + 1)
        $t = $arr[$i]; $arr[$i] = $arr[$j]; $arr[$j] = $t
    }
    return -join $arr
}

function Export-ParaCSV {
    param(
        [Parameter(Mandatory)][array]$Dados,
        [Parameter(Mandatory)][string]$NomeArquivo
    )
    $desktop = [Environment]::GetFolderPath('Desktop')
    $arquivo = Join-Path $desktop "$NomeArquivo`_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $Dados | Export-Csv -Path $arquivo -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Log "Exportado: $arquivo" 'SUCESSO'
    return $arquivo
}

# =============================================================
#  HELPERS DE API
# =============================================================
function Get-UsuarioByUPN {
    param(
        [Parameter(Mandatory)][string]$UPN,
        [string[]]$Propriedades = @('Id','DisplayName','UserPrincipalName','AccountEnabled','AssignedLicenses')
    )
    try {
        $props   = $Propriedades -join ','
        $usuario = Get-MgUser -Filter "userPrincipalName eq '$UPN'" -Property $props -ErrorAction Stop |
                   Select-Object -First 1
        if ($null -eq $usuario) {
            Write-Log "Nenhum usuario encontrado: '$UPN'." 'ERRO'
            return $null
        }
        return $usuario
    } catch {
        Write-Log "Erro ao buscar '$UPN': $($_.Exception.Message)" 'ERRO'
        return $null
    }
}

function Get-CatalogoSkus {
    if ($null -eq $Script:CacheSkus) {
        $Script:CacheSkus = @{}
        try {
            foreach ($s in (Get-MgSubscribedSku -ErrorAction Stop)) {
                $Script:CacheSkus[$s.SkuId] = $s.SkuPartNumber
            }
        } catch {
            Write-Log "Aviso catalogo SKUs: $($_.Exception.Message)" 'AVISO'
        }
    }
    return $Script:CacheSkus
}

function Get-NomeSku {
    param([string]$SkuId)
    $cat = Get-CatalogoSkus
    if ($cat.ContainsKey($SkuId)) { return $cat[$SkuId] }
    return $SkuId
}

# =============================================================
#  INSTALACAO DE MODULOS
# =============================================================
function Test-IsAdministrador {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Install-ModulosNecessarios {
    Show-Banner
    Write-Host "  Verificando dependencias necessarias..." -ForegroundColor Cyan
    Show-Separador
    Write-Host ""

    try {
        $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($null -ne $repo -and $repo.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }
    } catch { }

    $escopo = if (Test-IsAdministrador) { 'AllUsers' } else { 'CurrentUser' }

    foreach ($mod in $Script:CONFIG.ModulosNecessarios) {
        $nomeMod   = $mod.Nome
        $minVersao = [Version]$mod.MinVersion

        Write-Host "  Verificando: " -NoNewline -ForegroundColor Gray
        Write-Host $nomeMod -ForegroundColor White

        # Busca o modulo disponivel localmente
        $instalado = Get-Module -Name $nomeMod -ListAvailable |
                     Sort-Object Version -Descending | Select-Object -First 1

        $precisaInstalar = ($null -eq $instalado) -or ($instalado.Version -lt $minVersao)

        if ($precisaInstalar) {
            Write-Log "Instalando '$nomeMod' (escopo: $escopo)..." 'AVISO'
            try {
                # PnP.PowerShell precisa de -AllowPrerelease em alguns ambientes
                Install-Module -Name $nomeMod -MinimumVersion $minVersao -Scope $escopo `
                               -AllowClobber -Force -Repository PSGallery -ErrorAction Stop
                Write-Log "'$nomeMod' instalado com sucesso!" 'SUCESSO'
            } catch {
                Write-Log "Falha ao instalar '$nomeMod': $($_.Exception.Message)" 'ERRO'
                Write-Host "  Execute manualmente: Install-Module -Name $nomeMod -Scope CurrentUser -Force" -ForegroundColor Yellow
                if ($nomeMod -ne 'PnP.PowerShell') {
                    Read-Host "  ENTER para sair"; exit 1
                } else {
                    Write-Host "  O SharePoint ficara indisponivel ate o modulo ser instalado." -ForegroundColor Yellow
                }
            }
        } else {
            Write-Log "OK -- $nomeMod v$($instalado.Version)" 'SUCESSO'
        }

        # Importa imediatamente apos instalar/verificar
        try {
            if ($nomeMod -eq 'Microsoft.Online.SharePoint.PowerShell') {
                Import-Module 'Microsoft.Online.SharePoint.PowerShell' -DisableNameChecking -ErrorAction Stop
                Write-Log "Importado: Microsoft.Online.SharePoint.PowerShell" 'SUCESSO'
            }
        } catch {
            Write-Log "Aviso ao importar '$nomeMod': $($_.Exception.Message)" 'AVISO'
        }

        Write-Host ""
    }

    Write-Host "  Importando sub-modulos Graph..." -ForegroundColor Cyan
    Write-Host ""
    foreach ($sub in $Script:CONFIG.SubModulosGraph) {
        try {
            Import-Module $sub -DisableNameChecking -ErrorAction Stop
            Write-Log "Importado: $sub" 'SUCESSO'
        } catch {
            Write-Log "Aviso '$sub': $($_.Exception.Message)" 'AVISO'
        }
    }

    Write-Host ""
    Write-Log "Dependencias verificadas!" 'SUCESSO'
    Start-Sleep -Seconds 1
}

# =============================================================
#  AUTENTICACAO
# =============================================================
function Connect-Tenant {
    Show-Banner
    Write-Host "  Conectando ao Microsoft 365..." -ForegroundColor Cyan
    Show-Separador
    Write-Host ""

    $escopos = @(
        'User.ReadWrite.All',
        'Group.ReadWrite.All',
        'Directory.ReadWrite.All',
        'Domain.Read.All',
        'AuditLog.Read.All',
        'Reports.Read.All',
        'Mail.Send',
        'UserAuthenticationMethod.ReadWrite.All'
    )

    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -ne $ctx -and $null -ne $ctx.Account) {
            Write-Log "Sessao ativa: $($ctx.Account)" 'SUCESSO'
            $Script:OperadorLog = $ctx.Account
            return $true
        }
    } catch { }

    try {
        Write-Host "  Abrindo navegador para autenticacao..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes $escopos -ErrorAction Stop | Out-Null
        $ctx = Get-MgContext
        $Script:OperadorLog = $ctx.Account
        Write-Log "Conectado: $($ctx.Account)" 'SUCESSO'
        Start-Sleep -Seconds 1
        return $true
    } catch {
        Write-Log "Falha: $($_.Exception.Message)" 'ERRO'
        Read-Host "  ENTER para tentar novamente"
        return $false
    }
}

function Connect-ExchangeOnlineSafe {
    try {
        $cmd = Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue
        if ($null -ne $cmd) { return $true }
    } catch { }
    try {
        Write-Host "  Conectando ao Exchange Online..." -ForegroundColor Gray
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "Exchange Online conectado." 'SUCESSO'
        return $true
    } catch {
        Write-Log "Falha Exchange Online: $($_.Exception.Message)" 'ERRO'
        return $false
    }
}

# =============================================================
#  SECAO: GESTAO DE USUARIOS
# =============================================================
function Get-DominiosTenant {
    try {
        return Get-MgDomain -ErrorAction Stop |
               Where-Object { $_.IsVerified -eq $true } |
               Select-Object -ExpandProperty Id
    } catch {
        Write-Log "Erro dominios: $($_.Exception.Message)" 'ERRO'
        return @()
    }
}

function Select-Dominio {
    param([string[]]$Dominios)
    Write-Host ""
    Write-Host "  Dominios disponiveis:" -ForegroundColor Cyan
    Show-Separador
    for ($i = 0; $i -lt $Dominios.Count; $i++) {
        Write-Host "  [$($i+1)] $($Dominios[$i])" -ForegroundColor White
    }
    Show-Separador
    Write-Host ""
    $idx = 0
    do {
        Write-Host "  Selecione o dominio: " -NoNewline -ForegroundColor White
        $e = Read-Host; $p = 0
        $ok = [int]::TryParse($e, [ref]$p) -and $p -ge 1 -and $p -le $Dominios.Count
        if (-not $ok) { Write-Host "  >> Invalido." -ForegroundColor Yellow } else { $idx = $p }
    } while (-not $ok)
    return $Dominios[$idx - 1]
}

function Get-LicencaDisponivel {
    try {
        return Get-MgSubscribedSku -ErrorAction Stop |
               Where-Object {
                   $_.SkuPartNumber -in $Script:CONFIG.LicensaPartNames -and
                   ($_.PrepaidUnits.Enabled - $_.ConsumedUnits) -gt 0
               } | Select-Object -First 1
    } catch {
        Write-Log "Erro licencas: $($_.Exception.Message)" 'ERRO'
        return $null
    }
}

function New-UsuarioO365 {
    Show-Banner; Show-Titulo "CRIAR NOVO USUARIO"
    $pNome = Read-InputObrigatorio "Primeiro nome           : "
    $uNome = Read-InputObrigatorio "Ultimo nome             : "
    $nExib = "$pNome $uNome"
    Write-Host ""; Write-Host "  Buscando dominios..." -ForegroundColor Gray
    $doms  = Get-DominiosTenant
    if ($doms.Count -eq 0) { Write-Log "Nenhum dominio." 'ERRO'; PauseMenu; return }
    $dom   = Select-Dominio -Dominios $doms
    $aBase = ($pNome.ToLower() -replace '[^a-z0-9]','') + '.' + ($uNome.ToLower() -replace '[^a-z0-9]','')
    Write-Host ""; Write-Host "  Sugestao: $aBase" -ForegroundColor DarkGray
    Write-Host "  Alias [ENTER=sugestao]: " -NoNewline -ForegroundColor White
    $alias = Read-Host
    if ([string]::IsNullOrWhiteSpace($alias)) { $alias = $aBase }
    $alias = ($alias.Trim().ToLower() -replace '[^a-z0-9._-]','')
    $upn   = "$alias@$dom"
    Write-Host ""; Write-Host "  E-mail: " -NoNewline -ForegroundColor Gray; Write-Host $upn -ForegroundColor Green
    Write-Host ""; Write-Host "  Verificando disponibilidade..." -ForegroundColor Gray
    try {
        if ($null -ne (Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue)) {
            Write-Log "E-mail '$upn' ja existe." 'ERRO'; PauseMenu; return
        }
    } catch { }
    Write-Host "  Verificando licenca..." -ForegroundColor Gray
    $lic = Get-LicencaDisponivel
    if ($null -eq $lic) { Write-Log "Sem licencas disponiveis." 'ERRO'; PauseMenu; return }
    Write-Log "Licenca: $($lic.SkuPartNumber)" 'INFO'
    Write-Host ""; Show-Separador; Write-Host "  RESUMO:" -ForegroundColor Cyan; Show-Separador
    Write-Host "  Nome     : $nExib"                              -ForegroundColor White
    Write-Host "  E-mail   : $upn"                                -ForegroundColor White
    Write-Host "  Licenca  : $($lic.SkuPartNumber)"               -ForegroundColor White
    Write-Host "  Grupo    : $($Script:CONFIG.GrupoGeralEmail)"   -ForegroundColor White
    Write-Host "  Senha    : $($Script:CONFIG.SenhaPadrao)"       -ForegroundColor White
    Show-Separador; Write-Host ""
    Write-Host "  Confirmar? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    Write-Host ""; Write-Host "  Criando usuario..." -ForegroundColor Cyan
    try {
        $novo = New-MgUser -BodyParameter @{
            DisplayName = $nExib; GivenName = $pNome; Surname = $uNome
            UserPrincipalName = $upn; MailNickname = $alias
            AccountEnabled = $true; UsageLocation = $Script:CONFIG.PaisDefault
            PasswordProfile = @{ Password = $Script:CONFIG.SenhaPadrao; ForceChangePasswordNextSignIn = $true }
        } -ErrorAction Stop
        Write-Log "Usuario '$upn' criado." 'SUCESSO'
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO'; PauseMenu; return }
    Write-Host "  Atribuindo licenca..." -ForegroundColor Gray
    try {
        Set-MgUserLicense -UserId $novo.Id -BodyParameter @{
            AddLicenses = @(@{ SkuId = $lic.SkuId }); RemoveLicenses = @()
        } -ErrorAction Stop | Out-Null
        Write-Log "Licenca atribuida." 'SUCESSO'
    } catch { Write-Log "Erro licenca: $($_.Exception.Message)" 'AVISO' }
    Write-Host "  Adicionando ao grupo Yeb Geral..." -ForegroundColor Gray
    try {
        $g = Get-MgGroup -Filter "mail eq '$($Script:CONFIG.GrupoGeralEmail)'" -ErrorAction Stop | Select-Object -First 1
        if ($null -ne $g) {
            New-MgGroupMemberByRef -GroupId $g.Id -BodyParameter @{
                '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($novo.Id)"
            } -ErrorAction Stop
            Write-Log "Adicionado ao Yeb Geral." 'SUCESSO'
        } else { Write-Log "Grupo Yeb Geral nao encontrado." 'AVISO' }
    } catch { Write-Log "Erro grupo: $($_.Exception.Message)" 'AVISO' }
    Write-Log "ACAO: Criacao - $upn - $nExib" 'AUDITORIA'
    Write-Host ""; Show-Separador
    Write-Host "  >>> USUARIO CRIADO! <<<" -ForegroundColor Green; Write-Host ""
    Write-Host "  +--------------------------------------------------+" -ForegroundColor Green
    Write-Host "  |  E-mail : $upn" -ForegroundColor Green
    Write-Host "  |  Senha  : $($Script:CONFIG.SenhaPadrao) (trocar no 1 acesso) |" -ForegroundColor Green
    Write-Host "  +--------------------------------------------------+" -ForegroundColor Green
    Write-Host ""; PauseMenu
}

function Set-BloquearUsuario {
    Show-Banner; Show-Titulo "BLOQUEAR ACESSO"
    $upn = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $u = Get-UsuarioByUPN -UPN $upn
    if ($null -eq $u) { PauseMenu; return }
    Write-Host ""; Show-Separador
    Write-Host "  Nome   : $($u.DisplayName)" -ForegroundColor White
    Write-Host "  Status : $(if ($u.AccountEnabled) { 'ATIVO' } else { 'JA BLOQUEADO' })" -ForegroundColor $(if ($u.AccountEnabled) { 'Green' } else { 'Yellow' })
    Show-Separador
    if (-not $u.AccountEnabled) { Write-Log "Ja bloqueado." 'AVISO'; PauseMenu; return }
    Write-Host ""; Write-Host "  Confirmar BLOQUEIO? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        Update-MgUser -UserId $u.Id -BodyParameter @{ AccountEnabled = $false } -ErrorAction Stop
        Write-Log "ACAO: Bloqueio - $upn" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> ACESSO BLOQUEADO! <<<" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function Set-DesbloquearUsuario {
    Show-Banner; Show-Titulo "DESBLOQUEAR ACESSO"
    $upn = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $u = Get-UsuarioByUPN -UPN $upn
    if ($null -eq $u) { PauseMenu; return }
    Write-Host ""; Show-Separador
    Write-Host "  Nome   : $($u.DisplayName)" -ForegroundColor White
    Write-Host "  Status : $(if ($u.AccountEnabled) { 'ATIVO' } else { 'BLOQUEADO' })" -ForegroundColor $(if ($u.AccountEnabled) { 'Green' } else { 'Red' })
    Show-Separador
    if ($u.AccountEnabled) { Write-Log "Usuario ja esta ativo." 'AVISO'; PauseMenu; return }
    Write-Host ""; Write-Host "  Confirmar DESBLOQUEIO? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        Update-MgUser -UserId $u.Id -BodyParameter @{ AccountEnabled = $true } -ErrorAction Stop
        Write-Log "ACAO: Desbloqueio - $upn" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> ACESSO REATIVADO! <<<" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function Get-ListarUsuarios {
    Show-Banner; Show-Titulo "LISTAR USUARIOS DO TENANT"
    Write-Host "  Filtro:" -ForegroundColor Cyan
    Write-Host "  [1] Todos  [2] Ativos  [3] Bloqueados  [4] Sem licenca" -ForegroundColor White
    Write-Host ""; Write-Host "  Escolha: " -NoNewline -ForegroundColor White
    $filtro = Read-Host
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    try {
        $todos = Get-MgUser -All -Property 'Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses,CreatedDateTime' -ErrorAction Stop |
                 Where-Object { $_.UserPrincipalName -notlike '*#EXT#*' }
        switch ($filtro.Trim()) {
            '2' { $todos = $todos | Where-Object { $_.AccountEnabled -eq $true } }
            '3' { $todos = $todos | Where-Object { $_.AccountEnabled -eq $false } }
            '4' { $todos = $todos | Where-Object { $_.AssignedLicenses.Count -eq 0 } }
        }
        if (@($todos).Count -eq 0) { Write-Log "Nenhum usuario encontrado." 'AVISO'; PauseMenu; return }
        Write-Host ""; Show-Separador
        Write-Host ("  {0,-38} {1,-10} {2}" -f "E-MAIL", "STATUS", "LIC") -ForegroundColor Cyan
        Show-Separador
        $csv = @()
        foreach ($u in $todos | Sort-Object UserPrincipalName) {
            $st  = if ($u.AccountEnabled) { "Ativo   " } else { "Bloqueado" }
            $cor = if ($u.AccountEnabled) { 'White' } else { 'Yellow' }
            Write-Host ("  {0,-38} {1,-10} {2}" -f $u.UserPrincipalName, $st, @($u.AssignedLicenses).Count) -ForegroundColor $cor
            $csv += [PSCustomObject]@{ Nome = $u.DisplayName; Email = $u.UserPrincipalName; Status = $st.Trim(); Licencas = @($u.AssignedLicenses).Count; Criado = $u.CreatedDateTime }
        }
        Show-Separador
        Write-Host "  Total: $(@($todos).Count) usuario(s)" -ForegroundColor DarkGray
        Write-Host ""; Write-Host "  Exportar CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -match '^[Ss]$') { Export-ParaCSV -Dados $csv -NomeArquivo 'Usuarios' }
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    PauseMenu
}

function Set-AlterarUsuario {
    Show-Banner; Show-Titulo "ALTERAR NOME OU UPN"
    $upn = Read-EmailObrigatorio "E-mail atual            : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $u = Get-UsuarioByUPN -UPN $upn
    if ($null -eq $u) { PauseMenu; return }
    Write-Host ""; Show-Separador
    Write-Host "  Nome atual  : $($u.DisplayName)"       -ForegroundColor White
    Write-Host "  E-mail atual: $($u.UserPrincipalName)" -ForegroundColor White
    Show-Separador; Write-Host ""
    Write-Host "  [1] Nome de exibicao  [2] UPN  [3] Ambos: " -NoNewline -ForegroundColor White
    $op = Read-Host; $body = @{}
    if ($op -in @('1','3')) {
        Write-Host "  Novo nome: " -NoNewline -ForegroundColor White
        $n = Read-Host
        if (-not [string]::IsNullOrWhiteSpace($n)) { $body['DisplayName'] = $n.Trim() }
    }
    if ($op -in @('2','3')) {
        $nu = Read-EmailObrigatorio "Novo UPN                : "
        $body['UserPrincipalName'] = $nu
        $body['MailNickname'] = ($nu -split '@')[0]
    }
    if ($body.Count -eq 0) { Write-Log "Nenhuma alteracao." 'AVISO'; PauseMenu; return }
    Write-Host ""; Write-Host "  Confirmar? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        Update-MgUser -UserId $u.Id -BodyParameter $body -ErrorAction Stop
        Write-Log "ACAO: Alteracao - $upn - Campos: $($body.Keys -join ', ')" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> ALTERADO COM SUCESSO! <<<" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function Get-DetalhesUsuario {
    Show-Banner; Show-Titulo "DETALHES COMPLETOS DO USUARIO"
    $upn = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $props = @('Id','DisplayName','GivenName','Surname','UserPrincipalName','AccountEnabled','AssignedLicenses','SignInActivity','JobTitle','Department','MobilePhone')
    $u = Get-UsuarioByUPN -UPN $upn -Propriedades $props
    if ($null -eq $u) { PauseMenu; return }
    $grupos = @()
    try {
        $grupos = Get-MgUserMemberOf -UserId $u.Id -ErrorAction SilentlyContinue |
                  Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' } |
                  ForEach-Object { $_.AdditionalProperties['displayName'] }
    } catch { }
    $nLic = @(); foreach ($l in $u.AssignedLicenses) { $nLic += Get-NomeSku -SkuId $l.SkuId }
    $mfa  = 'Nao verificado'
    try {
        $m = Get-MgUserAuthenticationMethod -UserId $u.Id -ErrorAction SilentlyContinue
        $mfa = if ($m.Count -gt 1) { "Habilitado ($($m.Count) metodos)" } else { "Apenas senha (sem MFA)" }
    } catch { }
    $login = if ($null -ne $u.SignInActivity -and $null -ne $u.SignInActivity.LastSignInDateTime) {
        if ($u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) { $u.SignInActivity.LastSignInDateTime.ToString('dd/MM/yyyy HH:mm') } else { 'Nunca' }
    } else { 'Nunca' }
    Write-Host ""; Show-Separador; Write-Host "  DADOS DO USUARIO" -ForegroundColor Cyan; Show-Separador
    Write-Host "  Nome        : $($u.DisplayName)"    -ForegroundColor White
    Write-Host "  E-mail      : $($u.UserPrincipalName)" -ForegroundColor White
    Write-Host "  Cargo       : $(if ($u.JobTitle) { $u.JobTitle } else { '-' })" -ForegroundColor White
    Write-Host "  Depto       : $(if ($u.Department) { $u.Department } else { '-' })" -ForegroundColor White
    Write-Host "  Celular     : $(if ($u.MobilePhone) { $u.MobilePhone } else { '-' })" -ForegroundColor White
    Write-Host "  Status      : " -NoNewline -ForegroundColor White
    Write-Host $(if ($u.AccountEnabled) { 'ATIVO' } else { 'BLOQUEADO' }) -ForegroundColor $(if ($u.AccountEnabled) { 'Green' } else { 'Red' })
        Write-Host "  Criado em   : $(if ($u.CreatedDateTime) { $u.CreatedDateTime.ToString('dd/MM/yyyy') } else { '-' })" -ForegroundColor White
    Write-Host "  Ultimo login: $login" -ForegroundColor White
    Show-Separador; Write-Host "  LICENCAS ($($nLic.Count))" -ForegroundColor Cyan
    if ($nLic.Count -eq 0) { Write-Host "  Nenhuma." -ForegroundColor Yellow } else { $nLic | ForEach-Object { Write-Host "  - $_" -ForegroundColor White } }
    Show-Separador; Write-Host "  GRUPOS ($(@($grupos).Count))" -ForegroundColor Cyan
    if (@($grupos).Count -eq 0) { Write-Host "  Nenhum." -ForegroundColor Yellow } else { $grupos | ForEach-Object { Write-Host "  - $_" -ForegroundColor White } }
    Show-Separador; Write-Host "  SEGURANCA" -ForegroundColor Cyan
    Write-Host "  MFA : $mfa" -ForegroundColor White
    Show-Separador; Write-Host ""; PauseMenu
}

function Set-RedefinirSenha {
    Show-Banner; Show-Titulo "REDEFINIR SENHA"
    $upn = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $u = Get-UsuarioByUPN -UPN $upn
    if ($null -eq $u) { PauseMenu; return }
    Write-Host ""; Show-Separador
    Write-Host "  Nome   : $($u.DisplayName)" -ForegroundColor White
    Write-Host "  E-mail : $($u.UserPrincipalName)" -ForegroundColor White
    Show-Separador; Write-Host ""
    Write-Host ('  [1] Senha padrao (' + $Script:CONFIG.SenhaPadrao + ')  [2] Gerar aleatoria: ') -NoNewline -ForegroundColor White
    $tipo = Read-Host
    $nova = switch ($tipo.Trim()) {
        '1' { $Script:CONFIG.SenhaPadrao }
        '2' { New-SenhaAleatoria }
        default { Write-Log "Invalido." 'AVISO'; PauseMenu; return }
    }
    Write-Host ""; Write-Host "  Confirmar redefinicao? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        Update-MgUser -UserId $u.Id -BodyParameter @{
            PasswordProfile = @{ Password = $nova; ForceChangePasswordNextSignIn = $true }
        } -ErrorAction Stop
        Write-Log "ACAO: Redefinicao de senha - $upn" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> SENHA REDEFINIDA! <<<" -ForegroundColor Green; Write-Host ""
        Write-Host "  +--------------------------------------------------+" -ForegroundColor Green
        Write-Host "  |  E-mail : $upn" -ForegroundColor Green
        Write-Host "  |  Senha  : $nova" -ForegroundColor Green
        Write-Host "  |  * Trocar no proximo acesso                      |" -ForegroundColor Green
        Write-Host "  +--------------------------------------------------+" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function Remove-UsuarioO365 {
    Show-Banner; Show-Titulo "EXCLUIR USUARIO" 'Red'
    Write-Host "  ATENCAO: Remove licencas e exclui o usuario." -ForegroundColor Red
    Write-Host "  Ficara na lixeira por 30 dias." -ForegroundColor Yellow
    Write-Host ""; Show-Separador; Write-Host ""
    $upn = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $u = Get-UsuarioByUPN -UPN $upn -Propriedades @('Id','DisplayName','UserPrincipalName','AccountEnabled','AssignedLicenses')
    if ($null -eq $u) { PauseMenu; return }
    $qtd = if ($u.AssignedLicenses) { @($u.AssignedLicenses).Count } else { 0 }
    Write-Host ""; Show-Separador
    Write-Host "  Nome     : $($u.DisplayName)" -ForegroundColor White
    Write-Host "  E-mail   : $($u.UserPrincipalName)" -ForegroundColor White
    Write-Host "  Status   : $(if ($u.AccountEnabled) { 'ATIVO' } else { 'BLOQUEADO' })" -ForegroundColor White
    Write-Host "  Licencas : $qtd" -ForegroundColor White
    Show-Separador; Write-Host ""
    Write-Host "  CONFIRMACAO 1/2 [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    Write-Host "  CONFIRMACAO 2/2 - Digite o e-mail: " -NoNewline -ForegroundColor Yellow
    if ((Read-Host).Trim().ToLower() -ne $upn.ToLower()) { Write-Log "E-mail nao confere. Cancelado." 'ERRO'; PauseMenu; return }
    Write-Host ""
    Write-Host "  [1/3] Removendo licencas..." -ForegroundColor Gray
    if ($qtd -gt 0) {
        try {
            $skus = $u.AssignedLicenses | Select-Object -ExpandProperty SkuId
            Set-MgUserLicense -UserId $u.Id -BodyParameter @{ AddLicenses = @(); RemoveLicenses = $skus } -ErrorAction Stop | Out-Null
            Write-Log "$qtd licenca(s) removida(s)." 'SUCESSO'
        } catch { Write-Log "Erro licencas: $($_.Exception.Message)" 'AVISO' }
    } else { Write-Log "Sem licencas." 'INFO' }
    Write-Host "  [2/3] Bloqueando conta..." -ForegroundColor Gray
    try { Update-MgUser -UserId $u.Id -BodyParameter @{ AccountEnabled = $false } -ErrorAction Stop; Write-Log "Bloqueado." 'SUCESSO' } catch { }
    Write-Host "  [3/3] Excluindo..." -ForegroundColor Gray
    try {
        Remove-MgUser -UserId $u.Id -ErrorAction Stop
        Write-Log "ACAO: Exclusao - $upn" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> USUARIO EXCLUIDO! <<<" -ForegroundColor Green
        Write-Host "  Lixeira Azure AD: recuperavel por 30 dias." -ForegroundColor DarkGray
    } catch { Write-Log "Erro exclusao: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

# =============================================================
#  SECAO: GESTAO DE LICENCAS
# =============================================================
function Add-LicencaUsuario {
    Show-Banner; Show-Titulo "ATRIBUIR LICENCA A USUARIO"
    $upn = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $u = Get-UsuarioByUPN -UPN $upn
    if ($null -eq $u) { PauseMenu; return }
    Write-Host "  Buscando licencas disponiveis..." -ForegroundColor Gray
    try {
        $disp = Get-MgSubscribedSku -ErrorAction Stop |
                Where-Object { ($_.PrepaidUnits.Enabled - $_.ConsumedUnits) -gt 0 }
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO'; PauseMenu; return }
    if ($disp.Count -eq 0) { Write-Log "Nenhuma licenca disponivel." 'AVISO'; PauseMenu; return }
    Write-Host ""; Show-Separador; Write-Host "  Licencas disponiveis:" -ForegroundColor Cyan; Show-Separador
    for ($i = 0; $i -lt $disp.Count; $i++) {
        $liv = $disp[$i].PrepaidUnits.Enabled - $disp[$i].ConsumedUnits
        Write-Host "  [$($i+1)] $($disp[$i].SkuPartNumber)  (livres: $liv)" -ForegroundColor White
    }
    Show-Separador; Write-Host ""; $idx = 0
    do {
        Write-Host "  Selecione: " -NoNewline -ForegroundColor White
        $e = Read-Host; $p = 0
        $ok = [int]::TryParse($e, [ref]$p) -and $p -ge 1 -and $p -le $disp.Count
        if (-not $ok) { Write-Host "  >> Invalido." -ForegroundColor Yellow } else { $idx = $p }
    } while (-not $ok)
    $lic = $disp[$idx - 1]
    Write-Host ""; Write-Host "  Confirmar atribuicao de '$($lic.SkuPartNumber)' a '$upn'? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    if ([string]::IsNullOrWhiteSpace($u.UsageLocation)) {
        try { Update-MgUser -UserId $u.Id -BodyParameter @{ UsageLocation = $Script:CONFIG.PaisDefault } -ErrorAction Stop } catch { }
    }
    try {
        Set-MgUserLicense -UserId $u.Id -BodyParameter @{
            AddLicenses = @(@{ SkuId = $lic.SkuId }); RemoveLicenses = @()
        } -ErrorAction Stop | Out-Null
        Write-Log "ACAO: Atribuicao licenca - $upn - $($lic.SkuPartNumber)" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> LICENCA ATRIBUIDA! <<<" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function Remove-LicencasUsuario {
    Show-Banner; Show-Titulo "RETIRAR LICENCAS DO USUARIO"
    $upn = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $u = Get-UsuarioByUPN -UPN $upn
    if ($null -eq $u) { PauseMenu; return }
    if ($null -eq $u.AssignedLicenses -or @($u.AssignedLicenses).Count -eq 0) {
        Write-Log "Sem licencas atribuidas." 'AVISO'; PauseMenu; return
    }
    $lista = @(); Write-Host ""; Show-Separador
    Write-Host "  $($u.DisplayName) | $upn" -ForegroundColor Cyan; Show-Separador
    for ($i = 0; $i -lt @($u.AssignedLicenses).Count; $i++) {
        $nome = Get-NomeSku -SkuId $u.AssignedLicenses[$i].SkuId
        $lista += [PSCustomObject]@{ Idx = ($i+1); SkuId = $u.AssignedLicenses[$i].SkuId; Nome = $nome }
        Write-Host "  [$($i+1)] $nome" -ForegroundColor White
    }
    Write-Host "  [T] Remover TODAS" -ForegroundColor Yellow; Show-Separador; Write-Host ""
    Write-Host "  Numeros (ex: 1,2) ou T: " -NoNewline -ForegroundColor White
    $escolha = Read-Host; $skus = @()
    if ($escolha.Trim().ToUpper() -eq 'T') {
        $skus = $lista | Select-Object -ExpandProperty SkuId
    } else {
        foreach ($n in ($escolha -split ',')) {
            $p = 0
            if ([int]::TryParse($n.Trim(), [ref]$p) -and $p -ge 1 -and $p -le $lista.Count) {
                $skus += ($lista | Where-Object { $_.Idx -eq $p }).SkuId
            }
        }
    }
    if ($skus.Count -eq 0) { Write-Log "Nenhuma selecionada." 'AVISO'; PauseMenu; return }
    Write-Host ""; Write-Host "  Serao removidas:" -ForegroundColor Cyan
    foreach ($s in $skus) { Write-Host "  - $(Get-NomeSku -SkuId $s)" -ForegroundColor Yellow }
    Write-Host ""; Write-Host "  Confirmar? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        Set-MgUserLicense -UserId $u.Id -BodyParameter @{ AddLicenses = @(); RemoveLicenses = $skus } -ErrorAction Stop | Out-Null
        Write-Log "ACAO: Remocao licencas - $upn - $($skus.Count) removida(s)" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> LICENCAS REMOVIDAS! <<<" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function Get-RelatorioLicencas {
    Show-Banner; Show-Titulo "RELATORIO DE LICENCAS DO TENANT"
    Write-Host "  Buscando..." -ForegroundColor Gray
    try {
        $skus = Get-MgSubscribedSku -ErrorAction Stop
        Write-Host ""; Show-Separador
        Write-Host ("  {0,-42} {1,8} {2,8} {3,8}" -f "LICENCA", "TOTAL", "USADAS", "LIVRES") -ForegroundColor Cyan
        Show-Separador
        $csv = @()
        foreach ($s in $skus | Sort-Object SkuPartNumber) {
            $total = $s.PrepaidUnits.Enabled; $usadas = $s.ConsumedUnits; $livres = $total - $usadas
            $cor   = if ($livres -le 0) { 'Red' } elseif ($livres -le 5) { 'Yellow' } else { 'White' }
            Write-Host ("  {0,-42} {1,8} {2,8} {3,8}" -f $s.SkuPartNumber, $total, $usadas, $livres) -ForegroundColor $cor
            $csv += [PSCustomObject]@{ Licenca = $s.SkuPartNumber; Total = $total; Usadas = $usadas; Livres = $livres }
        }
        Show-Separador; Write-Host ""
        Write-Host "  Exportar CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -match '^[Ss]$') { Export-ParaCSV -Dados $csv -NomeArquivo 'Licencas' }
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    PauseMenu
}

function Get-UsuariosSemLicenca {
    Show-Banner; Show-Titulo "USUARIOS SEM LICENCA"
    Write-Host "  Buscando..." -ForegroundColor Gray
    try {
        $res = Get-MgUser -All -Property 'Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses,CreatedDateTime' -ErrorAction Stop |
               Where-Object { $_.UserPrincipalName -notlike '*#EXT#*' -and $_.AssignedLicenses.Count -eq 0 }
        if (@($res).Count -eq 0) { Write-Log "Todos possuem licenca." 'SUCESSO'; PauseMenu; return }
        Write-Host ""; Show-Separador
        Write-Host ("  {0,-40} {1}" -f "E-MAIL", "STATUS") -ForegroundColor Cyan; Show-Separador
        $csv = @()
        foreach ($u in $res | Sort-Object UserPrincipalName) {
            $st = if ($u.AccountEnabled) { "Ativo" } else { "Bloqueado" }
            Write-Host ("  {0,-40} {1}" -f $u.UserPrincipalName, $st) -ForegroundColor $(if ($u.AccountEnabled) { 'White' } else { 'Yellow' })
            $csv += [PSCustomObject]@{ Nome = $u.DisplayName; Email = $u.UserPrincipalName; Status = $st }
        }
        Show-Separador; Write-Host "  Total: $(@($res).Count)" -ForegroundColor DarkGray; Write-Host ""
        Write-Host "  Exportar CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -match '^[Ss]$') { Export-ParaCSV -Dados $csv -NomeArquivo 'SemLicenca' }
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    PauseMenu
}

# =============================================================
#  SECAO: GESTAO DE GRUPOS
# =============================================================
function Add-MembroGrupo {
    Show-Banner; Show-Titulo "ADICIONAR USUARIO A GRUPO"
    $upn = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $u = Get-UsuarioByUPN -UPN $upn
    if ($null -eq $u) { PauseMenu; return }
    try { $gs = Get-MgGroup -All -Property 'Id,DisplayName,Mail' -ErrorAction Stop | Sort-Object DisplayName }
    catch { Write-Log "Erro grupos: $($_.Exception.Message)" 'ERRO'; PauseMenu; return }
    Write-Host ""; Show-Separador; Write-Host "  Grupos disponiveis:" -ForegroundColor Cyan; Show-Separador
    for ($i = 0; $i -lt $gs.Count; $i++) {
        $m = if ($gs[$i].Mail) { " | $($gs[$i].Mail)" } else { "" }
        Write-Host "  [$($i+1)] $($gs[$i].DisplayName)$m" -ForegroundColor White
    }
    Show-Separador; Write-Host ""; $idx = 0
    do {
        Write-Host "  Selecione: " -NoNewline -ForegroundColor White
        $e = Read-Host; $p = 0
        $ok = [int]::TryParse($e, [ref]$p) -and $p -ge 1 -and $p -le $gs.Count
        if (-not $ok) { Write-Host "  >> Invalido." -ForegroundColor Yellow } else { $idx = $p }
    } while (-not $ok)
    $g = $gs[$idx - 1]
    try {
        $jaMembro = Get-MgGroupMember -GroupId $g.Id -ErrorAction SilentlyContinue | Where-Object { $_.Id -eq $u.Id }
        if ($null -ne $jaMembro) { Write-Log "Ja e membro." 'AVISO'; PauseMenu; return }
    } catch { }
    Write-Host ""; Write-Host "  Adicionar '$($u.DisplayName)' ao '$($g.DisplayName)'? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        New-MgGroupMemberByRef -GroupId $g.Id -BodyParameter @{
            '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($u.Id)"
        } -ErrorAction Stop
        Write-Log "ACAO: Adicao ao grupo - $upn - $($g.DisplayName)" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> ADICIONADO AO GRUPO! <<<" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function Remove-MembroGrupo {
    Show-Banner; Show-Titulo "REMOVER USUARIO DE GRUPO"
    $upn = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $u = Get-UsuarioByUPN -UPN $upn
    if ($null -eq $u) { PauseMenu; return }
    try {
        $mGrupos = Get-MgUserMemberOf -UserId $u.Id -ErrorAction Stop |
                   Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }
    } catch { Write-Log "Erro grupos: $($_.Exception.Message)" 'ERRO'; PauseMenu; return }
    if ($mGrupos.Count -eq 0) { Write-Log "Sem grupos." 'AVISO'; PauseMenu; return }
    Write-Host ""; Show-Separador; Write-Host "  Grupos de '$($u.DisplayName)':" -ForegroundColor Cyan; Show-Separador
    for ($i = 0; $i -lt $mGrupos.Count; $i++) {
        Write-Host "  [$($i+1)] $($mGrupos[$i].AdditionalProperties['displayName'])" -ForegroundColor White
    }
    Show-Separador; Write-Host ""; $idx = 0
    do {
        Write-Host "  Remover qual: " -NoNewline -ForegroundColor White
        $e = Read-Host; $p = 0
        $ok = [int]::TryParse($e, [ref]$p) -and $p -ge 1 -and $p -le $mGrupos.Count
        if (-not $ok) { Write-Host "  >> Invalido." -ForegroundColor Yellow } else { $idx = $p }
    } while (-not $ok)
    $gSel  = $mGrupos[$idx - 1]
    $gNome = $gSel.AdditionalProperties['displayName']
    Write-Host ""; Write-Host "  Remover de '$gNome'? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        Remove-MgGroupMemberByRef -GroupId $gSel.Id -DirectoryObjectId $u.Id -ErrorAction Stop
        Write-Log "ACAO: Remocao grupo - $upn - $gNome" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> REMOVIDO DO GRUPO! <<<" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function Get-ListarGrupos {
    Show-Banner; Show-Titulo "LISTAR GRUPOS E MEMBROS"
    Write-Host "  Buscando grupos..." -ForegroundColor Gray
    try {
        $gs = Get-MgGroup -All -Property 'Id,DisplayName,Mail' -ErrorAction Stop | Sort-Object DisplayName
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO'; PauseMenu; return }
    Write-Host ""; Show-Separador
    Write-Host ("  {0,-4} {1,-35} {2}" -f "#", "NOME", "E-MAIL") -ForegroundColor Cyan; Show-Separador
    for ($i = 0; $i -lt $gs.Count; $i++) {
        $m = if ($gs[$i].Mail) { $gs[$i].Mail } else { '-' }
        Write-Host ("  {0,-4} {1,-35} {2}" -f "[$($i+1)]", $gs[$i].DisplayName, $m) -ForegroundColor White
    }
    Show-Separador; Write-Host ""
    Write-Host "  Ver membros de um grupo? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { PauseMenu; return }
    $idx = 0
    do {
        Write-Host "  Numero do grupo: " -NoNewline -ForegroundColor White
        $e = Read-Host; $p = 0
        $ok = [int]::TryParse($e, [ref]$p) -and $p -ge 1 -and $p -le $gs.Count
        if (-not $ok) { Write-Host "  >> Invalido." -ForegroundColor Yellow } else { $idx = $p }
    } while (-not $ok)
    $g = $gs[$idx - 1]
    Write-Host "  Buscando membros..." -ForegroundColor Gray
    try {
        $mbs = Get-MgGroupMember -GroupId $g.Id -All -ErrorAction Stop
        Write-Host ""; Show-Separador; Write-Host "  Membros: $($g.DisplayName)" -ForegroundColor Cyan; Show-Separador
        if ($mbs.Count -eq 0) { Write-Host "  Sem membros." -ForegroundColor Yellow } else {
            $csv = @()
            foreach ($m in $mbs) {
                $n = $m.AdditionalProperties['displayName']
                $e = $m.AdditionalProperties['userPrincipalName']
                Write-Host "  - $n  |  $e" -ForegroundColor White
                $csv += [PSCustomObject]@{ Nome = $n; Email = $e }
            }
            Show-Separador; Write-Host "  Total: $($mbs.Count)" -ForegroundColor DarkGray; Write-Host ""
            Write-Host "  Exportar CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
            if (Read-Host -match '^[Ss]$') { Export-ParaCSV -Dados $csv -NomeArquivo "Membros_$($g.DisplayName -replace '\s','_')" }
        }
    } catch { Write-Log "Erro membros: $($_.Exception.Message)" 'ERRO' }
    PauseMenu
}

# =============================================================
#  SECAO: SEGURANCA E CONFORMIDADE
# =============================================================
function Revoke-SessoesUsuario {
    Show-Banner; Show-Titulo "REVOGAR SESSOES ATIVAS"
    $upn = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $u = Get-UsuarioByUPN -UPN $upn
    if ($null -eq $u) { PauseMenu; return }
    Write-Host ""; Show-Separador
    Write-Host "  Nome   : $($u.DisplayName)" -ForegroundColor White
    Write-Host "  E-mail : $($u.UserPrincipalName)" -ForegroundColor White
    Show-Separador; Write-Host ""
    Write-Host "  Desconecta Teams, Outlook, etc. Confirmar? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        Revoke-MgUserSignInSession -UserId $u.Id -ErrorAction Stop | Out-Null
        Write-Log "ACAO: Revogacao sessoes - $upn" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> SESSOES REVOGADAS! <<<" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function Set-ForcarMFA {
    Show-Banner; Show-Titulo "FORCAR MFA"
    $upn = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $u = Get-UsuarioByUPN -UPN $upn
    if ($null -eq $u) { PauseMenu; return }
    Write-Host ""; Show-Separador
    Write-Host "  Nome   : $($u.DisplayName)" -ForegroundColor White
    Show-Separador; Write-Host ""
    Write-Host "  Revoga sessoes e exige MFA no proximo login. Confirmar? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        Revoke-MgUserSignInSession -UserId $u.Id -ErrorAction Stop | Out-Null
        Write-Log "ACAO: Forcamento MFA + revogacao - $upn" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> SESSOES REVOGADAS - MFA EXIGIDO NO PROXIMO LOGIN! <<<" -ForegroundColor Green
        Write-Host "  Dica: configure Acesso Condicional no portal Azure para MFA permanente." -ForegroundColor DarkGray
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function Get-UsuariosInativos {
    Show-Banner; Show-Titulo "USUARIOS INATIVOS"
    Write-Host "  Sem login ha quantos dias? " -NoNewline -ForegroundColor White
    $ds = Read-Host; $dias = 30
    [int]::TryParse($ds, [ref]$dias) | Out-Null
    if ($dias -le 0) { $dias = 30 }
    $corte = (Get-Date).AddDays(-$dias).ToUniversalTime()
    Write-Host ""; Write-Host "  Buscando (pode demorar)..." -ForegroundColor Gray
    try {
        $todos = Get-MgUser -All -Property 'Id,DisplayName,UserPrincipalName,AccountEnabled,SignInActivity,CreatedDateTime' -ErrorAction Stop |
                 Where-Object { $_.UserPrincipalName -notlike '*#EXT#*' }
        $inat  = $todos | Where-Object {
            $_.AccountEnabled -eq $true -and (
                $null -eq $_.SignInActivity -or
                $null -eq $_.SignInActivity.LastSignInDateTime -or
                $_.SignInActivity.LastSignInDateTime -lt $corte
            )
        }
        if ($inat.Count -eq 0) { Write-Log "Nenhum inativo encontrado." 'SUCESSO'; PauseMenu; return }
        Write-Host ""; Show-Separador
        Write-Host ("  {0,-40} {1}" -f "E-MAIL", "ULTIMO LOGIN") -ForegroundColor Cyan; Show-Separador
        $csv = @()
        foreach ($u in $inat | Sort-Object UserPrincipalName) {
            $lg = if ($null -ne $u.SignInActivity -and $null -ne $u.SignInActivity.LastSignInDateTime) {
                if ($u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) { $u.SignInActivity.LastSignInDateTime.ToString('dd/MM/yyyy HH:mm') } else { 'Nunca' }
            } else { "Nunca" }
            Write-Host ("  {0,-40} {1}" -f $u.UserPrincipalName, $lg) -ForegroundColor Yellow
            $csv += [PSCustomObject]@{ Nome = $u.DisplayName; Email = $u.UserPrincipalName; UltimoLogin = $lg; Criado = $u.CreatedDateTime }
        }
        Show-Separador; Write-Host "  Total: $($inat.Count)" -ForegroundColor DarkGray; Write-Host ""
        Write-Host "  Exportar CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -match '^[Ss]$') { Export-ParaCSV -Dados $csv -NomeArquivo "Inativos_${dias}dias" }
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    PauseMenu
}

function Get-UsuariosSemMFA {
    Show-Banner; Show-Titulo "USUARIOS SEM MFA"
    Write-Host "  Buscando (pode demorar para tenants grandes)..." -ForegroundColor Gray
    try {
        $todos = Get-MgUser -All -Property 'Id,DisplayName,UserPrincipalName,AccountEnabled' -ErrorAction Stop |
                 Where-Object { $_.AccountEnabled -eq $true -and $_.UserPrincipalName -notlike '*#EXT#*' }
        $semMFA = @(); $c = 0
        foreach ($u in $todos) {
            $c++
            Write-Host "  Verificando $c/$(@($todos).Count): $($u.UserPrincipalName)...          `r" -NoNewline -ForegroundColor DarkGray
            try {
                $m = Get-MgUserAuthenticationMethod -UserId $u.Id -ErrorAction SilentlyContinue
                if ($m.Count -le 1) { $semMFA += $u }
            } catch { $semMFA += $u }
        }
        Write-Host "  Concluido.                                              " -ForegroundColor Gray
        if ($semMFA.Count -eq 0) { Write-Log "Todos com MFA." 'SUCESSO'; PauseMenu; return }
        Write-Host ""; Show-Separador; Write-Host "  Sem MFA ($($semMFA.Count)):" -ForegroundColor Cyan; Show-Separador
        $csv = @()
        foreach ($u in $semMFA | Sort-Object UserPrincipalName) {
            Write-Host "  - $($u.UserPrincipalName)" -ForegroundColor Yellow
            $csv += [PSCustomObject]@{ Nome = $u.DisplayName; Email = $u.UserPrincipalName }
        }
        Show-Separador; Write-Host ""
        Write-Host "  Exportar CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -match '^[Ss]$') { Export-ParaCSV -Dados $csv -NomeArquivo 'SemMFA' }
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    PauseMenu
}

# =============================================================
#  SECAO: EXCHANGE ONLINE
# =============================================================
function New-CaixaCompartilhada {
    Show-Banner; Show-Titulo "CRIAR CAIXA COMPARTILHADA"
    if (-not (Connect-ExchangeOnlineSafe)) { PauseMenu; return }
    $nome  = Read-InputObrigatorio "Nome da caixa           : "
    $alias = Read-InputObrigatorio "Alias (sem @dominio)    : "
    $alias = ($alias.Trim().ToLower() -replace '[^a-z0-9._-]','')
    $doms  = Get-DominiosTenant
    if ($doms.Count -eq 0) { Write-Log "Nenhum dominio." 'ERRO'; PauseMenu; return }
    $dom   = Select-Dominio -Dominios $doms
    $email = "$alias@$dom"
    Write-Host ""; Show-Separador
    Write-Host "  Nome   : $nome"  -ForegroundColor White
    Write-Host "  E-mail : $email" -ForegroundColor White
    Show-Separador; Write-Host ""
    Write-Host "  Confirmar? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        New-Mailbox -Shared -Name $nome -Alias $alias -PrimarySmtpAddress $email -ErrorAction Stop | Out-Null
        Write-Log "ACAO: Caixa compartilhada criada - $email" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> CAIXA CRIADA: $email <<<" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function Add-PermissaoCaixaCompartilhada {
    Show-Banner; Show-Titulo "PERMISSAO EM CAIXA COMPARTILHADA"
    if (-not (Connect-ExchangeOnlineSafe)) { PauseMenu; return }
    $caixa   = Read-EmailObrigatorio "E-mail da caixa         : "
    $usuario = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  [1] FullAccess  [2] SendAs  [3] Ambos: " -NoNewline -ForegroundColor White
    $tipo = Read-Host
    Write-Host ""; Write-Host "  Confirmar? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        if ($tipo -in @('1','3')) {
            Add-MailboxPermission -Identity $caixa -User $usuario -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop | Out-Null
            Write-Log "FullAccess: $usuario -> $caixa" 'SUCESSO'
        }
        if ($tipo -in @('2','3')) {
            Add-RecipientPermission -Identity $caixa -Trustee $usuario -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
            Write-Log "SendAs: $usuario -> $caixa" 'SUCESSO'
        }
        Write-Log "ACAO: Permissao caixa - $caixa -> $usuario" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> PERMISSAO CONCEDIDA! <<<" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function New-SalaOuEquipamento {
    Show-Banner; Show-Titulo "CRIAR SALA OU EQUIPAMENTO"
    if (-not (Connect-ExchangeOnlineSafe)) { PauseMenu; return }
    Write-Host "  [1] Sala de reuniao  [2] Equipamento: " -NoNewline -ForegroundColor White
    $tipo = Read-Host
    if ($tipo -notin @('1','2')) { Write-Log "Invalido." 'AVISO'; PauseMenu; return }
    $nome  = Read-InputObrigatorio "Nome do recurso         : "
    $alias = Read-InputObrigatorio "Alias (sem @dominio)    : "
    $alias = ($alias.Trim().ToLower() -replace '[^a-z0-9._-]','')
    $doms  = Get-DominiosTenant
    if ($doms.Count -eq 0) { Write-Log "Nenhum dominio." 'ERRO'; PauseMenu; return }
    $dom   = Select-Dominio -Dominios $doms
    $email = "$alias@$dom"
    $tNome = if ($tipo -eq '1') { 'Sala' } else { 'Equipamento' }
    Write-Host ""; Show-Separador
    Write-Host "  Tipo   : $tNome" -ForegroundColor White
    Write-Host "  Nome   : $nome"  -ForegroundColor White
    Write-Host "  E-mail : $email" -ForegroundColor White
    Show-Separador; Write-Host ""
    Write-Host "  Confirmar? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        if ($tipo -eq '1') {
            New-Mailbox -Room -Name $nome -Alias $alias -PrimarySmtpAddress $email -ErrorAction Stop | Out-Null
        } else {
            New-Mailbox -Equipment -Name $nome -Alias $alias -PrimarySmtpAddress $email -ErrorAction Stop | Out-Null
        }
        Write-Log "ACAO: Criacao $tNome - $email" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> $($tNome.ToUpper()) CRIADO: $email <<<" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

function Set-RedirecionamentoEmail {
    Show-Banner; Show-Titulo "REDIRECIONAMENTO DE E-MAIL"
    if (-not (Connect-ExchangeOnlineSafe)) { PauseMenu; return }
    $origem  = Read-EmailObrigatorio "E-mail de origem        : "
    $destino = Read-EmailObrigatorio "E-mail de destino       : "
    Write-Host ""; Write-Host "  Manter copia na origem? [S/N]: " -NoNewline -ForegroundColor Yellow
    $manter = (Read-Host) -match '^[Ss]$'
    Write-Host ""; Show-Separador
    Write-Host "  Origem  : $origem"  -ForegroundColor White
    Write-Host "  Destino : $destino" -ForegroundColor White
    Write-Host "  Copia   : $(if ($manter) { 'Sim' } else { 'Nao' })" -ForegroundColor White
    Show-Separador; Write-Host ""
    Write-Host "  Confirmar? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    try {
        Set-Mailbox -Identity $origem -ForwardingSmtpAddress $destino -DeliverToMailboxAndForward $manter -ErrorAction Stop
        Write-Log "ACAO: Redirecionamento - $origem -> $destino" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> REDIRECIONAMENTO CONFIGURADO! <<<" -ForegroundColor Green
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    Write-Host ""; PauseMenu
}

# =============================================================
#  SECAO: OFFBOARDING E ONBOARDING
# =============================================================
function Invoke-OffboardingCompleto {
    Show-Banner; Show-Titulo "OFFBOARDING COMPLETO" 'Red'
    Write-Host "  Executa: bloqueio + revogacao + licencas + grupos + resposta automatica" -ForegroundColor Yellow
    Write-Host ""; Show-Separador; Write-Host ""
    $upn = Read-EmailObrigatorio "E-mail do colaborador   : "
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    $u = Get-UsuarioByUPN -UPN $upn -Propriedades @('Id','DisplayName','UserPrincipalName','AccountEnabled','AssignedLicenses')
    if ($null -eq $u) { PauseMenu; return }
    Write-Host ""; Show-Separador
    Write-Host "  Nome   : $($u.DisplayName)" -ForegroundColor White
    Write-Host "  Status : $(if ($u.AccountEnabled) { 'ATIVO' } else { 'BLOQUEADO' })" -ForegroundColor White
    Show-Separador; Write-Host ""
    Write-Host "  Configurar resposta automatica? [S/N]: " -NoNewline -ForegroundColor White
    $confRA = (Read-Host) -match '^[Ss]$'
    $msgRA = ""; $emailGestor = ""
    if ($confRA) {
        $msgRA       = Read-InputObrigatorio "Mensagem de ausencia    : "
        $emailGestor = Read-EmailObrigatorio "E-mail do gestor        : "
    }
    Write-Host ""; Write-Host "  CONFIRMACAO 1/2 [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    Write-Host "  CONFIRMACAO 2/2 - Digite o e-mail: " -NoNewline -ForegroundColor Yellow
    if ((Read-Host).Trim().ToLower() -ne $upn.ToLower()) { Write-Log "Nao confere." 'ERRO'; PauseMenu; return }
    Write-Host ""
    Write-Host "  [1/5] Bloqueando conta..." -ForegroundColor Gray
    try { Update-MgUser -UserId $u.Id -BodyParameter @{ AccountEnabled = $false } -ErrorAction Stop; Write-Log "Bloqueado." 'SUCESSO' }
    catch { Write-Log "Erro bloqueio: $($_.Exception.Message)" 'AVISO' }
    Write-Host "  [2/5] Revogando sessoes..." -ForegroundColor Gray
    try { Revoke-MgUserSignInSession -UserId $u.Id -ErrorAction Stop | Out-Null; Write-Log "Sessoes revogadas." 'SUCESSO' }
    catch { Write-Log "Erro sessoes: $($_.Exception.Message)" 'AVISO' }
    Write-Host "  [3/5] Removendo licencas..." -ForegroundColor Gray
    try {
        if (@($u.AssignedLicenses).Count -gt 0) {
            $skus = $u.AssignedLicenses | Select-Object -ExpandProperty SkuId
            Set-MgUserLicense -UserId $u.Id -BodyParameter @{ AddLicenses = @(); RemoveLicenses = $skus } -ErrorAction Stop | Out-Null
            Write-Log "$($skus.Count) licenca(s) removida(s)." 'SUCESSO'
        } else { Write-Log "Sem licencas." 'INFO' }
    } catch { Write-Log "Erro licencas: $($_.Exception.Message)" 'AVISO' }
    Write-Host "  [4/5] Removendo grupos..." -ForegroundColor Gray
    try {
        $gList = Get-MgUserMemberOf -UserId $u.Id -ErrorAction SilentlyContinue |
                 Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }
        $rem = 0
        foreach ($g in $gList) {
            try { Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $u.Id -ErrorAction Stop; $rem++ } catch { }
        }
        Write-Log "$rem grupo(s) removido(s)." 'SUCESSO'
    } catch { Write-Log "Erro grupos: $($_.Exception.Message)" 'AVISO' }
    Write-Host "  [5/5] Resposta automatica..." -ForegroundColor Gray
    if ($confRA -and (Connect-ExchangeOnlineSafe)) {
        try {
            Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled `
                -InternalMessage $msgRA -ExternalMessage $msgRA -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($emailGestor)) {
                Set-Mailbox -Identity $upn -ForwardingSmtpAddress $emailGestor -DeliverToMailboxAndForward $true -ErrorAction SilentlyContinue
                Write-Log "Encaminhamento para $emailGestor configurado." 'SUCESSO'
            }
            Write-Log "Resposta automatica configurada." 'SUCESSO'
        } catch { Write-Log "Erro resposta automatica: $($_.Exception.Message)" 'AVISO' }
    } else { Write-Log "Resposta automatica pulada." 'INFO' }
    Write-Log "ACAO: Offboarding completo - $upn" 'AUDITORIA'
    Write-Host ""; Show-Separador
    Write-Host "  >>> OFFBOARDING CONCLUIDO! <<<" -ForegroundColor Green
    Write-Host "  Usuario : $($u.DisplayName) | $upn" -ForegroundColor White
    Write-Host ""; PauseMenu
}

function Invoke-OnboardingCompleto {
    Show-Banner; Show-Titulo "ONBOARDING COMPLETO"
    Write-Host "  Executa: criacao + licenca + grupos + e-mail de boas-vindas" -ForegroundColor Cyan
    Write-Host ""; Show-Separador; Write-Host ""
    $pNome = Read-InputObrigatorio "Primeiro nome           : "
    $uNome = Read-InputObrigatorio "Ultimo nome             : "
    $cargo = Read-InputObrigatorio "Cargo                   : "
    $depto = Read-InputObrigatorio "Departamento            : "
    $nExib = "$pNome $uNome"
    Write-Host ""; Write-Host "  Buscando dominios..." -ForegroundColor Gray
    $doms  = Get-DominiosTenant
    if ($doms.Count -eq 0) { Write-Log "Nenhum dominio." 'ERRO'; PauseMenu; return }
    $dom   = Select-Dominio -Dominios $doms
    $aBase = ($pNome.ToLower() -replace '[^a-z0-9]','') + '.' + ($uNome.ToLower() -replace '[^a-z0-9]','')
    Write-Host ""; Write-Host "  Sugestao: $aBase" -ForegroundColor DarkGray
    Write-Host "  Alias [ENTER=sugestao]: " -NoNewline -ForegroundColor White
    $alias = Read-Host
    if ([string]::IsNullOrWhiteSpace($alias)) { $alias = $aBase }
    $alias = ($alias.Trim().ToLower() -replace '[^a-z0-9._-]','')
    $upn   = "$alias@$dom"
    $eContato = Read-EmailObrigatorio "E-mail pessoal (boas-vindas): "
    Write-Host ""; Write-Host "  Verificando..." -ForegroundColor Gray
    try {
        if ($null -ne (Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue)) {
            Write-Log "E-mail '$upn' ja existe." 'ERRO'; PauseMenu; return
        }
    } catch { }
    $lic = Get-LicencaDisponivel
    if ($null -eq $lic) { Write-Log "Sem licencas." 'ERRO'; PauseMenu; return }
    Write-Host ""; Show-Separador; Write-Host "  RESUMO ONBOARDING:" -ForegroundColor Cyan; Show-Separador
    Write-Host "  Nome         : $nExib"   -ForegroundColor White
    Write-Host "  E-mail       : $upn"      -ForegroundColor White
    Write-Host "  Cargo        : $cargo"    -ForegroundColor White
    Write-Host "  Departamento : $depto"    -ForegroundColor White
    Write-Host "  Licenca      : $($lic.SkuPartNumber)" -ForegroundColor White
    Write-Host "  Boas-vindas  : $eContato" -ForegroundColor White
    Show-Separador; Write-Host ""
    Write-Host "  Confirmar onboarding? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    Write-Host ""
    Write-Host "  [1/5] Criando usuario..." -ForegroundColor Gray
    $novo = $null
    try {
        $novo = New-MgUser -BodyParameter @{
            DisplayName = $nExib; GivenName = $pNome; Surname = $uNome
            UserPrincipalName = $upn; MailNickname = $alias
            JobTitle = $cargo; Department = $depto
            AccountEnabled = $true; UsageLocation = $Script:CONFIG.PaisDefault
            PasswordProfile = @{ Password = $Script:CONFIG.SenhaPadrao; ForceChangePasswordNextSignIn = $true }
        } -ErrorAction Stop
        Write-Log "Usuario criado: $upn" 'SUCESSO'
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO'; PauseMenu; return }
    Write-Host "  [2/5] Atribuindo licenca..." -ForegroundColor Gray
    try {
        Set-MgUserLicense -UserId $novo.Id -BodyParameter @{
            AddLicenses = @(@{ SkuId = $lic.SkuId }); RemoveLicenses = @()
        } -ErrorAction Stop | Out-Null
        Write-Log "Licenca atribuida." 'SUCESSO'
    } catch { Write-Log "Erro licenca: $($_.Exception.Message)" 'AVISO' }
    Write-Host "  [3/5] Grupo Yeb Geral..." -ForegroundColor Gray
    try {
        $gG = Get-MgGroup -Filter "mail eq '$($Script:CONFIG.GrupoGeralEmail)'" -ErrorAction Stop | Select-Object -First 1
        if ($null -ne $gG) {
            New-MgGroupMemberByRef -GroupId $gG.Id -BodyParameter @{
                '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($novo.Id)"
            } -ErrorAction Stop
            Write-Log "Adicionado ao Yeb Geral." 'SUCESSO'
        } else { Write-Log "Grupo Yeb Geral nao encontrado." 'AVISO' }
    } catch { Write-Log "Erro grupo: $($_.Exception.Message)" 'AVISO' }
    Write-Host "  [4/5] Grupos adicionais? [S/N]: " -NoNewline -ForegroundColor Gray
    if ((Read-Host) -match '^[Ss]$') {
        try {
            $tGs = Get-MgGroup -All -Property 'Id,DisplayName' -ErrorAction Stop | Sort-Object DisplayName
            do {
                Write-Host ""; for ($i = 0; $i -lt $tGs.Count; $i++) { Write-Host "  [$($i+1)] $($tGs[$i].DisplayName)" -ForegroundColor White }
                Write-Host "  [0] Continuar" -ForegroundColor DarkGray
                Write-Host "  Selecione: " -NoNewline -ForegroundColor White
                $e = Read-Host; $p = 0
                if ([int]::TryParse($e, [ref]$p) -and $p -ge 1 -and $p -le $tGs.Count) {
                    try {
                        New-MgGroupMemberByRef -GroupId $tGs[$p-1].Id -BodyParameter @{
                            '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($novo.Id)"
                        } -ErrorAction Stop
                        Write-Log ("Adicionado ao grupo: " + $tGs[$p-1].DisplayName) 'SUCESSO'
                    } catch { Write-Log ("Erro grupo: " + $tGs[$p-1].DisplayName) 'AVISO' }
                } elseif ($p -eq 0) { break }
            } while ($true)
        } catch { Write-Log "Erro grupos adicionais: $($_.Exception.Message)" 'AVISO' }
    }
    Write-Host "  [5/5] Enviando e-mail de boas-vindas..." -ForegroundColor Gray
    try {
        $corpo = "Ola $pNome,<br><br>Seja bem-vindo(a) a YEB!<br><br>Suas credenciais:<br><b>E-mail:</b> $upn<br><b>Senha inicial:</b> $($Script:CONFIG.SenhaPadrao)<br><br>Altere sua senha no primeiro acesso.<br><br>Att,<br>TI - YEB"
        $ctx = Get-MgContext
        Send-MgUserMail -UserId $ctx.Account -BodyParameter @{
            Message = @{
                Subject      = "Bem-vindo(a) a YEB - Credenciais de acesso"
                Body         = @{ ContentType = "HTML"; Content = $corpo }
                ToRecipients = @(@{ EmailAddress = @{ Address = $eContato } })
            }
            SaveToSentItems = $false
        } -ErrorAction Stop
        Write-Log "Boas-vindas enviado para $eContato." 'SUCESSO'
    } catch { Write-Log "Erro e-mail boas-vindas: $($_.Exception.Message)" 'AVISO' }
    Write-Log "ACAO: Onboarding completo - $upn - $nExib" 'AUDITORIA'
    Write-Host ""; Show-Separador
    Write-Host "  >>> ONBOARDING CONCLUIDO! <<<" -ForegroundColor Green; Write-Host ""
    Write-Host "  +--------------------------------------------------+" -ForegroundColor Green
    Write-Host "  |  E-mail : $upn" -ForegroundColor Green
    Write-Host "  |  Senha  : $($Script:CONFIG.SenhaPadrao) (trocar no 1 acesso) |" -ForegroundColor Green
    Write-Host "  +--------------------------------------------------+" -ForegroundColor Green
    Write-Host ""; PauseMenu
}

# =============================================================
#  SECAO: RELATORIOS
# =============================================================
function Get-RelatorioAuditoria {
    Show-Banner; Show-Titulo "RELATORIO DE AUDITORIA"
    $logPath = Join-Path $PSScriptRoot 'logs\auditoria.log'
    if (-not (Test-Path $logPath)) { Write-Log "Sem registros ainda." 'AVISO'; PauseMenu; return }
    Write-Host "  [1] Hoje  [2] 7 dias  [3] 30 dias  [4] Tudo: " -NoNewline -ForegroundColor White
    $f = Read-Host
    $corte = switch ($f.Trim()) {
        '1' { (Get-Date).Date }
        '2' { (Get-Date).AddDays(-7) }
        '3' { (Get-Date).AddDays(-30) }
        default { [DateTime]::MinValue }
    }
    $linhas = Get-Content $logPath -Encoding UTF8 | Where-Object {
        try { ([DateTime]::Parse($_.Substring(0,19))) -ge $corte } catch { $false }
    }
    if ($linhas.Count -eq 0) { Write-Log "Sem registros no periodo." 'AVISO'; PauseMenu; return }
    Write-Host ""; Show-Separador
    Write-Host "  AUDITORIA ($($linhas.Count) entradas):" -ForegroundColor Cyan; Show-Separador; Write-Host ""
    $linhas | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
    Write-Host ""; Write-Host "  Exportar CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -match '^[Ss]$') {
        $csv = $linhas | ForEach-Object {
            if ($_ -match '(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}) \[([^\]]+)\] \[([^\]]+)\] (.+)') {
                [PSCustomObject]@{ DataHora = $Matches[1]; Nivel = $Matches[2]; Operador = $Matches[3]; Mensagem = $Matches[4] }
            }
        } | Where-Object { $null -ne $_ }
        Export-ParaCSV -Dados $csv -NomeArquivo 'Auditoria'
    }
    PauseMenu
}

function Get-RelatorioUsuariosRecentes {
    Show-Banner; Show-Titulo "USUARIOS CRIADOS RECENTEMENTE"
    Write-Host "  Criados nos ultimos quantos dias? " -NoNewline -ForegroundColor White
    $ds = Read-Host; $dias = 30
    [int]::TryParse($ds, [ref]$dias) | Out-Null
    if ($dias -le 0) { $dias = 30 }
    $corte = (Get-Date).AddDays(-$dias).ToUniversalTime()
    Write-Host ""; Write-Host "  Buscando..." -ForegroundColor Gray
    try {
        $res = Get-MgUser -All `
            -Property 'Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses,CreatedDateTime' `
            -Filter "createdDateTime ge $($corte.ToString('yyyy-MM-ddTHH:mm:ssZ'))" `
            -ErrorAction Stop |
            Where-Object { $_.UserPrincipalName -notlike '*#EXT#*' } |
            Sort-Object LastContentModifiedDate -Descending
        if (@($res).Count -eq 0) { Write-Log "Nenhum no periodo." 'INFO'; PauseMenu; return }
        Write-Host ""; Show-Separador
        Write-Host ("  {0,-38} {1,-10} {2}" -f "E-MAIL", "STATUS", "CRIADO EM") -ForegroundColor Cyan; Show-Separador
        $csv = @()
        foreach ($u in $res) {
            $st = if ($u.AccountEnabled) { "Ativo" } else { "Bloqueado" }
            $dt = $u.CreatedDateTime.ToLocalTime().ToString('dd/MM/yyyy HH:mm')
            Write-Host ("  {0,-38} {1,-10} {2}" -f $u.UserPrincipalName, $st, $dt) -ForegroundColor $(if ($u.AccountEnabled) { 'White' } else { 'Yellow' })
            $csv += [PSCustomObject]@{ Nome = $u.DisplayName; Email = $u.UserPrincipalName; Status = $st; CriadoEm = $dt; Licencas = @($u.AssignedLicenses).Count }
        }
        Show-Separador; Write-Host "  Total: $(@($res).Count)" -ForegroundColor DarkGray; Write-Host ""
        Write-Host "  Exportar CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -match '^[Ss]$') { Export-ParaCSV -Dados $csv -NomeArquivo "UsuariosCriados_${dias}dias" }
    } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    PauseMenu
}

# =============================================================
#  MENUS
# =============================================================
function Show-Header {
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if ($null -ne $ctx -and $null -ne $ctx.Account) {
        Write-Host "  Conectado : " -NoNewline -ForegroundColor DarkGray
        Write-Host $ctx.Account -ForegroundColor Green
        Write-Host ""
    }
}

function Start-MenuUsuarios {
    $loop = $true
    while ($loop) {
        Show-Banner; Show-Header
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |         GESTAO DE USUARIOS               |" -ForegroundColor Cyan
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |  [1]  Criar novo usuario                 |" -ForegroundColor White
        Write-Host "  |  [2]  Redefinir senha                    |" -ForegroundColor White
        Write-Host "  |  [3]  Bloquear acesso                    |" -ForegroundColor White
        Write-Host "  |  [4]  Desbloquear acesso                 |" -ForegroundColor White
        Write-Host "  |  [5]  Listar usuarios (filtros)          |" -ForegroundColor White
        Write-Host "  |  [6]  Alterar nome ou UPN                |" -ForegroundColor White
        Write-Host "  |  [7]  Detalhes completos do usuario      |" -ForegroundColor White
        Write-Host "  |  [8]  Excluir usuario                    |" -ForegroundColor White
        Write-Host "  |                                          |" -ForegroundColor Cyan
        Write-Host "  |  [0]  Voltar                             |" -ForegroundColor DarkGray
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host ""; Write-Host "  Escolha: " -NoNewline -ForegroundColor White
        switch ((Read-Host).Trim()) {
            '1' { New-UsuarioO365 }
            '2' { Set-RedefinirSenha }
            '3' { Set-BloquearUsuario }
            '4' { Set-DesbloquearUsuario }
            '5' { Get-ListarUsuarios }
            '6' { Set-AlterarUsuario }
            '7' { Get-DetalhesUsuario }
            '8' { Remove-UsuarioO365 }
            '0' { $loop = $false }
            default { Write-Host "  >> Invalido." -ForegroundColor Yellow; Start-Sleep -Seconds 1 }
        }
    }
}

function Start-MenuLicencas {
    $loop = $true
    while ($loop) {
        Show-Banner; Show-Header
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |         GESTAO DE LICENCAS               |" -ForegroundColor Cyan
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |  [1]  Atribuir licenca a usuario         |" -ForegroundColor White
        Write-Host "  |  [2]  Retirar licencas                   |" -ForegroundColor White
        Write-Host "  |  [3]  Relatorio de licencas do tenant    |" -ForegroundColor White
        Write-Host "  |  [4]  Usuarios sem licenca               |" -ForegroundColor White
        Write-Host "  |                                          |" -ForegroundColor Cyan
        Write-Host "  |  [0]  Voltar                             |" -ForegroundColor DarkGray
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host ""; Write-Host "  Escolha: " -NoNewline -ForegroundColor White
        switch ((Read-Host).Trim()) {
            '1' { Add-LicencaUsuario }
            '2' { Remove-LicencasUsuario }
            '3' { Get-RelatorioLicencas }
            '4' { Get-UsuariosSemLicenca }
            '0' { $loop = $false }
            default { Write-Host "  >> Invalido." -ForegroundColor Yellow; Start-Sleep -Seconds 1 }
        }
    }
}

function Start-MenuGrupos {
    $loop = $true
    while ($loop) {
        Show-Banner; Show-Header
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |         GESTAO DE GRUPOS                 |" -ForegroundColor Cyan
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |  [1]  Adicionar usuario a grupo          |" -ForegroundColor White
        Write-Host "  |  [2]  Remover usuario de grupo           |" -ForegroundColor White
        Write-Host "  |  [3]  Listar grupos e membros            |" -ForegroundColor White
        Write-Host "  |                                          |" -ForegroundColor Cyan
        Write-Host "  |  [0]  Voltar                             |" -ForegroundColor DarkGray
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host ""; Write-Host "  Escolha: " -NoNewline -ForegroundColor White
        switch ((Read-Host).Trim()) {
            '1' { Add-MembroGrupo }
            '2' { Remove-MembroGrupo }
            '3' { Get-ListarGrupos }
            '0' { $loop = $false }
            default { Write-Host "  >> Invalido." -ForegroundColor Yellow; Start-Sleep -Seconds 1 }
        }
    }
}

function Start-MenuSeguranca {
    $loop = $true
    while ($loop) {
        Show-Banner; Show-Header
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |      SEGURANCA E CONFORMIDADE            |" -ForegroundColor Cyan
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |  [1]  Revogar sessoes ativas             |" -ForegroundColor White
        Write-Host "  |  [2]  Forcar MFA                         |" -ForegroundColor White
        Write-Host "  |  [3]  Usuarios inativos                  |" -ForegroundColor White
        Write-Host "  |  [4]  Usuarios sem MFA                   |" -ForegroundColor White
        Write-Host "  |                                          |" -ForegroundColor Cyan
        Write-Host "  |  [0]  Voltar                             |" -ForegroundColor DarkGray
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host ""; Write-Host "  Escolha: " -NoNewline -ForegroundColor White
        switch ((Read-Host).Trim()) {
            '1' { Revoke-SessoesUsuario }
            '2' { Set-ForcarMFA }
            '3' { Get-UsuariosInativos }
            '4' { Get-UsuariosSemMFA }
            '0' { $loop = $false }
            default { Write-Host "  >> Invalido." -ForegroundColor Yellow; Start-Sleep -Seconds 1 }
        }
    }
}

function Start-MenuExchange {
    $loop = $true
    while ($loop) {
        Show-Banner; Show-Header
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |      CAIXAS DE E-MAIL E RECURSOS         |" -ForegroundColor Cyan
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |  [1]  Criar caixa compartilhada          |" -ForegroundColor White
        Write-Host "  |  [2]  Permissao em caixa compartilhada   |" -ForegroundColor White
        Write-Host "  |  [3]  Criar sala ou equipamento          |" -ForegroundColor White
        Write-Host "  |  [4]  Configurar redirecionamento        |" -ForegroundColor White
        Write-Host "  |                                          |" -ForegroundColor Cyan
        Write-Host "  |  [0]  Voltar                             |" -ForegroundColor DarkGray
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host ""; Write-Host "  Escolha: " -NoNewline -ForegroundColor White
        switch ((Read-Host).Trim()) {
            '1' { New-CaixaCompartilhada }
            '2' { Add-PermissaoCaixaCompartilhada }
            '3' { New-SalaOuEquipamento }
            '4' { Set-RedirecionamentoEmail }
            '0' { $loop = $false }
            default { Write-Host "  >> Invalido." -ForegroundColor Yellow; Start-Sleep -Seconds 1 }
        }
    }
}

function Start-MenuRelatorios {
    $loop = $true
    while ($loop) {
        Show-Banner; Show-Header
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |      RELATORIOS E EXPORTACAO             |" -ForegroundColor Cyan
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |  [1]  Auditoria do script                |" -ForegroundColor White
        Write-Host "  |  [2]  Usuarios criados recentemente      |" -ForegroundColor White
        Write-Host "  |  [3]  Relatorio de licencas              |" -ForegroundColor White
        Write-Host "  |  [4]  Usuarios sem licenca               |" -ForegroundColor White
        Write-Host "  |  [5]  Usuarios inativos                  |" -ForegroundColor White
        Write-Host "  |                                          |" -ForegroundColor Cyan
        Write-Host "  |  [0]  Voltar                             |" -ForegroundColor DarkGray
        Write-Host "  +------------------------------------------+" -ForegroundColor Cyan
        Write-Host ""; Write-Host "  Escolha: " -NoNewline -ForegroundColor White
        switch ((Read-Host).Trim()) {
            '1' { Get-RelatorioAuditoria }
            '2' { Get-RelatorioUsuariosRecentes }
            '3' { Get-RelatorioLicencas }
            '4' { Get-UsuariosSemLicenca }
            '5' { Get-UsuariosInativos }
            '0' { $loop = $false }
            default { Write-Host "  >> Invalido." -ForegroundColor Yellow; Start-Sleep -Seconds 1 }
        }
    }
}


# =============================================================
#  SECAO: SHAREPOINT ONLINE
# =============================================================
$Script:SPOConectado = $false
$Script:SPOAdminUrl   = $null
$Script:SPOUsarGraph  = $false
$Script:SPOStorageProp = $null

function Get-SPOAdminUrl {
    # Descobre a URL admin do SharePoint via Graph
    try {
        $dominios = Get-MgDomain -ErrorAction Stop |
                    Where-Object { $_.IsVerified -eq $true -and $_.Id -like '*.onmicrosoft.com' }
        $tenantDomain = ($dominios | Select-Object -First 1).Id -replace '\.onmicrosoft\.com$', ''
        return "https://${tenantDomain}-admin.sharepoint.com"
    } catch {
        Write-Log "Nao foi possivel determinar a URL admin: $($_.Exception.Message)" 'ERRO'
        return $null
    }
}



function Connect-SharePointOnline {
    if ($Script:SPOConectado) { return $true }

    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if ($null -eq $ctx -or $null -eq $ctx.Account) {
        Write-Log "Conecte ao Microsoft 365 primeiro (opcao de login)." 'ERRO'
        return $false
    }

    # Tenta modulo SPO nativo (PS 5.1)
    $moduloSPO = Get-Module -Name 'Microsoft.Online.SharePoint.PowerShell' -ListAvailable |
                 Sort-Object Version -Descending | Select-Object -First 1

    if ($null -ne $moduloSPO) {
        try {
            Import-Module 'Microsoft.Online.SharePoint.PowerShell' -DisableNameChecking -Force -ErrorAction Stop
        } catch { }
    }

    # Instala se necessario
    if ($null -eq (Get-Command Connect-SPOService -ErrorAction SilentlyContinue)) {
        Write-Host "  Instalando Microsoft.Online.SharePoint.PowerShell..." -ForegroundColor Gray
        $escopo = if (Test-IsAdministrador) { 'AllUsers' } else { 'CurrentUser' }
        try {
            Install-Module 'Microsoft.Online.SharePoint.PowerShell' -Scope $escopo -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Import-Module 'Microsoft.Online.SharePoint.PowerShell' -DisableNameChecking -Force -ErrorAction Stop
        } catch {
            Write-Log "Falha ao instalar modulo SPO: $($_.Exception.Message)" 'AVISO'
        }
    }

    $adminUrl = Get-SPOAdminUrl
    if ($null -eq $adminUrl) { return $false }

    Write-Host "  Admin URL : $adminUrl" -ForegroundColor DarkGray
    Write-Host "  Conta     : $($ctx.Account)" -ForegroundColor DarkGray

    # METODO 1: Connect-SPOService com token do MgGraph (moderno, sem senha)
    if ($null -ne (Get-Command Connect-SPOService -ErrorAction SilentlyContinue)) {
        try {
            Write-Host "  Conectando via token de acesso atual..." -ForegroundColor Gray
            # Obtem token do MgGraph para reutilizar
            # Conecta via autenticacao moderna
            Connect-SPOService -Url $adminUrl -Region Default -ErrorAction Stop
            Get-SPOTenant -ErrorAction Stop | Out-Null
            $Script:SPOConectado = $true
            $Script:SPOAdminUrl  = $adminUrl
            Write-Log "SharePoint Online conectado via SPO Service: $adminUrl" 'SUCESSO'
            return $true
        } catch {
            Write-Log "Metodo 1 falhou: $($_.Exception.Message)" 'AVISO'
        }

        # METODO 2: Connect-SPOService com ModernAuth explicito
        try {
            Write-Host "  Tentando autenticacao moderna..." -ForegroundColor Gray
            Connect-SPOService -Url $adminUrl -ModernAuth $true -ErrorAction Stop
            Get-SPOTenant -ErrorAction Stop | Out-Null
            $Script:SPOConectado = $true
            $Script:SPOAdminUrl  = $adminUrl
            Write-Log "SharePoint Online conectado (ModernAuth): $adminUrl" 'SUCESSO'
            return $true
        } catch {
            Write-Log "Metodo 2 falhou: $($_.Exception.Message)" 'AVISO'
        }
    }

    # METODO 3: Graph REST API direto (sem modulo SPO, usa token do MgGraph)
    Write-Host "  Usando Microsoft Graph REST API diretamente..." -ForegroundColor Gray
    try {
        $teste = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites?search=*" -OutputType PSObject -ErrorAction Stop
        if ($null -ne $teste) {
            $Script:SPOConectado = $true
            $Script:SPOAdminUrl  = $adminUrl
            $Script:SPOUsarGraph = $true
            Write-Log "SharePoint conectado via Microsoft Graph REST API." 'SUCESSO'
            return $true
        }
    } catch {
        Write-Log "Metodo Graph falhou: $($_.Exception.Message)" 'ERRO'
    }

    Write-Log "Todos os metodos de conexao falharam." 'ERRO'
    Write-Host "  Tente reconectar ao M365 escolhendo a opcao [0] Sair e executando novamente." -ForegroundColor Yellow
    return $false
}

function Get-SPOTenantInfo {
    # Wrapper que funciona com modulo SPO ou Graph
    if ($Script:SPOUsarGraph) {
        try {
            # Busca via Graph
            $r = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/admin/sharepoint/settings" -OutputType PSObject -ErrorAction Stop
            return $r
        } catch {
            # Fallback: busca dados do tenant via sites
            return $null
        }
    }
    return Get-SPOTenant -ErrorAction Stop
}



function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return "$([math]::Round($Bytes/1TB, 2)) TB" }
    if ($Bytes -ge 1GB) { return "$([math]::Round($Bytes/1GB, 2)) GB" }
    if ($Bytes -ge 1MB) { return "$([math]::Round($Bytes/1MB, 2)) MB" }
    if ($Bytes -ge 1KB) { return "$([math]::Round($Bytes/1KB, 2)) KB" }
    return "$Bytes B"
}

function Show-BarraUso {
    param([double]$Porcentagem, [int]$Largura = 30)
    # Protege contra valores invalidos (negativo, acima de 100, NaN)
    if ($Porcentagem -lt 0 -or [double]::IsNaN($Porcentagem) -or [double]::IsInfinity($Porcentagem)) { $Porcentagem = 0 }
    if ($Porcentagem -gt 100) { $Porcentagem = 100 }
    $preenchido = [int][math]::Floor($Porcentagem / 100 * $Largura)
    $vazio      = $Largura - $preenchido
    # Garante que preenchido e vazio nunca sejam negativos
    if ($preenchido -lt 0) { $preenchido = 0 }
    if ($vazio      -lt 0) { $vazio      = 0 }
    $barra = '[' + ('#' * $preenchido) + ('-' * $vazio) + ']'
    $cor   = if ($Porcentagem -ge 90) { 'Red' } elseif ($Porcentagem -ge 70) { 'Yellow' } else { 'Green' }
    Write-Host "  $barra $([math]::Round($Porcentagem,1))%" -ForegroundColor $cor
}

# -- SPO 1: Visao geral do tenant SharePoint --

function Get-SPOSiteStorage {
    param([object]$Site)
    $mb = 0
    try {
        # Descobre qual propriedade de storage existe neste objeto
        $props = $Site | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
        if ('StorageUsage'        -in $props -and $Site.StorageUsage        -gt 0) { return [long]$Site.StorageUsage }
        if ('StorageUsageCurrent' -in $props -and $Site.StorageUsageCurrent -gt 0) { return [long]$Site.StorageUsageCurrent }
        if ('StorageMaximumLevel' -in $props -and $Site.StorageMaximumLevel -gt 0) {
            # StorageMaximumLevel eh a quota, nao o uso - ignora
        }
        # Nenhuma propriedade util encontrada no objeto bulk
        # Busca individualmente (mais lento mas confiavel)
        $det   = Get-SPOSite -Identity $Site.Url -ErrorAction SilentlyContinue
        if ($null -ne $det) {
            $dProps = $det | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
            if ('StorageUsage'        -in $dProps -and $det.StorageUsage        -gt 0) { return [long]$det.StorageUsage }
            if ('StorageUsageCurrent' -in $dProps -and $det.StorageUsageCurrent -gt 0) { return [long]$det.StorageUsageCurrent }
        }
    } catch { }
    return [long]$mb
}

# Cache do nome da propriedade de storage (descoberto uma vez, reutilizado)
$Script:SPOStorageProp = $null

function Get-SPOStoragePropName {
    # Retorna o nome correto da propriedade de storage do tenant atual
    if ($null -ne $Script:SPOStorageProp) { return $Script:SPOStorageProp }
    try {
        $site1 = Get-SPOSite -Limit 1 -ErrorAction Stop | Select-Object -First 1
        if ($null -ne $site1) {
            $props = $site1 | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
            if ('StorageUsage'        -in $props) { $Script:SPOStorageProp = 'StorageUsage';        return $Script:SPOStorageProp }
            if ('StorageUsageCurrent' -in $props) { $Script:SPOStorageProp = 'StorageUsageCurrent'; return $Script:SPOStorageProp }
            if ('StorageQuotaWarningLevel' -in $props) {
                # Log quais props existem para diagnostico
                Write-Log "Propriedades de storage disponiveis: $($props -join ', ')" 'INFO'
            }
        }
    } catch { }
    $Script:SPOStorageProp = 'StorageUsage' # fallback
    return $Script:SPOStorageProp
}

function Get-SPOVisaoGeral {
    Show-Banner; Show-Titulo "SHAREPOINT - VISAO GERAL DO TENANT"
    if (-not (Connect-SharePointOnline)) { PauseMenu; return }
    Write-Host "  Buscando informacoes do tenant..." -ForegroundColor Gray
    try {
        $tenant = Get-SPOTenantInfo -ErrorAction Stop
        $sites  = @(Get-SPOSite -Limit All -ErrorAction Stop)

        # Calculo de armazenamento
        $totalGB = if ($null -ne $tenant -and $tenant.StorageQuota -gt 0) { [math]::Round($tenant.StorageQuota / 1024, 2) } else { 0 }
        Write-Host "  Calculando armazenamento..." -ForegroundColor Gray
        $propNome = Get-SPOStoragePropName
        Write-Log "Propriedade storage detectada: $propNome" 'INFO'
        $usadoMB  = 0
        foreach ($__s in $sites) {
            try { $val = $__s.$propNome; if ($val -gt 0) { $usadoMB += [long]$val } } catch { }
        }
        $usadoGB    = [math]::Round($usadoMB / 1024, 2)
        $dispGB     = [math]::Round($totalGB - $usadoGB, 2)
        $porcentagem = 0
        if ($totalGB -gt 0 -and $usadoGB -ge 0) {
            $porcentagem = [math]::Round(($usadoGB / $totalGB) * 100, 1)
            if ($porcentagem -gt 100) { $porcentagem = 100 }
        }

        $sitesAtivos  = @($sites | Where-Object { $_.Status -eq 'Active' -or $_.LockState -eq 'Unlock' }).Count
        $sitesBloq    = @($sites | Where-Object { $_.LockState -ne 'Unlock' }).Count
        $sitesTeams   = @($sites | Where-Object { $_.Template -like 'TEAMCHANNEL*' -or $_.Template -like 'GROUP*' }).Count
        $sitesComm    = @($sites | Where-Object { $_.Template -like 'SITEPAGEPUBLISHING*' -or $_.Template -like 'COMMUNICATION*' }).Count
        $sitesPess    = @($sites | Where-Object { $_.Template -like 'SPSPERS*' }).Count

        Write-Host ""
        Write-Host "  ======================================================" -ForegroundColor Cyan
        Write-Host "  ARMAZENAMENTO DO TENANT" -ForegroundColor Cyan
        Write-Host "  ======================================================" -ForegroundColor Cyan
        Write-Host ("  Total disponivel : {0,10}" -f (Format-Bytes ($totalGB * 1GB))) -ForegroundColor White
        Write-Host ("  Espaco usado     : {0,10}" -f (Format-Bytes ($usadoGB * 1GB))) -ForegroundColor White
        Write-Host ("  Espaco livre     : {0,10}" -f (Format-Bytes ($dispGB * 1GB))) -ForegroundColor $(if ($dispGB -lt 50) { 'Red' } elseif ($dispGB -lt 200) { 'Yellow' } else { 'Green' })
        Write-Host ""
        Write-Host "  Uso do armazenamento:" -ForegroundColor Gray
        Show-BarraUso -Porcentagem $porcentagem
        Write-Host ""
        Write-Host "  ======================================================" -ForegroundColor Cyan
        Write-Host "  SITES" -ForegroundColor Cyan
        Write-Host "  ======================================================" -ForegroundColor Cyan
        Write-Host ("  Total de sites   : {0}" -f @($sites).Count)        -ForegroundColor White
        Write-Host ("  Sites ativos     : {0}" -f $sitesAtivos)         -ForegroundColor Green
        Write-Host ("  Sites bloqueados : {0}" -f $sitesBloq)           -ForegroundColor $(if ($sitesBloq -gt 0) { 'Yellow' } else { 'White' })
        Write-Host ("  Teams/Grupos     : {0}" -f $sitesTeams)          -ForegroundColor White
        Write-Host ("  Comunicacao      : {0}" -f $sitesComm)           -ForegroundColor White
        Write-Host ("  OneDrive pessoal : {0}" -f $sitesPess)           -ForegroundColor White
        Write-Host ""
        Write-Host "  ======================================================" -ForegroundColor Cyan
        Write-Host "  CONFIGURACOES DO TENANT" -ForegroundColor Cyan
        Write-Host "  ======================================================" -ForegroundColor Cyan
        Write-Host ("  Compartilhamento externo: {0}" -f $tenant.SharingCapability) -ForegroundColor White
        Write-Host ("  Quota padrao por site   : {0}" -f (Format-Bytes ([long]$tenant.OneDriveStorageQuota * 1GB))) -ForegroundColor White
        Write-Host ""
    } catch {
        Write-Log "Erro ao buscar informacoes: $($_.Exception.Message)" 'ERRO'
    }
    PauseMenu
}

# -- SPO 2: Listar todos os sites com uso de armazenamento --
function Get-SPOListarSites {
    Show-Banner; Show-Titulo "SHAREPOINT - LISTAR TODOS OS SITES"
    if (-not (Connect-SharePointOnline)) { PauseMenu; return }

    Write-Host "  Ordenar por:" -ForegroundColor Cyan
    Write-Host "  [1] Uso de armazenamento (maior primeiro)" -ForegroundColor White
    Write-Host "  [2] Nome do site (A-Z)"                   -ForegroundColor White
    Write-Host "  [3] Ultima modificacao (mais recente)"       -ForegroundColor White
    Write-Host "  [4] Somente sites bloqueados"             -ForegroundColor White
    Write-Host ""
    Write-Host "  Escolha: " -NoNewline -ForegroundColor White
    $ord = Read-Host

    Write-Host ""; Write-Host "  Buscando sites..." -ForegroundColor Gray
    try {
        $sites = @(Get-SPOSite -Limit All -ErrorAction Stop)

        switch ($ord.Trim()) {
            '2' { $sites = $sites | Sort-Object Title }
            '3' { $sites = $sites | Where-Object { $null -ne $_.LastContentModifiedDate } | Sort-Object LastContentModifiedDate -Descending }
            '4' { $sites = $sites | Where-Object { $_.LockState -ne 'Unlock' } }
            default { $sites = $sites | Where-Object { $null -ne $_.StorageUsageCurrent } | Sort-Object StorageUsageCurrent -Descending }
        }

        if (@($sites).Count -eq 0) { Write-Log "Nenhum site encontrado." 'AVISO'; PauseMenu; return }

        Write-Host ""
        Show-Separador
        Write-Host ("  {0,-45} {1,12} {2,12} {3,-12}" -f "SITE", "USADO", "QUOTA", "STATUS") -ForegroundColor Cyan
        Show-Separador

        $csv = @()
        foreach ($s in $sites) {
            $usadoMB = Get-SPOSiteStorage -Site $s
            $quotaMB = $s.StorageQuota
            $usadoStr = Format-Bytes ($usadoMB * 1MB)
            $quotaStr = if ($quotaMB -gt 0) { Format-Bytes ($quotaMB * 1MB) } else { "Herdado" }
            $status   = $s.LockState
            $cor      = if ($status -ne 'Unlock') { 'Yellow' } else { 'White' }
            $tNome    = if ($s.Title) { $s.Title } else { $s.Url }
            $titulo   = if ($tNome.Length -gt 43) { $tNome.Substring(0,43) + '..' } else { $tNome }
            Write-Host ("  {0,-45} {1,12} {2,12} {3}" -f $titulo, $usadoStr, $quotaStr, $status) -ForegroundColor $cor
            $csv += [PSCustomObject]@{
                Nome       = $s.Title
                URL        = $s.Url
                UsadoMB    = [long]$usadoMB
                QuotaMB    = $quotaMB
                Status     = $status
                Template   = $s.Template
                Proprietario = $s.Owner
                Criado     = if ($s.LastContentModifiedDate) { $s.LastContentModifiedDate } else { '' }
            }
        }
        Show-Separador
        Write-Host "  Total: $(@($sites).Count) site(s)" -ForegroundColor DarkGray
        Write-Host ""; Write-Host "  Exportar CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -match '^[Ss]$') { Export-ParaCSV -Dados $csv -NomeArquivo 'SPO_Sites' }
    } catch {
        Write-Log "Erro: $($_.Exception.Message)" 'ERRO'
    }
    PauseMenu
}

# -- SPO 3: Quem usa mais armazenamento --
function Get-SPOTopConsumidores {
    Show-Banner; Show-Titulo "SHAREPOINT - MAIORES CONSUMIDORES DE ESPACO"
    if (-not (Connect-SharePointOnline)) { PauseMenu; return }

    Write-Host "  Mostrar top quantos sites? [ENTER = 20]: " -NoNewline -ForegroundColor White
    $topStr = Read-Host; $top = 20
    [int]::TryParse($topStr, [ref]$top) | Out-Null
    if ($top -le 0) { $top = 20 }

    Write-Host ""; Write-Host "  Buscando e calculando uso..." -ForegroundColor Gray
    try {
        $sites = @(Get-SPOSite -Limit All -ErrorAction Stop) | Select-Object -First $top
        $totalMB = ($sites | Measure-Object -Property StorageMB -Sum).Sum

        Write-Host ""
        Write-Host "  TOP $top - SITES POR USO DE ARMAZENAMENTO" -ForegroundColor Cyan
        Show-Separador
        Write-Host ("  {0,-4} {1,-40} {2,12} {3,8}" -f "#", "SITE", "USADO", "% DO TOP") -ForegroundColor Cyan
        Show-Separador

        $csv = @(); $pos = 1
        foreach ($s in $sites) {
            $usadoStr = Format-Bytes ($s.StorageMB * 1MB)
            $pct = if ($totalMB -gt 0) { [math]::Min(100, [math]::Round($s.StorageMB / $totalMB * 100, 1)) } else { 0 }
            $cor      = if ($pos -le 3) { 'Red' } elseif ($pos -le 10) { 'Yellow' } else { 'White' }
            $tNome2   = if ($s.Title) { $s.Title } else { $s.Url }
            $titulo   = if ($tNome2.Length -gt 38) { $tNome2.Substring(0,38) + '..' } else { $tNome2 }
            Write-Host ("  {0,-4} {1,-40} {2,12} {3,7}%" -f $pos, $titulo, $usadoStr, $pct) -ForegroundColor $cor
            $csv += [PSCustomObject]@{
                Posicao  = $pos
                Site     = $s.Title
                URL      = $s.Url
                UsadoMB  = $s.StorageMB
                Usado    = $usadoStr
                PctDoTop = "$pct%"
                Owner    = $s.Owner
            }
            $pos++
        }
        Show-Separador
        Write-Host ("  Total dos {0} sites: {1}" -f $top, (Format-Bytes ($totalMB * 1MB))) -ForegroundColor DarkGray
        Write-Host ""; Write-Host "  Exportar CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -match '^[Ss]$') { Export-ParaCSV -Dados $csv -NomeArquivo "SPO_TopConsumidores_Top$top" }
    } catch {
        Write-Log "Erro: $($_.Exception.Message)" 'ERRO'
    }
    PauseMenu
}

# -- SPO 4: Detalhes de um site especifico --
function Get-SPODetalhesSite {
    Show-Banner; Show-Titulo "SHAREPOINT - DETALHES DO SITE"
    if (-not (Connect-SharePointOnline)) { PauseMenu; return }

    Write-Host "  URL do site (ex: https://empresa.sharepoint.com/sites/nome): " -NoNewline -ForegroundColor White
    $url = Read-Host
    if ([string]::IsNullOrWhiteSpace($url)) { Write-Log "URL invalida." 'AVISO'; PauseMenu; return }

    Write-Host ""; Write-Host "  Buscando detalhes..." -ForegroundColor Gray
    try {
        $s = Get-SPOSite -Identity $url -ErrorAction Stop

        $usadoMB  = (Get-SPOSiteStorage -Site $s)
        $quotaMB  = $s.StorageQuota
        $pct      = if ($quotaMB -gt 0) { [math]::Round($usadoMB / $quotaMB * 100, 1) } else { 0 }

        Write-Host ""; Show-Separador
        Write-Host "  INFORMACOES DO SITE" -ForegroundColor Cyan
        Show-Separador
        Write-Host ("  Titulo       : {0}" -f $s.Title)                              -ForegroundColor White
        Write-Host ("  URL          : {0}" -f $s.Url)                                -ForegroundColor White
        Write-Host ("  Template     : {0}" -f $s.Template)                           -ForegroundColor White
        Write-Host ("  Proprietario : {0}" -f $s.Owner)                              -ForegroundColor White
        Write-Host ("  Status       : {0}" -f $s.LockState)                          -ForegroundColor $(if ($s.LockState -ne 'Unlock') { 'Yellow' } else { 'Green' })
        Write-Host ("  Criado em    : N/D (use portal admin)") -ForegroundColor DarkGray
        $ultimoMod = if ($s.LastContentModifiedDate) { $s.LastContentModifiedDate.ToString('dd/MM/yyyy') } else { '-' }
        Write-Host ("  Ultimo acesso: {0}" -f $ultimoMod) -ForegroundColor White
        Show-Separador
        Write-Host "  ARMAZENAMENTO" -ForegroundColor Cyan
        Show-Separador
        Write-Host ("  Usado        : {0}" -f (Format-Bytes ($usadoMB * 1MB)))       -ForegroundColor White
        Write-Host ("  Quota do site: {0}" -f (if ($quotaMB -gt 0) { Format-Bytes ($quotaMB * 1MB) } else { "Herdado do tenant" })) -ForegroundColor White
        if ($quotaMB -gt 0) {
            Write-Host "  Uso:" -ForegroundColor Gray
            Show-BarraUso -Porcentagem $pct
        }
        Show-Separador
        Write-Host "  COMPARTILHAMENTO" -ForegroundColor Cyan
        Show-Separador
        Write-Host ("  Politica     : {0}" -f $s.SharingCapability)                  -ForegroundColor White
        Write-Host ("  Membros      : {0}" -f $s.MembersCanShare)                    -ForegroundColor White
        Show-Separador
        Write-Host ""
    } catch {
        Write-Log "Erro ao buscar site: $($_.Exception.Message)" 'ERRO'
    }
    PauseMenu
}

# -- SPO 5: Gerenciar quota de um site --
function Set-SPOQuotaSite {
    Show-Banner; Show-Titulo "SHAREPOINT - ALTERAR QUOTA DO SITE"
    if (-not (Connect-SharePointOnline)) { PauseMenu; return }

    Write-Host "  URL do site: " -NoNewline -ForegroundColor White
    $url = Read-Host
    if ([string]::IsNullOrWhiteSpace($url)) { Write-Log "URL invalida." 'AVISO'; PauseMenu; return }

    Write-Host ""; Write-Host "  Buscando site..." -ForegroundColor Gray
    try {
        $s = Get-SPOSite -Identity $url -ErrorAction Stop
    } catch {
        Write-Log "Site nao encontrado: $($_.Exception.Message)" 'ERRO'; PauseMenu; return
    }

    $usadoMB = (Get-SPOSiteStorage -Site $s)
    $quotaMB = $s.StorageQuota

    Write-Host ""; Show-Separador
    Write-Host ("  Site  : {0}" -f $s.Title) -ForegroundColor White
    Write-Host ("  Usado : {0}" -f (Format-Bytes ($usadoMB * 1MB))) -ForegroundColor White
    Write-Host ("  Quota : {0}" -f (if ($quotaMB -gt 0) { Format-Bytes ($quotaMB * 1MB) } else { "Herdado" })) -ForegroundColor White
    Show-Separador; Write-Host ""
    Write-Host "  Nova quota em GB (minimo: $([math]::Ceiling($usadoMB/1024) + 1) GB): " -NoNewline -ForegroundColor White
    $novaGBStr = Read-Host; $novaGB = 0
    if (-not [int]::TryParse($novaGBStr, [ref]$novaGB) -or $novaGB -le 0) {
        Write-Log "Valor invalido." 'AVISO'; PauseMenu; return
    }
    $novaMB = $novaGB * 1024
    $minimoMB = [math]::Ceiling($usadoMB) + 100
    if ($novaMB -lt $minimoMB) {
        Write-Log "Quota minima para este site: $([math]::Ceiling($minimoMB/1024)) GB (tem $([math]::Round($usadoMB/1024,2)) GB em uso)." 'ERRO'
        PauseMenu; return
    }
    Write-Host ""; Write-Host "  Definir quota de $novaGB GB para '$($s.Title)'? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }

    try {
        Set-SPOSite -Identity $url -StorageQuota $novaMB -StorageQuotaWarningLevel ($novaMB * 0.9) -ErrorAction Stop
        Write-Log "ACAO: Quota alterada - $url - $novaGB GB" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> QUOTA ATUALIZADA PARA $novaGB GB! <<<" -ForegroundColor Green
    } catch {
        Write-Log "Erro ao alterar quota: $($_.Exception.Message)" 'ERRO'
    }
    PauseMenu
}

# -- SPO 6: Bloquear/desbloquear site --
function Set-SPOBloqueioSite {
    Show-Banner; Show-Titulo "SHAREPOINT - BLOQUEAR / DESBLOQUEAR SITE"
    if (-not (Connect-SharePointOnline)) { PauseMenu; return }

    Write-Host "  URL do site: " -NoNewline -ForegroundColor White
    $url = Read-Host
    if ([string]::IsNullOrWhiteSpace($url)) { Write-Log "URL invalida." 'AVISO'; PauseMenu; return }

    Write-Host ""; Write-Host "  Buscando site..." -ForegroundColor Gray
    try { $s = Get-SPOSite -Identity $url -ErrorAction Stop }
    catch { Write-Log "Site nao encontrado." 'ERRO'; PauseMenu; return }

    $bloqueado = $s.LockState -ne 'Unlock'
    Write-Host ""; Show-Separador
    Write-Host ("  Site   : {0}" -f $s.Title) -ForegroundColor White
    Write-Host ("  Status : {0}" -f $s.LockState) -ForegroundColor $(if ($bloqueado) { 'Red' } else { 'Green' })
    Show-Separador; Write-Host ""

    if ($bloqueado) {
        Write-Host "  Site BLOQUEADO. Desbloquear? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
        try {
            Set-SPOSite -Identity $url -LockState Unlock -ErrorAction Stop
            Write-Log "ACAO: Site desbloqueado - $url" 'AUDITORIA'
            Write-Host ""; Write-Host "  >>> SITE DESBLOQUEADO! <<<" -ForegroundColor Green
        } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    } else {
        Write-Host "  Tipo de bloqueio:" -ForegroundColor Cyan
        Write-Host "  [1] ReadOnly  - somente leitura (recomendado)"  -ForegroundColor White
        Write-Host "  [2] NoAccess  - sem acesso algum"               -ForegroundColor White
        Write-Host ""; Write-Host "  Escolha: " -NoNewline -ForegroundColor White
        $tipo = Read-Host
        $lock = switch ($tipo.Trim()) { '2' { 'NoAccess' } default { 'ReadOnly' } }
        Write-Host ""; Write-Host "  Bloquear '$($s.Title)' como '$lock'? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
        try {
            Set-SPOSite -Identity $url -LockState $lock -ErrorAction Stop
            Write-Log "ACAO: Site bloqueado ($lock) - $url" 'AUDITORIA'
            Write-Host ""; Write-Host "  >>> SITE BLOQUEADO ($lock)! <<<" -ForegroundColor Yellow
        } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
    }
    PauseMenu
}

# -- SPO 7: Gerenciar proprietarios de um site --
function Set-SPOProprietarioSite {
    Show-Banner; Show-Titulo "SHAREPOINT - GERENCIAR PROPRIETARIO DO SITE"
    if (-not (Connect-SharePointOnline)) { PauseMenu; return }

    Write-Host "  URL do site: " -NoNewline -ForegroundColor White
    $url = Read-Host
    if ([string]::IsNullOrWhiteSpace($url)) { Write-Log "URL invalida." 'AVISO'; PauseMenu; return }

    Write-Host ""; Write-Host "  Buscando site..." -ForegroundColor Gray
    try { $s = Get-SPOSite -Identity $url -ErrorAction Stop }
    catch { Write-Log "Site nao encontrado." 'ERRO'; PauseMenu; return }

    Write-Host ""; Show-Separador
    Write-Host ("  Site              : {0}" -f $s.Title)      -ForegroundColor White
    Write-Host ("  Proprietario atual: {0}" -f $s.Owner)      -ForegroundColor White
    Show-Separador; Write-Host ""

    Write-Host "  Acoes:" -ForegroundColor Cyan
    Write-Host "  [1] Alterar proprietario principal" -ForegroundColor White
    Write-Host "  [2] Ver todos os administradores"   -ForegroundColor White
    Write-Host "  [3] Adicionar administrador"        -ForegroundColor White
    Write-Host "  [4] Remover administrador"          -ForegroundColor White
    Write-Host ""; Write-Host "  Escolha: " -NoNewline -ForegroundColor White
    $op = Read-Host

    switch ($op.Trim()) {
        '1' {
            $novoOwner = Read-EmailObrigatorio "Novo proprietario (e-mail): "
            Write-Host ""; Write-Host "  Confirmar? [S/N]: " -NoNewline -ForegroundColor Yellow
            if (Read-Host -notmatch '^[Ss]$') { PauseMenu; return }
            try {
                Set-SPOSite -Identity $url -Owner $novoOwner -ErrorAction Stop
                Write-Log "ACAO: Proprietario alterado - $url -> $novoOwner" 'AUDITORIA'
                Write-Host ""; Write-Host "  >>> PROPRIETARIO ATUALIZADO! <<<" -ForegroundColor Green
            } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
        }
        '2' {
            try {
                # SPO nao precisa de reconexao por site
                $admins = Get-SPOUser -Site $url -Limit All | Where-Object { $_.IsSiteAdmin } -ErrorAction Stop
                Write-Host ""; Show-Separador; Write-Host "  Administradores:" -ForegroundColor Cyan; Show-Separador
                $admins | ForEach-Object { Write-Host "  - $($_.Email) ($($_.Title))" -ForegroundColor White }
                # reconexao nao necessaria com SPO
            } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
        }
        '3' {
            $novoAdmin = Read-EmailObrigatorio "E-mail do novo admin    : "
            try {
                Set-SPOUser -Site $url -LoginName $novoAdmin -ErrorAction Stop
                Write-Log "ACAO: Admin adicionado - $url -> $novoAdmin" 'AUDITORIA'
                Write-Host ""; Write-Host "  >>> ADMIN ADICIONADO! <<<" -ForegroundColor Green
            } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
        }
        '4' {
            $remAdmin = Read-EmailObrigatorio "E-mail do admin remover : "
            Write-Host ""; Write-Host "  Confirmar remocao? [S/N]: " -NoNewline -ForegroundColor Yellow
            if (Read-Host -notmatch '^[Ss]$') { PauseMenu; return }
            try {
                Set-SPOUser -Site $url -LoginName $remAdmin -ErrorAction Stop
                Write-Log "ACAO: Admin removido - $url -> $remAdmin" 'AUDITORIA'
                Write-Host ""; Write-Host "  >>> ADMIN REMOVIDO! <<<" -ForegroundColor Green
            } catch { Write-Log "Erro: $($_.Exception.Message)" 'ERRO' }
        }
        default { Write-Log "Opcao invalida." 'AVISO' }
    }
    PauseMenu
}

# -- SPO 8: Sites sem atividade (inativos) --
function Get-SPOSitesInativos {
    Show-Banner; Show-Titulo "SHAREPOINT - SITES INATIVOS"
    if (-not (Connect-SharePointOnline)) { PauseMenu; return }

    Write-Host "  Sites sem modificacao ha quantos dias? [ENTER = 90]: " -NoNewline -ForegroundColor White
    $ds = Read-Host; $dias = 90
    [int]::TryParse($ds, [ref]$dias) | Out-Null
    if ($dias -le 0) { $dias = 90 }
    $corte = (Get-Date).AddDays(-$dias)

    Write-Host ""; Write-Host "  Buscando sites inativos..." -ForegroundColor Gray
    try {
        $todos   = @(Get-SPOSite -Limit All -ErrorAction Stop)
        # Filtra inativos de forma compativel com PS 5.1 (sem try inline)
        $inativos = $todos | Where-Object {
            if ($null -eq $_.LastContentModifiedDate) { return $false }
            if ($_.Template -like 'SPSPERS*') { return $false }
            try {
                $dtMod = [datetime]$_.LastContentModifiedDate
                return $dtMod -lt $corte
            } catch {
                return $false
            }
        } | Sort-Object LastContentModifiedDate

        if (@($inativos).Count -eq 0) {
            Write-Log "Nenhum site inativo ha mais de $dias dias." 'SUCESSO'; PauseMenu; return
        }

        Write-Host ""; Show-Separador
        Write-Host ("  {0,-40} {1,12} {2,-20}" -f "SITE", "USADO", "ULTIMA MODIF.") -ForegroundColor Cyan
        Show-Separador

        $csv = @()
        foreach ($s in $inativos) {
            $usadoStr = Format-Bytes ((Get-SPOSiteStorage -Site $s) * 1MB)
            $dtMod = try { if ($s.LastContentModifiedDate) { ([datetime]$s.LastContentModifiedDate).ToString('dd/MM/yyyy') } else { 'Nunca' } } catch { 'Nunca' }
            $tNome2   = if ($s.Title) { $s.Title } else { $s.Url }
            $titulo   = if ($tNome2.Length -gt 38) { $tNome2.Substring(0,38) + '..' } else { $tNome2 }
            Write-Host ("  {0,-40} {1,12} {2}" -f $titulo, $usadoStr, $dtMod) -ForegroundColor Yellow
            $csv += [PSCustomObject]@{
                Site          = $s.Title
                URL           = $s.Url
                UsadoMB       = Get-SPOSiteStorage -Site $s
                UltimaModif   = $dtMod
                Owner         = $s.Owner
                Template      = $s.Template
            }
        }
        Show-Separador
        Write-Host "  Total: $(@($inativos).Count) site(s) inativo(s)" -ForegroundColor DarkGray
        Write-Host ""; Write-Host "  Exportar CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -match '^[Ss]$') { Export-ParaCSV -Dados $csv -NomeArquivo "SPO_SitesInativos_${dias}dias" }
    } catch {
        Write-Log "Erro: $($_.Exception.Message)" 'ERRO'
    }
    PauseMenu
}

# -- SPO 9: Relatorio de uso OneDrive por usuario --
function Get-SPOOneDriveUsuarios {
    Show-Banner; Show-Titulo "SHAREPOINT - ONEDRIVE POR USUARIO"
    if (-not (Connect-SharePointOnline)) { PauseMenu; return }

    Write-Host "  Ordenar por:" -ForegroundColor Cyan
    Write-Host "  [1] Uso (maior primeiro)  [2] Nome (A-Z): " -NoNewline -ForegroundColor White
    $ord = Read-Host

    Write-Host ""; Write-Host "  Buscando OneDrives..." -ForegroundColor Gray
    try {
        $ods = @(Get-SPOSite -IncludePersonalSite $true -Template SPSPERS -Limit All)

        if ($ord.Trim() -eq '2') { $ods = $ods | Sort-Object Owner }
        else { $ods = $ods | Sort-Object { Get-SPOSiteStorage -Site $_ } -Descending }

        $propN3  = Get-SPOStoragePropName
        $totalMB = 0; foreach ($__od in $ods) { try { $v = $__od.$propN3; if ($v -gt 0) { $totalMB += [long]$v } } catch { } }
        $totalStr = Format-Bytes ($totalMB * 1MB)

        Write-Host ""; Show-Separador
        Write-Host ("  {0,-42} {1,12} {2,10}" -f "USUARIO", "USADO", "% DO TOTAL") -ForegroundColor Cyan
        Show-Separador

        $csv = @()
        foreach ($od in $ods) {
            $__odMB   = Get-SPOSiteStorage -Site $od
            $usadoStr = Format-Bytes ($__odMB * 1MB)
            $pct      = if ($totalMB -gt 0) { [math]::Round($__odMB / $totalMB * 100, 1) } else { 0 }
            $cor      = if ($__odMB * 1MB -gt 50GB) { 'Red' } elseif ($__odMB * 1MB -gt 20GB) { 'Yellow' } else { 'White' }
            $ownerSafe = if ($od.Owner) { $od.Owner } else { $od.Url }
            $owner     = if ($ownerSafe.Length -gt 40) { $ownerSafe.Substring(0,40) + '..' } else { $ownerSafe }
            Write-Host ("  {0,-42} {1,12} {2,9}%" -f $owner, $usadoStr, $pct) -ForegroundColor $cor
            $csv += [PSCustomObject]@{
                Usuario  = $od.Owner
                URL      = $od.Url
                UsadoMB  = $__odMB
                Usado    = $usadoStr
                PctTotal = "$pct%"
            }
        }
        Show-Separador
        Write-Host "  Total OneDrive do tenant: $totalStr ($(@($ods).Count) usuarios)" -ForegroundColor DarkGray
        Write-Host ""; Write-Host "  Exportar CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -match '^[Ss]$') { Export-ParaCSV -Dados $csv -NomeArquivo 'SPO_OneDrive_Usuarios' }
    } catch {
        Write-Log "Erro: $($_.Exception.Message)" 'ERRO'
    }
    PauseMenu
}

# -- SPO 10: Excluir site --
function Remove-SPOSite {
    Show-Banner; Show-Titulo "SHAREPOINT - EXCLUIR SITE" 'Red'
    if (-not (Connect-SharePointOnline)) { PauseMenu; return }

    Write-Host "  ATENCAO: Esta acao move o site para a lixeira do SharePoint." -ForegroundColor Red
    Write-Host "  O site pode ser recuperado em ate 93 dias." -ForegroundColor Yellow
    Write-Host ""; Show-Separador; Write-Host ""

    Write-Host "  URL do site: " -NoNewline -ForegroundColor White
    $url = Read-Host
    if ([string]::IsNullOrWhiteSpace($url)) { Write-Log "URL invalida." 'AVISO'; PauseMenu; return }

    Write-Host ""; Write-Host "  Buscando site..." -ForegroundColor Gray
    try { $s = Get-SPOSite -Identity $url -ErrorAction Stop }
    catch { Write-Log "Site nao encontrado." 'ERRO'; PauseMenu; return }

    Write-Host ""; Show-Separador
    Write-Host ("  Titulo   : {0}" -f $s.Title)   -ForegroundColor White
    Write-Host ("  URL      : {0}" -f $s.Url)     -ForegroundColor White
    Write-Host ("  Usado    : {0}" -f (Format-Bytes ((Get-SPOSiteStorage -Site $s) * 1MB))) -ForegroundColor White
    Write-Host ("  Owner    : {0}" -f $s.Owner)   -ForegroundColor White
    Show-Separador; Write-Host ""

    Write-Host "  CONFIRMACAO 1/2 - Excluir este site? [S/N]: " -NoNewline -ForegroundColor Yellow
    if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }
    Write-Host "  CONFIRMACAO 2/2 - Digite o nome do site para confirmar: " -NoNewline -ForegroundColor Yellow
    if ((Read-Host).Trim() -ne $s.Title) { Write-Log "Nome nao confere. Cancelado." 'ERRO'; PauseMenu; return }

    try {
        Remove-SPOSite -Identity $url -Confirm:$false -ErrorAction Stop
        Write-Log "ACAO: Site excluido - $url - $($s.Title)" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> SITE MOVIDO PARA A LIXEIRA! <<<" -ForegroundColor Yellow
        Write-Host "  Recuperavel por 93 dias no centro de admin do SharePoint." -ForegroundColor DarkGray
    } catch {
        Write-Log "Erro: $($_.Exception.Message)" 'ERRO'
    }
    PauseMenu
}

# -- SPO 11: Politica de compartilhamento externo --
function Set-SPOCompartilhamentoExterno {
    Show-Banner; Show-Titulo "SHAREPOINT - COMPARTILHAMENTO EXTERNO"
    if (-not (Connect-SharePointOnline)) { PauseMenu; return }

    Write-Host "  Buscando configuracao atual..." -ForegroundColor Gray
    try {
        $tenant = Get-SPOTenantInfo -ErrorAction Stop
        Write-Host ""; Show-Separador
        Write-Host "  Compartilhamento atual: $($tenant.SharingCapability)" -ForegroundColor White
        Show-Separador; Write-Host ""
        Write-Host "  Opcoes de compartilhamento:" -ForegroundColor Cyan
        Write-Host "  [1] Disabled      - Sem compartilhamento externo (mais seguro)" -ForegroundColor White
        Write-Host "  [2] ExistingExternalUserSharingOnly - Apenas usuarios externos ja conhecidos" -ForegroundColor White
        Write-Host "  [3] ExternalUserSharingOnly - Usuarios externos (com conta Microsoft)" -ForegroundColor White
        Write-Host "  [4] ExternalUserAndGuestSharing - Qualquer pessoa (link anonimo)" -ForegroundColor White
        Write-Host ""; Write-Host "  Escolha: " -NoNewline -ForegroundColor White
        $op = Read-Host

        $nivel = switch ($op.Trim()) {
            '1' { 'Disabled' }
            '2' { 'ExistingExternalUserSharingOnly' }
            '3' { 'ExternalUserSharingOnly' }
            '4' { 'ExternalUserAndGuestSharing' }
            default { $null }
        }
        if ($null -eq $nivel) { Write-Log "Opcao invalida." 'AVISO'; PauseMenu; return }

        Write-Host ""; Write-Host "  Confirmar alteracao para '$nivel'? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -notmatch '^[Ss]$') { Write-Log "Cancelado." 'AVISO'; PauseMenu; return }

        Set-SPOTenant -SharingCapability $nivel -ErrorAction Stop
        Write-Log "ACAO: Compartilhamento externo alterado - $nivel" 'AUDITORIA'
        Write-Host ""; Write-Host "  >>> CONFIGURACAO ATUALIZADA: $nivel <<<" -ForegroundColor Green
    } catch {
        Write-Log "Erro: $($_.Exception.Message)" 'ERRO'
    }
    PauseMenu
}

# -- SPO 12: Relatorio geral de armazenamento --
function Get-SPORelatorioArmazenamento {
    Show-Banner; Show-Titulo "SHAREPOINT - RELATORIO COMPLETO DE ARMAZENAMENTO"
    if (-not (Connect-SharePointOnline)) { PauseMenu; return }

    Write-Host "  Coletando dados completos (pode demorar)..." -ForegroundColor Gray
    try {
        $tenant   = Get-SPOTenant -ErrorAction Stop
        $sites    = Get-SPOSite -Limit All -ErrorAction Stop
        $onedrives = $sites | Where-Object { $_.Url -like '*-my.sharepoint.com/personal/*' }
        $sitesNorm = $sites | Where-Object { $_.Url -notlike '*-my.sharepoint.com/personal/*' }

        $totalGB       = $tenant.StorageQuota / 1024
        # Calcula totais via helper (campo pode vir vazio no bulk)
        $propN2 = Get-SPOStoragePropName
        $usadoSiteMB = 0; foreach ($__s in $sitesNorm) { try { $v = $__s.$propN2; if ($v -gt 0) { $usadoSiteMB += [long]$v } } catch { } }
        $usadoODMB   = 0; foreach ($__s in $onedrives)  { try { $v = $__s.$propN2; if ($v -gt 0) { $usadoODMB   += [long]$v } } catch { } }
        $usadoTotalMB = $usadoSiteMB + $usadoODMB
        $usadoTotalGB  = $usadoTotalMB / 1024
        $livreGB       = $totalGB - $usadoTotalGB
        $pct = 0
        if ($totalGB -gt 0 -and $usadoTotalGB -ge 0) {
            $pct = [math]::Round(($usadoTotalGB / $totalGB) * 100, 1)
            if ($pct -gt 100) { $pct = 100 }
        }

        Write-Host ""
        Write-Host "  ======================================================" -ForegroundColor Cyan
        Write-Host "  RESUMO GERAL DE ARMAZENAMENTO" -ForegroundColor Cyan
        Write-Host "  ======================================================" -ForegroundColor Cyan
        Write-Host ("  Quota total do tenant    : {0,15}" -f (Format-Bytes ($totalGB * 1GB)))     -ForegroundColor White
        Write-Host ("  Total em uso             : {0,15}" -f (Format-Bytes ($usadoTotalMB * 1MB))) -ForegroundColor White
        Write-Host ("  Livre                    : {0,15}" -f (Format-Bytes ($livreGB * 1GB)))      -ForegroundColor $(if ($livreGB -lt 50) { 'Red' } elseif ($livreGB -lt 100) { 'Yellow' } else { 'Green' })
        Write-Host ("  Uso percentual           : {0,14}%" -f $pct) -ForegroundColor White
        Write-Host ""
        Write-Host "  Barra de uso:" -ForegroundColor Gray
        Show-BarraUso -Porcentagem $pct -Largura 40
        Write-Host ""
        Write-Host "  ======================================================" -ForegroundColor Cyan
        Write-Host "  DETALHAMENTO POR TIPO" -ForegroundColor Cyan
        Write-Host "  ======================================================" -ForegroundColor Cyan
        Write-Host ("  Sites SharePoint ({0,4} sites) : {1,15}" -f @($sitesNorm).Count, (Format-Bytes ($usadoSiteMB * 1MB))) -ForegroundColor White
        Write-Host ("  OneDrive pessoal ({0,4} users) : {1,15}" -f @($onedrives).Count, (Format-Bytes ($usadoODMB * 1MB)))  -ForegroundColor White
        Write-Host ""
        Write-Host "  ======================================================" -ForegroundColor Cyan
        Write-Host "  TOP 5 - MAIORES CONSUMIDORES" -ForegroundColor Cyan
        Write-Host "  ======================================================" -ForegroundColor Cyan
        $top5 = $sites | ForEach-Object { $mb2 = Get-SPOSiteStorage -Site $_; $_ | Add-Member -NotePropertyName StorageMB2 -NotePropertyValue $mb2 -Force -PassThru } | Sort-Object StorageMB2 -Descending | Select-Object -First 5
        $pos = 1
        foreach ($s in $top5) {
            Write-Host ("  [{0}] {1,-40} {2}" -f $pos, ($s.Title -replace '.{38}$','..'), (Format-Bytes ($s.StorageMB2 * 1MB))) -ForegroundColor White
            $pos++
        }
        Write-Host ""

        # Monta CSV completo
        $csv = $sites | ForEach-Object {
            $__mb = Get-SPOSiteStorage -Site $_
            [PSCustomObject]@{
                Tipo       = if ($_.Url -like '*-my.sharepoint.com/personal/*') { 'OneDrive' } else { 'SharePoint' }
                Nome       = $_.Title
                URL        = $_.Url
                UsadoMB    = $__mb
                Usado      = Format-Bytes ($__mb * 1MB)
                QuotaMB    = $_.StorageQuota
                Status     = $_.LockState
                Owner      = $_.Owner
                Template   = $_.Template
                UltimaModif = $_.LastContentModifiedDate
            }
        }
        Write-Host "  Exportar relatorio completo CSV? [S/N]: " -NoNewline -ForegroundColor Yellow
        if (Read-Host -match '^[Ss]$') { Export-ParaCSV -Dados $csv -NomeArquivo 'SPO_RelatorioCompleto' }
    } catch {
        Write-Log "Erro: $($_.Exception.Message)" 'ERRO'
    }
    PauseMenu
}

# =============================================================
#  MENU PRINCIPAL
# =============================================================
function Show-MenuPrincipal {
    Show-Banner; Show-Header
    Write-Host "  +============================================+" -ForegroundColor Cyan
    Write-Host "  |            MENU PRINCIPAL                 |" -ForegroundColor Cyan
    Write-Host "  +============================================+" -ForegroundColor Cyan
    Write-Host "  |                                           |" -ForegroundColor Cyan
    Write-Host "  |  [1]  Gestao de Usuarios                 |" -ForegroundColor White
    Write-Host "  |  [2]  Gestao de Licencas                 |" -ForegroundColor White
    Write-Host "  |  [3]  Gestao de Grupos                   |" -ForegroundColor White
    Write-Host "  |  [4]  Seguranca e Conformidade           |" -ForegroundColor White
    Write-Host "  |  [5]  Caixas de E-mail e Recursos        |" -ForegroundColor White
    Write-Host "  |  [6]  Relatorios e Exportacao            |" -ForegroundColor White
    Write-Host "  |                                          |" -ForegroundColor Cyan
    Write-Host "  |  [7]  Offboarding Completo               |" -ForegroundColor Yellow
    Write-Host "  |  [8]  Onboarding Completo                |" -ForegroundColor Yellow
    Write-Host "  |  [9]  Gestao do SharePoint               |" -ForegroundColor Magenta
    Write-Host "  |                                          |" -ForegroundColor Cyan
    Write-Host "  |  [0]  Sair                               |" -ForegroundColor DarkGray
    Write-Host "  |                                          |" -ForegroundColor Cyan
    Write-Host "  +============================================+" -ForegroundColor Cyan
    Write-Host ""; Write-Host "  Escolha: " -NoNewline -ForegroundColor White
}


function Start-MenuSharePoint {
    $loop = $true
    while ($loop) {
        Show-Banner; Show-Header
        Write-Host "  +------------------------------------------+" -ForegroundColor Magenta
        Write-Host "  |      GESTAO DO SHAREPOINT ONLINE         |" -ForegroundColor Magenta
        Write-Host "  +------------------------------------------+" -ForegroundColor Magenta
        Write-Host "  |  [1]  Visao geral do tenant              |" -ForegroundColor White
        Write-Host "  |  [2]  Listar todos os sites              |" -ForegroundColor White
        Write-Host "  |  [3]  Top consumidores de espaco         |" -ForegroundColor White
        Write-Host "  |  [4]  Detalhes de um site               |" -ForegroundColor White
        Write-Host "  |  [5]  Alterar quota de um site          |" -ForegroundColor White
        Write-Host "  |  [6]  Bloquear / Desbloquear site       |" -ForegroundColor White
        Write-Host "  |  [7]  Gerenciar proprietario do site    |" -ForegroundColor White
        Write-Host "  |  [8]  Sites inativos                    |" -ForegroundColor White
        Write-Host "  |  [9]  OneDrive por usuario              |" -ForegroundColor White
        Write-Host "  |  [10] Excluir site                      |" -ForegroundColor White
        Write-Host "  |  [11] Compartilhamento externo          |" -ForegroundColor White
        Write-Host "  |  [12] Relatorio completo armazenamento  |" -ForegroundColor White
        Write-Host "  |                                         |" -ForegroundColor Magenta
        Write-Host "  |  [0]  Voltar                            |" -ForegroundColor DarkGray
        Write-Host "  +-----------------------------------------+" -ForegroundColor Magenta
        Write-Host ""; Write-Host "  Escolha: " -NoNewline -ForegroundColor White
        switch ((Read-Host).Trim()) {
            '1'  { Get-SPOVisaoGeral }
            '2'  { Get-SPOListarSites }
            '3'  { Get-SPOTopConsumidores }
            '4'  { Get-SPODetalhesSite }
            '5'  { Set-SPOQuotaSite }
            '6'  { Set-SPOBloqueioSite }
            '7'  { Set-SPOProprietarioSite }
            '8'  { Get-SPOSitesInativos }
            '9'  { Get-SPOOneDriveUsuarios }
            '10' { Remove-SPOSite }
            '11' { Set-SPOCompartilhamentoExterno }
            '12' { Get-SPORelatorioArmazenamento }
            '0'  { $loop = $false }
            default { Write-Host "  >> Invalido." -ForegroundColor Yellow; Start-Sleep -Seconds 1 }
        }
    }
}

# =============================================================
#  PONTO DE ENTRADA
# =============================================================
function Start-App {
    Install-ModulosNecessarios
    $conectado = $false
    while (-not $conectado) { $conectado = Connect-Tenant }
    $sair = $false
    while (-not $sair) {
        Show-MenuPrincipal
        switch ((Read-Host).Trim()) {
            '1' { Start-MenuUsuarios }
            '2' { Start-MenuLicencas }
            '3' { Start-MenuGrupos }
            '4' { Start-MenuSeguranca }
            '5' { Start-MenuExchange }
            '6' { Start-MenuRelatorios }
            '7' { Invoke-OffboardingCompleto }
            '8' { Invoke-OnboardingCompleto }
            '9' { Start-MenuSharePoint }
            '0' {
                Write-Host ""
                Write-Log "Encerrando. Ate logo!" 'INFO'
                try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
                try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }
                $sair = $true
            }
            default { Write-Host "  >> Invalido." -ForegroundColor Yellow; Start-Sleep -Seconds 1 }
        }
    }
}

Start-App