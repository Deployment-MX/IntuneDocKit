<#
.SYNOPSIS
    IntuneDocKit - Script Maestro v5.0

.DESCRIPTION
    Primera vez: Crea App Registration y exporta en sesiones separadas automaticamente.
    Siguientes veces: Usa -SkipAppCreation para ir directo a exportar.

.EXAMPLE
    # Primera vez (crea App y exporta automaticamente)
    .\MasterExport-MEM.ps1 -TenantId "62066f2a-c776-4693-af2b-9a98c1ad0b6d"

    # Siguientes veces
    .\MasterExport-MEM.ps1 -TenantId "62066f2a-c776-4693-af2b-9a98c1ad0b6d" -SkipAppCreation

.NOTES
    Proyecto : IntuneDocKit
    Version  : 5.0 - Sesiones separadas para Graph y MSAL
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true)]  [string]$TenantId,
    [Parameter(Mandatory = $false)] [string]$AppDisplayName = "IntuneDocKit",
    [Parameter(Mandatory = $false)] [string]$XmlConfigPath = "",
    [Parameter(Mandatory = $false)] [string]$ExportScriptPath = "",
    [Parameter(Mandatory = $false)] [string]$ExportPath = "",
    [Parameter(Mandatory = $false)] [string]$ClientName = "",
    [Parameter(Mandatory = $false)] [switch]$SkipAppCreation,
    [Parameter(Mandatory = $false)] [switch]$OnlyCreateApp,
    [Parameter(Mandatory = $false)] [switch]$ExportOnly  # Uso interno
)

#region -- Helpers
function Write-Step { param([string]$m) Write-Host ""; Write-Host "=================================================" -ForegroundColor Cyan; Write-Host "  $m" -ForegroundColor Cyan; Write-Host "=================================================" -ForegroundColor Cyan }
function Write-OK   { param([string]$m) Write-Host "  [OK] $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "  [!]  $m" -ForegroundColor Yellow }
function Write-Fail { param([string]$m) Write-Host "  [X]  $m" -ForegroundColor Red }
#endregion

#region -- Rutas por defecto
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ([string]::IsNullOrEmpty($XmlConfigPath))    { $XmlConfigPath    = Join-Path $ScriptRoot "Export-MEMConfiguration.xml" }
if ([string]::IsNullOrEmpty($ExportScriptPath)) { $ExportScriptPath = Join-Path $ScriptRoot "Export-MEMConfiguration.ps1" }
foreach ($file in @($XmlConfigPath, $ExportScriptPath)) {
    if (!(Test-Path $file)) { Write-Fail "No se encontro: $file"; Exit 1 }
}
#endregion

#region -- Modo ExportOnly: solo PASOS 1 y 5 (sesion limpia sin Graph)
if ($ExportOnly) {

    Write-Step "PASO 1 -- Cargando MSAL.PS"
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

    $msalBase = "C:\Program Files\WindowsPowerShell\Modules\MSAL.PS"
    $msalPsd1 = Get-ChildItem -Path $msalBase -Filter "MSAL.PS.psd1" -Recurse -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending | Select-Object -First 1
    if (-not $msalPsd1) {
        Write-Fail "MSAL.PS no encontrado. Instala con: Install-Module MSAL.PS -Scope AllUsers -Force -AcceptLicense"
        Exit 1
    }
    Remove-Module MSAL.PS -Force -ErrorAction SilentlyContinue
    Import-Module $msalPsd1.FullName -Force
    Write-OK "MSAL.PS cargado: $($msalPsd1.FullName)"

    Write-Step "PASO 5 -- Exportando"

    if ([string]::IsNullOrEmpty($ClientName)) {
        Write-Host ""
        Write-Host "  Este nombre aparecera en la portada del Word." -ForegroundColor Cyan
        $ClientName = Read-Host "  Nombre del cliente (ej: Contoso Inc.)"
        while ([string]::IsNullOrWhiteSpace($ClientName)) {
            $ClientName = Read-Host "  Nombre del cliente (ej: Contoso Inc.)"
        }
    }
    Write-OK "Client name: $ClientName"

    Write-Host ""
    Write-Host "  El script pedira confirmar el tenant (escribe Y)." -ForegroundColor Cyan
    Write-Host "  Luego abrira Device Code -- usa una cuenta con lectura en Intune." -ForegroundColor Cyan
    Write-Host ""

    $exportArgs = @{ Config=$XmlConfigPath; ClientName=$ClientName; Force=$true }
    if (-not [string]::IsNullOrEmpty($ExportPath)) { $exportArgs["ExportPath"] = $ExportPath }

    try {
        & $ExportScriptPath @exportArgs *>&1
        Write-OK "Exportacion completada."
    } catch {
        Write-Fail "Error: $($_.Exception.Message)"
        Exit 1
    }

    Write-Host ""
    Write-Host "=================================================" -ForegroundColor Green
    Write-Host "  COMPLETADO -- IntuneDocKit finalizado OK." -ForegroundColor Green
    Write-Host "=================================================" -ForegroundColor Green
    Exit 0
}
#endregion

#region -- Flujo normal: PASOS 1-4 (con Graph) luego relanza en sesion limpia para PASO 5

Write-Step "PASO 1 -- Verificando modulos"
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
Write-OK "ExecutionPolicy = Bypass"

# ✅ NUEVO: Instalar NuGet y confiar en PSGallery ANTES del loop
Write-Host "[*] Preparando proveedor NuGet..." -ForegroundColor Cyan
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null

Write-Host "[*] Confiando en PSGallery..." -ForegroundColor Cyan
Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted

$mods = @(
    "PSWriteWord",
    "PSWriteHTML",
    "ImportExcel",
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Applications",
    "Microsoft.Graph.Identity.SignIns"
)

foreach ($mod in $mods) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Warn "Instalando $mod..."
        Install-Module $mod -Scope CurrentUser -Force -AllowClobber -ErrorAction SilentlyContinue
    }
    Write-OK "$mod OK"
}

if (-not (Get-Module -ListAvailable -Name "AzureAD" | Where-Object { $_.Version -eq "2.0.2.140" })) {
    Install-Module AzureAD -RequiredVersion 2.0.2.140 -Scope CurrentUser -Force -AllowClobber -ErrorAction SilentlyContinue
}
Write-OK "AzureAD OK"

$msalBase = "C:\Program Files\WindowsPowerShell\Modules\MSAL.PS"
if (-not (Test-Path $msalBase)) {
    Write-Warn "Instalando MSAL.PS..."
    Install-Module MSAL.PS -Scope AllUsers -Force -AllowClobber
}
Write-OK "MSAL.PS OK"

if (-not $SkipAppCreation) {

    Write-Step "PASO 2 -- Creando App Registration"
    $graphMods = @("Microsoft.Graph.Authentication","Microsoft.Graph.Applications","Microsoft.Graph.Identity.SignIns")
    foreach ($gm in $graphMods) { Import-Module $gm -Force -ErrorAction SilentlyContinue }

    try {
        Connect-MgGraph -TenantId $TenantId -Scopes @("Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All","DelegatedPermissionGrant.ReadWrite.All") -ErrorAction Stop
        Write-OK "Conectado a Microsoft Graph."
    } catch { Write-Fail "Error Graph: $($_.Exception.Message)"; Exit 1 }

    $app = Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -ErrorAction SilentlyContinue
    if ($app) {
        Write-Warn "App '$AppDisplayName' ya existe. AppId: $($app.AppId)"
    } else {
        $app = New-MgApplication -DisplayName $AppDisplayName -SignInAudience "AzureADMyOrg"
        Write-OK "App creada: $($app.AppId)"
    }

    try { Update-MgApplication -ApplicationId $app.Id -IsFallbackPublicClient:$true; Write-OK "Public client flows habilitado." }
    catch { Write-Warn "Habilita manualmente: Azure Portal > $AppDisplayName > Authentication > Allow public client flows = Yes" }

    $sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue
    if (-not $sp) { $sp = New-MgServicePrincipal -AppId $app.AppId; Write-OK "Service Principal creado." }

    Write-Step "PASO 3 -- Permisos + Admin Consent"
    $graphAppId = "00000003-0000-0000-c000-000000000000"
    $graphSP    = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
    $perms = @(
        @{ Name="openid";                                  Id="37f7f235-527c-4136-accd-4a02d197296e" },
        @{ Name="profile";                                 Id="14dad69e-099b-42c9-810b-d002981feec1" },
        @{ Name="offline_access";                          Id="7427e0e9-2fba-42fe-b0c0-848c9e6a8182" },
        @{ Name="User.Read";                               Id="e1fe6dd8-ba31-4d61-89e7-88639da4683d" },
        @{ Name="DeviceManagementApps.Read.All";           Id="4edf5f54-4666-44af-9de9-0144fb4b6e8c" },
        @{ Name="DeviceManagementConfiguration.Read.All";  Id="f1493658-876a-4c87-8729-b9e6a8e64d53" },
        @{ Name="DeviceManagementManagedDevices.Read.All"; Id="314874da-47d6-4978-88dc-cf0d37f0bb82" },
        @{ Name="DeviceManagementRBAC.Read.All";           Id="49f0cc30-024c-4dfd-ab3e-82e137ee5714" },
        @{ Name="DeviceManagementServiceConfig.Read.All";  Id="8696daa5-bce3-4941-b3c1-3664e4211262" },
        @{ Name="DeviceManagementScripts.Read.All";        Id="b3630069-03ba-431a-9061-a86fb8aa7b88" },
        @{ Name="Directory.Read.All";                      Id="06da0dbc-49e2-44d2-8312-53f166ab848a" },
        @{ Name="User.Read.All";                           Id="a154be20-db9c-4678-8ab7-66f6cc099a59" },
        @{ Name="Group.Read.All";                          Id="5f8c59db-677d-491f-a6b8-5f174b11ec1d" },
        @{ Name="Organization.Read.All";                   Id="4908d5b9-3fb2-4b1e-9336-1888b7937185" },
        @{ Name="Application.Read.All";                    Id="c79f8feb-a9db-4090-85f9-90d820caa0eb" },
        @{ Name="Policy.Read.All";                         Id="572fea84-0151-49b2-9301-11cb16974376" },
        @{ Name="AuditLog.Read.All";                       Id="e4c9e354-4dc5-45b8-9e7c-e1393b0b1a20" },
        @{ Name="Agreement.Read.All";                      Id="af2819c9-df71-4dd3-ade7-4d7c9dc653b7" },
        @{ Name="CloudPC.Read.All";                        Id="5252ec4e-fd40-4d92-8c68-89dd1d3c6110" }
    )

    # ---- Permisos adicionales para nuevas secciones ----
    $perms += @(
        @{ Name="Policy.ReadWrite.ConditionalAccess"; Id="ad902697-1014-4ef5-81ef-2b4301988e8e" }
       )

    $appRoles = $perms | ForEach-Object { @{ id=$_.Id; type="Scope" } }
    try { Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess @(@{ resourceAppId=$graphAppId; resourceAccess=$appRoles }); Write-OK "Permisos asignados." }
    catch { Write-Fail "Error permisos: $($_.Exception.Message)"; Exit 1 }

    Start-Sleep -Seconds 5
    $scopeStr = ($perms | ForEach-Object { $_.Name }) -join " "
    $grant = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($sp.Id)' and resourceId eq '$($graphSP.Id)'" -ErrorAction SilentlyContinue
    try {
        if ($grant) { Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $grant.Id -Scope $scopeStr | Out-Null }
        else { New-MgOauth2PermissionGrant -ClientId $sp.Id -ConsentType "AllPrincipals" -ResourceId $graphSP.Id -Scope $scopeStr | Out-Null }
        Write-OK "Admin Consent otorgado."
    } catch { Write-Warn "Consent manual requerido: Azure Portal > $AppDisplayName > API Permissions > Grant admin consent" }

    $newAppId = $app.AppId
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

} else {
    Write-Step "PASO 2/3 -- Omitidos (SkipAppCreation)"
    $newAppId = $null
}

Write-Step "PASO 4 -- Actualizando XML"
[xml]$xml = Get-Content $XmlConfigPath -Encoding UTF8
if (-not $SkipAppCreation -and $newAppId) {
    $xml.root.Configuration.Tenant   = $TenantId
    $xml.root.Configuration.ClientId = $newAppId
    Write-OK "TenantId: $TenantId"
    Write-OK "ClientId: $newAppId"
} else {
    Write-OK "TenantId en XML: $($xml.root.Configuration.Tenant)"
    Write-OK "ClientId  en XML: $($xml.root.Configuration.ClientId)"
}
if (-not [string]::IsNullOrEmpty($ExportPath)) { $xml.root.Configuration.ExportPath = $ExportPath }
$xml.Save($XmlConfigPath)
Write-OK "XML guardado: $XmlConfigPath"

if ([string]::IsNullOrEmpty($xml.root.Configuration.Tenant) -or [string]::IsNullOrEmpty($xml.root.Configuration.ClientId)) {
    Write-Fail "Tenant o ClientId vacio en XML. Corre sin -SkipAppCreation."
    Exit 1
}

# Si solo se queria crear la App -- salir aqui
if ($OnlyCreateApp) {
    Write-Host ""
    Write-Host "=================================================" -ForegroundColor Green
    Write-Host "  App Registration creada y XML actualizado OK." -ForegroundColor Green
    Write-Host "  Siguiente paso -- exportar con -SkipAppCreation" -ForegroundColor Green
    Write-Host "=================================================" -ForegroundColor Green
    Exit 0
}

# Pedir ClientName ahora antes de relanzar para pasarlo como parametro
if ([string]::IsNullOrEmpty($ClientName)) {
    Write-Host ""
    Write-Host "  Este nombre aparecera en la portada del Word." -ForegroundColor Cyan
    $ClientName = Read-Host "  Nombre del cliente (ej: Contoso Inc.)"
    while ([string]::IsNullOrWhiteSpace($ClientName)) {
        $ClientName = Read-Host "  Nombre del cliente (ej: Contoso Inc.)"
    }
}
Write-OK "Client name: $ClientName"

# Relanzar en sesion nueva y limpia para el PASO 5 (sin contexto de Graph)
Write-Host ""
Write-Host "  Iniciando sesion limpia para exportacion (sin contexto Graph)..." -ForegroundColor Cyan

# Cuando corre desde IntuneDocKit GUI, ejecutar directamente sin relanzar
# para mantener stdout redirigido y capturar el Device Code
if ($env:INTUMEDOCKIT_GUI -eq "1") {
    Write-Host "  [IntuneDocKit] Modo GUI - ejecutando ExportOnly en proceso actual..." -ForegroundColor Cyan
    
    # Desconectar Graph para sesion limpia
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Remove-Module Microsoft.Graph.Authentication -Force -ErrorAction SilentlyContinue

    # Recargar MSAL.PS limpio
    $msalBase = "C:\Program Files\WindowsPowerShell\Modules\MSAL.PS"
    $msalPsd1 = Get-ChildItem -Path $msalBase -Filter "MSAL.PS.psd1" -Recurse -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending | Select-Object -First 1
    if ($msalPsd1) {
        Remove-Module MSAL.PS -Force -ErrorAction SilentlyContinue
        Import-Module $msalPsd1.FullName -Force
        Write-Host "  [OK] MSAL.PS cargado para exportacion." -ForegroundColor Green
    }

    $exportArgs = @{ Config=$XmlConfigPath; ClientName=$ClientName; Force=$true }
    if (-not [string]::IsNullOrEmpty($ExportPath)) { $exportArgs["ExportPath"] = $ExportPath }

    $exportArgStr = "-NoProfile -ExecutionPolicy Bypass -File `"$ExportScriptPath`" -Config `"$XmlConfigPath`" -ClientName `"$ClientName`" -Force"
    if (-not [string]::IsNullOrEmpty($ExportPath)) { $exportArgStr += " -ExportPath `"$ExportPath`"" }
    Write-Host "  Iniciando exportacion con Device Code..." -ForegroundColor Cyan
    $exportProc = Start-Process -FilePath "powershell.exe" -ArgumentList $exportArgStr -PassThru -Wait
    Exit $exportProc.ExitCode
}
$ps51    = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$thisScript = $MyInvocation.MyCommand.Path

$relaunchArgs = @(
    "-ExecutionPolicy", "Bypass",
    "-File", $thisScript,
    "-TenantId", $TenantId,
    "-XmlConfigPath", $XmlConfigPath,
    "-ExportScriptPath", $ExportScriptPath,
    "-ClientName", $ClientName,
    "-ExportOnly"
)
if (-not [string]::IsNullOrEmpty($ExportPath)) {
    $relaunchArgs += "-ExportPath"
    $relaunchArgs += $ExportPath
}

& $ps51 @relaunchArgs
Exit $LASTEXITCODE
#endregion
