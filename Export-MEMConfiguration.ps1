#Install-Module AzureAD -RequiredVersion 2.0.2.140 opcional
#Install-Module MSAL.PS -Scope CurrentUser -Force
#Install-Module MSAL.PS -Scope CurrentUser -Force
#Import-Module MSAL.PS
#Install-Module PSWriteWord  -Scope CurrentUser -Force
#Install-Module PSWriteHTML  -Scope CurrentUser -Force
#Install-Module ImportExcel  -Scope CurrentUser -Force
#Install-Module MSAL.PS      -Scope CurrentUser -Force
#verificacion de modulos 
#
#"PSWriteWord","PSWriteHTML","ImportExcel","MSAL.PS" | ForEach-Object {
#  "{0} -> {1}" -f $_, ([bool](Get-Module -ListAvailable -Name $_))
#}
#Requires -Modules PSWriteWord, MSAL.PS

#region --------------------------------------------------[Script Parameters]------------------------------------------------------
Param (
    [Parameter(Mandatory = $False)] [string]$Tenant      = "",
    [Parameter(Mandatory = $False)] [string]$ExportPath  = "",
    [Parameter(Mandatory = $False)] [string]$DocumentName = "",
    [Parameter(Mandatory = $False)] [string]$ClientName  = "",
    [Parameter(Mandatory = $False)] [string]$Config      = "",
    [Parameter(Mandatory = $False)] [switch]$Force       = $false
)
#endregion

#region ---------------------------------------------------[Declarations]----------------------------------------------------------
$DateTimeRegex   = [regex]"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z"
$script:User     = ""
$script:Tenant   = $Tenant

$global:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogPath           = Join-Path $global:ScriptPath "$([io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)).log"

if ([string]::IsNullOrEmpty($Config)) {
    $Config = Join-Path $global:ScriptPath "Export-MEMConfiguration.xml"
}
$DocumentTemplate = Join-Path $global:ScriptPath "MEMDocumentationTempl.docx"
#endregion

#region ---------------------------------------------------[Logging]------------------------------------------------------------
Function Start-Log {
    [CmdletBinding()]
    param (
        [ValidateScript({ Split-Path $_ -Parent | Test-Path })]
        [string]$FilePath
    )
    try {
        if (!(Test-Path $FilePath)) { New-Item $FilePath -ItemType File | Out-Null }
        $global:ScriptLogFilePath = $FilePath
    }
    catch { Write-Error $_.Exception.Message }
}

Function Write-Log {
    param (
        [Parameter(Mandatory = $true)] [string]$Message,
        [Parameter()] [ValidateSet(1, 2, 3)] [int]$LogLevel = 1
    )
    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'

    $component = if ($MyInvocation.ScriptName) {
        "$($MyInvocation.ScriptName | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)"
    } else { "Unknown" }

    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), $component, $LogLevel
    $Line = $Line -f $LineFormat

    if (Test-Path $global:ScriptLogFilePath) {
        if ((Get-Item $global:ScriptLogFilePath).Length -ge $maxlogfilesize) {
            $backup = $global:ScriptLogFilePath.TrimEnd('g') + '_'
            if (Test-Path $backup) { Remove-Item $backup -Force }
            Rename-Item -Path $global:ScriptLogFilePath -NewName $backup -Force
        }
    }
    Add-Content -Value $Line -Path $global:ScriptLogFilePath
}
#endregion

#region ---------------------------------------------------[Authentication]------------------------------------------------------------
Function Get-AuthToken {
    <#
    .SYNOPSIS
    Acquires a Graph API token using MSAL Device Code flow.
    ClientId and TenantId are read from script-scope variables set during XML load.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string]$User
    )

    # Read-only scopes — script never writes data
    $Scopes = @(
        "openid", "profile", "offline_access",
        "User.Read",
        "Directory.Read.All",
        "Organization.Read.All",
        "User.Read.All",
        "Group.Read.All",
        "Application.Read.All",
        "Policy.Read.All",
        "Agreement.Read.All",
        "DeviceManagementApps.Read.All",
        "DeviceManagementConfiguration.Read.All",
        "DeviceManagementRBAC.Read.All",
        "DeviceManagementServiceConfig.Read.All",
        "DeviceManagementManagedDevices.Read.All",
        "DeviceManagementScripts.Read.All",
        "CloudPC.Read.All"
    )

    # ---- Scopes adicionales para nuevas secciones ----
    $Scopes += @(
        "Policy.ReadWrite.ConditionalAccess"
    )

    Write-Host "Initiating Device Code authentication for tenant: $($script:Tenant)" -ForegroundColor Cyan
    Write-Log "Initiating Device Code authentication for tenant: $($script:Tenant)"

    try {
        # Notificar a IntuneDocKit GUI que device code esta por generarse
        $dcFile = [System.IO.Path]::Combine($env:TEMP, "IDK_dcode.tmp")
        $transcriptFile = [System.IO.Path]::Combine($env:TEMP, "IDK_transcript.tmp")
        [System.IO.File]::WriteAllText($dcFile, "WAITING", [System.Text.Encoding]::UTF8)

        # Usar transcript para capturar Write-Host de MSAL (incluye device code)
        Start-Transcript -Path $transcriptFile -Force | Out-Null
        $tokenResult = Get-MsalToken `
            -ClientId  $script:ClientId `
            -TenantId  $script:Tenant `
            -Scopes    $Scopes `
            -DeviceCode `
            -ErrorAction Stop
        Stop-Transcript | Out-Null

        # Extraer el codigo del transcript
        if (Test-Path $transcriptFile) {
            $transcriptContent = Get-Content $transcriptFile -Raw
            if ($transcriptContent -match "([A-Z0-9]{8,9})") {
                [System.IO.File]::WriteAllText($dcFile, "CODE:$($matches[1])", [System.Text.Encoding]::UTF8)
            }
        }

        [System.IO.File]::WriteAllText($dcFile, "AUTHDONE", [System.Text.Encoding]::UTF8)
        # Auto-resolve the UPN from the token claims when possible
        if ([string]::IsNullOrEmpty($script:User)) {
            $upnClaim = $tokenResult.Account.Username
            if (![string]::IsNullOrEmpty($upnClaim)) {
                $script:User = $upnClaim
                Write-Log "UPN resolved from token: $($script:User)"
            }
        }

        $authHeader = @{
            'Content-Type'  = 'application/json'
            'Authorization' = "Bearer $($tokenResult.AccessToken)"
            'ExpiresOn'     = $tokenResult.ExpiresOn.UtcDateTime   # always a [DateTime]
        }
        return $authHeader
    }
    catch {
        Write-Host "Authentication failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Authentication failed: $($_.Exception.Message)" -LogLevel 3
        throw
    }
}

Function Update-AuthToken {
    <#
    .SYNOPSIS
    Checks token expiry and renews it when necessary.
    #>

    # Helper: prompt for UPN if still empty
    $ensureUser = {
        if ([string]::IsNullOrEmpty($script:User)) {
            $script:User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
            Write-Log "Connecting using user: $($script:User)"
        }
    }

    if ($global:authToken) {
        $DateTime = (Get-Date).ToUniversalTime()

        # ExpiresOn is stored as [DateTime] — safe direct cast
        $expires = [datetime]$global:authToken['ExpiresOn']
        $TokenExpiresMins = [int](($expires - $DateTime).TotalMinutes)

        if ($TokenExpiresMins -le 0) {
            Write-Host "Authentication Token expired $([Math]::Abs($TokenExpiresMins)) minute(s) ago. Re-authenticating..." -ForegroundColor Yellow
            Write-Log "Authentication Token expired $([Math]::Abs($TokenExpiresMins)) minute(s) ago." -LogLevel 2
            & $ensureUser
            Write-Log "Refreshing authToken for the Graph API"
            $global:authToken = Get-AuthToken -User $script:User
        }
        # Token still valid — nothing to do
    }
    else {
        & $ensureUser
        Write-Log "Acquiring authToken for the Graph API"
        $global:authToken = Get-AuthToken -User $script:User
    }
}
#endregion

#region ---------------------------------------------------[Graph API]------------------------------------------------------------
Function Invoke-GraphRequest {
    <#
    .SYNOPSIS
    Single REST call to Graph API with automatic retry on 429 TooManyRequests.
    Respects Retry-After header when present, otherwise uses exponential backoff.
    #>
    param (
        [Parameter(Mandatory)] [string]$Uri,
        [int]$MaxRetries = 5
    )

    $attempt = 0
    while ($attempt -le $MaxRetries) {
        try {
            return Invoke-RestMethod -Uri $Uri -Headers $global:authToken -Method Get
        }
        catch {
            $statusCode = [int]$_.Exception.Response.StatusCode

            if ($statusCode -eq 429) {
                # Read Retry-After header if available, otherwise exponential backoff
                $retryAfter = $_.Exception.Response.Headers['Retry-After']
                $waitSec    = if ($retryAfter) { [int]$retryAfter } else { [math]::Pow(2, $attempt + 1) }
                $waitSec    = [math]::Max($waitSec, 2)   # minimum 2 seconds

                $attempt++
                Write-Host "  Generando el Documento del Cliente....." -ForegroundColor Yellow
                Write-Log  "429 on $Uri — waiting $waitSec sec (retry $attempt/$MaxRetries)" -LogLevel 2
                Start-Sleep -Seconds $waitSec
            }
            else {
                # Non-429 error — log and return null
                $ex           = $_.Exception
                $errorStream  = $ex.Response.GetResponseStream()
                $reader       = New-Object System.IO.StreamReader($errorStream)
                $reader.BaseStream.Position = 0
                $reader.DiscardBufferedData()
                $responseBody = $reader.ReadToEnd()

                Write-Host "HTTP Error — $($ex.Response.StatusCode) $($ex.Response.StatusDescription)" -ForegroundColor Red
                Write-Host "Response body:`n$responseBody" -ForegroundColor Red
                Write-Log  "Request to $Uri failed: $($ex.Response.StatusCode) $($ex.Response.StatusDescription)" -LogLevel 3
                Write-Log  "Response body: $responseBody" -LogLevel 3
                return $null
            }
        }
    }

    Write-Host "  [Graph] Max retries ($MaxRetries) reached for: $Uri" -ForegroundColor Red
    Write-Log  "Max retries reached for $Uri" -LogLevel 3
    return $null
}

Function Get-GraphUri {
    <#
    .SYNOPSIS
    Generic wrapper for Graph API REST calls with automatic pagination and 429 retry.
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)] [ValidateSet("Beta","v1.0")] [string]$ApiVersion,
        [Parameter(Mandatory)] [string]$Class,
        [Parameter()] [string]$Id     = "",
        [Parameter()] [string]$OData  = "",
        [Parameter()] [switch]$Value,
        [Parameter()] [switch]$AuditData
    )

    $uri = "https://graph.microsoft.com/$ApiVersion/$Class"
    if ($Id    -ne "") { $uri = $uri.TrimEnd('/') + "/$Id" }
    if ($OData -ne "") { $uri = $uri + $OData }

    Write-Verbose "GET $uri"
    Write-Log    "GET $uri"
    Update-AuthToken

    if ($Value) {
        if ($AuditData) {
            # Single page only — avoids pulling full audit history
            $result = Invoke-GraphRequest -Uri $uri
            return $result.value
        }
        else {
            # Follow @odata.nextLink for full pagination, retrying each page on 429
            $page     = Invoke-GraphRequest -Uri $uri
            if ($null -eq $page) { return $null }
            $response = $page.value

            while ($page.'@odata.nextLink') {
                $page      = Invoke-GraphRequest -Uri $page.'@odata.nextLink'
                if ($null -eq $page) { break }
                $response += $page.value
            }
            return $response
        }
    }
    else {
        return Invoke-GraphRequest -Uri $uri
    }
}
#endregion

#region ---------------------------------------------------[Helpers]------------------------------------------------------------
Function Format-DataToString {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()][AllowNull()] $Data
    )
    if ($null -eq $Data) { return "" }

    if ($Data -is [array]) { $Data = $Data -join "," }

    $Data = [string]$Data
    $Data = $Data -replace '^@\{|(?<=^@\{.*)\}$|^#microsoft\.graph\.|^@odata\.', ""

    if ($Data -match $DateTimeRegex) {
        try {
            [DateTime]$Date = [DateTime]::Parse($Data)
            $Data = "$($Date.ToShortDateString()) $($Date.ToShortTimeString())"
        }
        catch {}
    }

    if ($Data.Length -ge $MaxStringLength) {
        $Data = $Data.Substring(0, $MaxStringLength) + "..."
    }
    return $Data
}

Function Format-PropertyName {
    <#
    .SYNOPSIS
    Converts camelCase or PascalCase property names into readable words.
    Example: configurationManagerComplianceRequired -> Configuration Manager Compliance Required
    #>
    param ([string]$Name)
    if ([string]::IsNullOrEmpty($Name)) { return $Name }
    # Insert space before each uppercase letter that follows a lowercase letter or digit
    $spaced = [regex]::Replace($Name, '(?<=[a-z0-9])([A-Z])', ' $1')
    # Capitalize first letter
    return $spaced.Substring(0,1).ToUpper() + $spaced.Substring(1)
}

Function Export-JSONData {
    <#
    .SYNOPSIS
    Saves Graph API response data to a JSON file (and optionally CSV).
    Returns the filename created.
    #>
    param (
        [Parameter(Mandatory)] $JSON,
        [Parameter(Mandatory)] [string]$ExportPath,
        [Parameter()] [string]$FileName = "",
        [Parameter()] [switch]$Force
    )
    try {
        if ($null -eq $JSON -or $JSON -eq "") {
            Write-Host "No JSON data provided." -ForegroundColor Red
            return $null
        }
        if ([string]::IsNullOrEmpty($ExportPath)) {
            Write-Host "No export path provided." -ForegroundColor Red
            return $null
        }

        if (!(Test-Path $ExportPath)) {
            if (!$Force) {
                Write-Host "Path '$ExportPath' doesn't exist. Create it? (Y/N)" -ForegroundColor Yellow
                if ((Read-Host) -notin 'y','Y') {
                    Write-Host "Export cancelled." -ForegroundColor Red
                    return $null
                }
            }
            New-Item -ItemType Directory -Path $ExportPath | Out-Null
        }

        $JSON1        = ConvertTo-Json $JSON -Depth 5
        $JSON_Convert = $JSON1 | ConvertFrom-Json

        $displayName = if ([string]::IsNullOrEmpty($FileName)) { $JSON_Convert.displayName } else { $FileName }
        $displayName = $displayName -replace '\<|\>|:|"|/|\\|\||\?|\*', "_"

        $dateSuffix  = if ($script:AppendDate) { "_$(Get-Date -f dd-MM-yyyy-H-mm-ss)" } else { "" }
        $fileNameJSON = "$displayName$dateSuffix.json"

        $JSON1 | Set-Content -LiteralPath (Join-Path $ExportPath $fileNameJSON) -Encoding UTF8
        Write-Host "JSON exported: $(Join-Path $ExportPath $fileNameJSON)" -ForegroundColor Cyan
        Write-Log  "JSON exported: $(Join-Path $ExportPath $fileNameJSON)"

        if ($ExportCSV) {
            $fileNameCSV = "$displayName$dateSuffix.csv"
            $Properties  = ($JSON_Convert | Get-Member -MemberType NoteProperty).Name
            $Object      = New-Object System.Object
            foreach ($prop in $Properties) {
                $Object | Add-Member -MemberType NoteProperty -Name $prop -Value $JSON_Convert.$prop
            }
            $Object | Export-Csv -LiteralPath (Join-Path $ExportPath $fileNameCSV) -Delimiter "," -NoTypeInformation -Append -Encoding UTF8
            Write-Host "CSV  exported: $(Join-Path $ExportPath $fileNameCSV)" -ForegroundColor Cyan
        }

        return $fileNameJSON
    }
    catch {
        Write-Host "Export-JSONData error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log  "Export-JSONData error: $($_.Exception.Message)" -LogLevel 3
        return $null
    }
}

Function Get-AuditInfo {
    <#
    .SYNOPSIS
    Returns a hashtable with Last Change By / Action for a given resource id.
    #>
    param ([string]$ResourceId)
    $result = @{}
    $audit  = Get-GraphUri -ApiVersion $script:graphApiVersion `
                           -Class "deviceManagement/auditEvents" `
                           -OData "?`$filter=resources/any(d:d/resourceId eq '$ResourceId')&`$top=1" `
                           -Value -AuditData
    if ($null -ne $audit) {
        $result["Last Change By"]     = Format-DataToString $audit.actor.userPrincipalName
        $result["Last Change Action"] = Format-DataToString $audit.activityOperationType
    }
    return $result
}

Function Add-AssignmentToDocument {
    <#
    .SYNOPSIS
    Resolves group assignments and appends bulleted lists to the Word document.
    Supports both intent-based (apps) and simple (policies) assignments.
    #>
    param (
        [Parameter(Mandatory)] $Assignments,
        [Parameter(Mandatory)] $Groups
    )

    if ($Assignments.Count -lt 1) { return }

    Add-WordText -WordDocument $WordDocument -Text 'This item has been assigned to the following groups:' -Supress $True

    $hasIntent = ($Assignments | Where-Object { $_.intent }) -ne $null

    if ($hasIntent) {
        $buckets = @{
            available                  = @()
            required                   = @()
            uninstall                  = @()
            availableWithoutEnrollment = @()
        }

        foreach ($a in $Assignments) {
            $groupName = if ($null -ne $a.target.groupId) {
                ($Groups | Where-Object { $_.id -eq $a.target.groupId }).displayName
            } else {
                switch ($a.target.'@odata.type') {
                    '#microsoft.graph.allLicensedUsersAssignmentTarget' { "All Users" }
                    '#microsoft.graph.allDevicesAssignmentTarget'       { "All Devices" }
                    default { $null }
                }
            }
            if ($null -ne $groupName -and $buckets.ContainsKey($a.intent)) {
                $buckets[$a.intent] += $groupName
            }
        }

        foreach ($intent in @('available','required','uninstall','availableWithoutEnrollment')) {
            # Filtrar nulls y vacios antes de Add-WordList
            $cleanBucket = @($buckets[$intent] | Where-Object { -not [string]::IsNullOrEmpty($_) })
            if ($cleanBucket.Count -ge 1) {
                Add-WordText -WordDocument $WordDocument -Text $intent -Supress $True
                Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $cleanBucket -Supress $True
            }
        }
    }
    else {
        $list = @()
        foreach ($a in $Assignments) {
            if ($null -ne $a.target.groupId) {
                $groupName = ($Groups | Where-Object { $_.id -eq $a.target.groupId }).displayName
                if (-not [string]::IsNullOrEmpty($groupName)) { $list += $groupName }
            } else {
                switch ($a.target.'@odata.type') {
                    '#microsoft.graph.allLicensedUsersAssignmentTarget' { $list += "All Users" }
                    '#microsoft.graph.allDevicesAssignmentTarget'       { $list += "All Devices" }
                }
            }
        }
        # Filtrar nulls y vacios antes de Add-WordList para evitar error en PSWriteWord
        $list = @($list | Where-Object { -not [string]::IsNullOrEmpty($_) })
        if ($list.Count -ge 1) {
            Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $list -Supress $True
        }
    }
}
#endregion

#region ---------------------------------------------------[Graph Class Processors]------------------------------------------------------------
Function Invoke-GraphClass {
    <#
    .SYNOPSIS
    Fetches a Graph class, exports JSON, and optionally writes a Word section.
    #>
    param (
        [Parameter(Mandatory)] [string]$Class,
        [Parameter(Mandatory)] [string]$Title,
        [Parameter()] [array]$Properties     = @(),
        [Parameter()] [string]$PropForFileName = "",
        [Parameter()] [switch]$Value,
        [Parameter()] [switch]$GetLastChange
    )

    $items = if ($Value) {
        Get-GraphUri -ApiVersion $script:graphApiVersion -Class $Class -Value
    } else {
        Get-GraphUri -ApiVersion $script:graphApiVersion -Class $Class
    }
    [array]$items = @($items)

    if ($Document) { Add-WordText -WordDocument $WordDocument -Text $Title -HeadingType Heading1 -Supress $True }

    foreach ($item in $items) {
        if ($null -eq $item) { continue }
        $classpath = $Class -replace "/", "\"

        $JSONFileName = $null
        if ($Export) {
            $fn = if ($PropForFileName -ne "") { Format-DataToString $item.$PropForFileName } else { "" }
            $JSONFileName = Export-JSONData -JSON $item -ExportPath "$ExportPath\$classpath" -FileName $fn -Force
        }

        if ($Document) {
            $src = if ($Properties.Count -gt 0) { $item | Select-Object -Property $Properties } else { $item }

            $ht = [ordered]@{}
            foreach ($prop in $src.psobject.properties) {
                $ht[(Format-PropertyName $prop.Name)] = Format-DataToString $prop.Value
            }

            if ($GetLastChange) {
                foreach ($kv in (Get-AuditInfo -ResourceId $item.id).GetEnumerator()) { $ht[$kv.Key] = $kv.Value }
            }

            $heading = if ($PropForFileName -ne "") { Format-DataToString $item.$PropForFileName } else { $item.displayName }
            Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
            Add-WordText -WordDocument $WordDocument -Text $heading -HeadingType Heading3 -Supress $True
            Add-WordTable -WordDocument $WordDocument -DataTable $ht -Design LightGridAccent1 -AutoFit Contents -Supress $True
        }
    }
}

Function Invoke-GraphClassExpand {
    <#
    .SYNOPSIS
    Like Invoke-GraphClass but also expands and documents group assignments.
    #>
    param (
        [Parameter(Mandatory)] [string]$Class,
        [Parameter(Mandatory)] [string]$Title,
        [Parameter()] [array]$Properties      = @(),
        [Parameter()] [string]$PropForFileName = "",
        [Parameter()] [switch]$Value,
        [Parameter()] [switch]$GetLastChange
    )

    $items = if ($Value) {
        Get-GraphUri -ApiVersion $script:graphApiVersion -Class $Class -Value
    } else {
        Get-GraphUri -ApiVersion $script:graphApiVersion -Class $Class
    }
    [array]$items = @($items)

    if ($Document) { Add-WordText -WordDocument $WordDocument -Text $Title -HeadingType Heading1 -Supress $True }

    foreach ($item in $items) {
        if ($null -eq $item) { continue }
        $classpath = $Class -replace "/", "\"

        $JSONFileName = $null
        if ($Export) {
            $fn = if ($PropForFileName -ne "") { Format-DataToString $item.$PropForFileName } else { "" }
            $JSONFileName = Export-JSONData -JSON $item -ExportPath "$ExportPath\$classpath" -FileName $fn -Force
        }

        if ($Document) {
            $src = if ($Properties.Count -gt 0) { $item | Select-Object -Property $Properties } else { $item }

            $ht = [ordered]@{}
            foreach ($prop in $src.psobject.properties) {
                $ht[(Format-PropertyName $prop.Name)] = Format-DataToString $prop.Value
            }

            if ($GetLastChange) {
                foreach ($kv in (Get-AuditInfo -ResourceId $item.id).GetEnumerator()) { $ht[$kv.Key] = $kv.Value }
            }

            $heading = if ($PropForFileName -ne "") { Format-DataToString $item.$PropForFileName } else { $item.displayName }
            Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
            Add-WordText -WordDocument $WordDocument -Text $heading -HeadingType Heading3 -Supress $True
            Add-WordTable -WordDocument $WordDocument -DataTable $ht -Design LightGridAccent1 -AutoFit Contents -Supress $True

            # Assignments
            $expanded = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $Class -Id $item.id -OData '?$expand=assignments'
            if ($expanded -and $expanded.assignments.Count -ge 1) {
                Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                Add-AssignmentToDocument -Assignments $expanded.assignments -Groups $Groups
            }
        }
    }
}
#endregion

#region ---------------------------------------------------[Execution]------------------------------------------------------------
Start-Log -FilePath $LogPath
Write-Log "---------- Script Starting ----------"

#region Load XML config
if (!(Test-Path $Config)) {
    Write-Log "Cannot find config file: $Config" -LogLevel 3
    Write-Error "Cannot find config file: $Config"
    Exit 1
}
try {
    $Xml = [xml](Get-Content -Path $Config -Encoding UTF8)
    Write-Log "Loaded config: $Config"
}
catch {
    Write-Log "Failed to read config: $($_.Exception.Message)" -LogLevel 3
    Exit 1
}

try {
    # Configuration section
    [bool]  $Document          = [System.Convert]::ToBoolean($Xml.root.Configuration.Document)
    [bool]  $Export            = [System.Convert]::ToBoolean($Xml.root.Configuration.Export)
    [bool]  $DocumentLastChange= [System.Convert]::ToBoolean($Xml.root.Configuration.DocumentLastChange)
    [bool]  $script:AppendDate = [System.Convert]::ToBoolean($Xml.root.Configuration.AppendDate)
    [bool]  $ExportCSV         = [System.Convert]::ToBoolean($Xml.root.Configuration.ExportCSV)
    [int]   $MaxStringLength   = [int]$Xml.root.Configuration.MaxStringLength
    [int]   $maxlogfilesize    = [int]$Xml.root.Configuration.maxlogfilesize * 1MB
    [string]$script:graphApiVersion = $Xml.root.Configuration.graphApiVersion
    # ClientId now lives in the XML — no more hardcoded value in the script
    [string]$script:ClientId   = $Xml.root.Configuration.ClientId

    if ([string]::IsNullOrEmpty($DocumentName)) { $DocumentName = $Xml.root.Configuration.DocumentName }
    # Reemplazar {ClientName} en el nombre del documento si existe
    if (-not [string]::IsNullOrEmpty($script:ClientName)) {
        $safeClientName = $script:ClientName -replace '[\/:*?"<>|]', "_"
        $DocumentName = $DocumentName -replace '\{ClientName\}', $safeClientName
    }
    if ([string]::IsNullOrEmpty($Tenant))       { $Tenant = $Xml.root.Configuration.Tenant }
    if ([string]::IsNullOrEmpty($ExportPath))   { $ExportPath = [string]$Xml.root.Configuration.ExportPath -f (Get-Date -Format "yyyyMMddHHmm") }

    # Process flags
    [bool]$ProcessmanagedDeviceOverview          = [System.Convert]::ToBoolean($Xml.root.Process.managedDeviceOverview)
    [bool]$ProcesstermsAndConditions             = [System.Convert]::ToBoolean($Xml.root.Process.termsAndConditions)
    [bool]$ProcessdeviceCompliancePolicies       = [System.Convert]::ToBoolean($Xml.root.Process.deviceCompliancePolicies)
    [bool]$ProcessdeviceEnrollmentConfigurations = [System.Convert]::ToBoolean($Xml.root.Process.deviceEnrollmentConfigurations)
    [bool]$ProcessdeviceConfigurations           = [System.Convert]::ToBoolean($Xml.root.Process.deviceConfigurations)
    [bool]$ProcesswindowsAutopilotDeploymentProfiles = [System.Convert]::ToBoolean($Xml.root.Process.windowsAutopilotDeploymentProfiles)
    [bool]$ProcessmobileApps                    = [System.Convert]::ToBoolean($Xml.root.Process.mobileApps)
    [bool]$ProcessSfBApps                       = [System.Convert]::ToBoolean($Xml.root.Process.SfBApps)
    [bool]$ProcessapplePushNotificationCertificate = [System.Convert]::ToBoolean($Xml.root.Process.applePushNotificationCertificate)
    [bool]$ProcessvppTokens                     = [System.Convert]::ToBoolean($Xml.root.Process.vppTokens)
    [bool]$Processpolicysets                    = [System.Convert]::ToBoolean($Xml.root.Process.policysets)
    [bool]$ProcessgroupPolicyConfigurations     = [System.Convert]::ToBoolean($Xml.root.Process.groupPolicyConfigurations)
    [bool]$ProcessdeviceManagementScripts       = [System.Convert]::ToBoolean($Xml.root.Process.deviceManagementScripts)
    [bool]$ProcessGroups                        = [System.Convert]::ToBoolean($Xml.root.Process.Groups)

    # ---- Nuevas secciones opcionales ----
    [bool]$ProcessconditionalAccess         = if ($Xml.root.Process.conditionalAccess)              { [System.Convert]::ToBoolean($Xml.root.Process.conditionalAccess) }              else { $false }
    [bool]$ProcessroleDefinitions           = if ($Xml.root.Process.roleDefinitions)               { [System.Convert]::ToBoolean($Xml.root.Process.roleDefinitions) }               else { $false }
    [bool]$ProcessroleAssignments           = if ($Xml.root.Process.roleAssignments)               { [System.Convert]::ToBoolean($Xml.root.Process.roleAssignments) }               else { $false }
    [bool]$ProcessdeviceCategories          = if ($Xml.root.Process.deviceCategories)              { [System.Convert]::ToBoolean($Xml.root.Process.deviceCategories) }              else { $false }
    [bool]$ProcesshealthScripts             = if ($Xml.root.Process.deviceHealthScripts)           { [System.Convert]::ToBoolean($Xml.root.Process.deviceHealthScripts) }           else { $false }
    [bool]$ProcesswindowsQualityUpdates     = if ($Xml.root.Process.windowsQualityUpdateProfiles)  { [System.Convert]::ToBoolean($Xml.root.Process.windowsQualityUpdateProfiles) }  else { $false }
    [bool]$ProcesswindowsFeatureUpdates     = if ($Xml.root.Process.windowsFeatureUpdateProfiles)  { [System.Convert]::ToBoolean($Xml.root.Process.windowsFeatureUpdateProfiles) }  else { $false }
    [bool]$ProcessnotificationTemplates     = if ($Xml.root.Process.notificationMessageTemplates)  { [System.Convert]::ToBoolean($Xml.root.Process.notificationMessageTemplates) }  else { $false }
    [bool]$ProcessscopeTags                 = if ($Xml.root.Process.roleScopeTags)                 { [System.Convert]::ToBoolean($Xml.root.Process.roleScopeTags) }                 else { $false }
    [bool]$ProcessmanagedDevices            = if ($Xml.root.Process.managedDevices)                { [System.Convert]::ToBoolean($Xml.root.Process.managedDevices) }                 else { $false }
    [bool]$ProcesssettingsCatalog           = if ($Xml.root.Process.settingsCatalog)               { [System.Convert]::ToBoolean($Xml.root.Process.settingsCatalog) }               else { $false }
    [bool]$ProcessappProtectionPolicies     = if ($Xml.root.Process.appProtectionPolicies)         { [System.Convert]::ToBoolean($Xml.root.Process.appProtectionPolicies) }         else { $false }
    [bool]$ProcessassignmentFilters         = if ($Xml.root.Process.assignmentFilters)             { [System.Convert]::ToBoolean($Xml.root.Process.assignmentFilters) }             else { $false }
    [bool]$ProcessendpointSecurity          = if ($Xml.root.Process.endpointSecurity)              { [System.Convert]::ToBoolean($Xml.root.Process.endpointSecurity) }              else { $false }
    [bool]$ProcessstaleDevices              = if ($Xml.root.Process.staleDevices)                  { [System.Convert]::ToBoolean($Xml.root.Process.staleDevices) }                  else { $false }
    [bool]$ProcessnonCompliantDevices       = if ($Xml.root.Process.nonCompliantDevices)           { [System.Convert]::ToBoolean($Xml.root.Process.nonCompliantDevices) }           else { $false }
    [bool]$ProcessosVersionSummary          = if ($Xml.root.Process.osVersionSummary)              { [System.Convert]::ToBoolean($Xml.root.Process.osVersionSummary) }              else { $false }

    Write-Log "All config settings loaded successfully."
}
catch {
    Write-Log "Failed to parse config XML: $($_.Exception.Message)" -LogLevel 3
    Exit 1
}
#endregion Load XML config

$script:Tenant = $Tenant

if ($Export -eq $false -and $Document -eq $false) {
    Write-Error "At least one of Document or Export must be True in the config file."
    Write-Log   "Neither Document nor Export is enabled — aborting." -LogLevel 3
    Exit 1
}

if ([string]::IsNullOrEmpty($script:ClientId)) {
    Write-Error "ClientId is not set in the config XML (<ClientId>). Please add it and retry."
    Write-Log   "ClientId missing from config XML." -LogLevel 3
    Exit 1
}

#region Connect
Write-Host "Connecting to tenant: $script:Tenant — continue? (Y/N)" -ForegroundColor Yellow
if ((Read-Host) -notin 'y','Y') {
    Write-Log "User aborted at tenant confirmation." -LogLevel 3
    Exit 1
}
$script:Tenant = $Tenant
Update-AuthToken
Write-Host "Connected as: $script:User" -ForegroundColor Green
Write-Log  "Connected as: $script:User"
#endregion Connect

#region Export path
if ([string]::IsNullOrEmpty($ExportPath)) {
    $ExportPath = Read-Host -Prompt "Export path (e.g. C:\MEMOutput)"
}
$ExportPath = $ExportPath.Trim('"')

# Client name for cover page
if ([string]::IsNullOrEmpty($ClientName)) {
    $ClientName = Read-Host -Prompt "Client name for the cover page (e.g. Contoso Inc.)"
}
$script:ClientName = $ClientName
Write-Log "Client name set to: $($script:ClientName)"

if (!(Test-Path $ExportPath)) {
    if (!$Force) {
        Write-Host "Path '$ExportPath' doesn't exist. Create it? (Y/N)" -ForegroundColor Yellow
        if ((Read-Host) -notin 'y','Y') {
            Write-Log "Directory creation cancelled." -LogLevel 3
            Exit 1
        }
    }
    New-Item -ItemType Directory -Path $ExportPath | Out-Null
}
#endregion Export path

$FullDocumentationPath = Join-Path $ExportPath $DocumentName
Write-Log "Template        : $DocumentTemplate"
Write-Log "Export path     : $ExportPath"
Write-Log "Documentation   : $FullDocumentationPath"

#region Word setup
if ($Document) {
    $WordDocument = Get-WordDocument -FilePath $DocumentTemplate

    # Resolver el nombre del usuario si aun esta vacio
    if ([string]::IsNullOrEmpty($script:User)) {
        try {
            $meUri  = "https://graph.microsoft.com/v1.0/me"
            $meData = Invoke-RestMethod -Uri $meUri -Headers $global:authToken -Method Get
            $script:User = $meData.userPrincipalName
        } catch { $script:User = "IntuneDocKit" }
    }

    foreach ($para in $WordDocument.Paragraphs) {
        $para.ReplaceText('#DATE#',       (Get-Date -Format "yyyy.MM.dd HH:mm"))
        $para.ReplaceText('#TENANT#',     $script:Tenant)
        $para.ReplaceText('#USERNAME#',   $script:User)
        $para.ReplaceText('#CLIENTNAME#', $script:ClientName)
    }
    Add-WordPageBreak -WordDocument $WordDocument -Supress $True
    Add-WordTOC      -WordDocument $WordDocument -Title 'Table of Contents' -HeaderStyle Heading1 -Supress $True
    Add-WordSection  -WordDocument $WordDocument -PageBreak -Supress $True
}
#endregion Word setup

# Pre-load all groups once (used for resolving assignments throughout)
Write-Log "Loading all groups..."
$Groups = Get-GraphUri -ApiVersion $script:graphApiVersion -Class "groups" -Value

#region managedDeviceOverview
if ($ProcessmanagedDeviceOverview -and $Document) {
    $overview = Get-GraphUri -ApiVersion $script:graphApiVersion -Class "deviceManagement/managedDeviceOverview"

    $dtHt = [ordered]@{}
    foreach ($p in ($overview | Select-Object enrolledDeviceCount, mdmEnrolledCount, dualEnrolledDeviceCount, managedDeviceModelsAndManufacturers, lastModifiedDateTime).psobject.properties) {
        $dtHt[(Format-DataToString $p.Name)] = Format-DataToString $p.Value
    }
    $osHt = [ordered]@{}
    foreach ($p in $overview.deviceOperatingSystemSummary.psobject.properties) {
        $osHt[(Format-DataToString $p.Name)] = Format-DataToString $p.Value
    }
    $eaHt = [ordered]@{}
    foreach ($p in $overview.deviceExchangeAccessStateSummary.psobject.properties) {
        $eaHt[(Format-DataToString $p.Name)] = Format-DataToString $p.Value
    }

    Add-WordText     -WordDocument $WordDocument -Text 'Device Overview' -HeadingType Heading1 -Supress $True
    Add-WordText     -WordDocument $WordDocument -Text '' -Supress $True
    Add-WordTable    -WordDocument $WordDocument -DataTable $dtHt -Design LightGridAccent1 -AutoFit Window -Supress $True
    Add-WordText     -WordDocument $WordDocument -Text '' -Supress $True
    Add-WordPieChart -WordDocument $WordDocument -ChartName 'Operating System Summary'      -Names $osHt.Keys -Values $osHt.Values
  #  Add-WordText     -WordDocument $WordDocument -Text '' -Supress $True
  #  Add-WordPieChart -WordDocument $WordDocument -ChartName 'Exchange Access State Summary' -Names $eaHt.Keys -Values $eaHt.Values
  #  Add-WordText     -WordDocument $WordDocument -Text '' -Supress $True
}
#endregion

#region managedDevices - Device Inventory
if ($ProcessmanagedDevices) {
    $selectFields = "deviceName,userPrincipalName,operatingSystem,osVersion,complianceState," +
                    "lastSyncDateTime,manufacturer,model,serialNumber,isEncrypted," +
                    "managedDeviceOwnerType,enrolledDateTime"

    [array]$allDevices = @(Get-GraphUri -ApiVersion $script:graphApiVersion `
        -Class "deviceManagement/managedDevices" `
        -OData "?`$select=$selectFields" `
        -Value)

    if ($Export) {
        foreach ($dev in $allDevices) {
            if ($null -ne $dev) {
                Export-JSONData -JSON $dev `
                    -ExportPath "$ExportPath\deviceManagement\managedDevices" `
                    -FileName ($dev.deviceName -replace '[\/:*?"<>|]', "_") `
                    -Force | Out-Null
            }
        }
    }

    if ($Document) {
        Add-WordText -WordDocument $WordDocument -Text 'Device Inventory' -HeadingType Heading1 -Supress $True
        Add-WordText -WordDocument $WordDocument -Text '' -Supress $True

        # --- Summary table ---
        $totalDevices  = $allDevices.Count
        $compliantCnt  = ($allDevices | Where-Object { $_.complianceState -eq 'compliant' }).Count
        $nonCompCnt    = ($allDevices | Where-Object { $_.complianceState -eq 'noncompliant' }).Count
        $unknownCnt    = $totalDevices - $compliantCnt - $nonCompCnt

        $summaryHt = [ordered]@{
            'Total Managed Devices' = [string]$totalDevices
            'Compliant'             = [string]$compliantCnt
            'Non-Compliant'         = [string]$nonCompCnt
            'Unknown / Other'       = [string]$unknownCnt
        }
        Add-WordTable -WordDocument $WordDocument -DataTable $summaryHt -Design LightGridAccent1 -AutoFit Contents -Supress $True
        Add-WordText -WordDocument $WordDocument -Text '' -Supress $True

        # --- Device list table ---
        $deviceTable = $allDevices | Where-Object { $null -ne $_ } | Sort-Object deviceName |
            Select-Object `
                @{ Name='Device Name'; Expression={ Format-DataToString $_.deviceName } },
                @{ Name='User (UPN)';  Expression={ Format-DataToString $_.userPrincipalName } },
                @{ Name='OS';          Expression={ Format-DataToString $_.operatingSystem } },
                @{ Name='Version';     Expression={ Format-DataToString $_.osVersion } },
                @{ Name='Status';      Expression={ Format-DataToString $_.complianceState } },
                @{ Name='Last Sync';   Expression={
                    $raw = $_.lastSyncDateTime
                    if ($raw) {
                        try { ([DateTime]::Parse($raw)).ToString('yyyy-MM-dd') }
                        catch { Format-DataToString $raw }
                    } else { '' }
                }},
                @{ Name='Encrypted';   Expression={ if ($_.isEncrypted) { 'Yes' } else { 'No' } } }

        if ($deviceTable.Count -ge 1) {
            Add-WordTable -WordDocument $WordDocument -DataTable $deviceTable -Design LightGridAccent1 -AutoFit Window -Supress $True
        }
    }
}
#endregion managedDevices

#region staleDevices
if ($ProcessstaleDevices) {
    $staleSelect = "deviceName,userPrincipalName,operatingSystem,osVersion,lastSyncDateTime"
    [array]$allForStale = @(Get-GraphUri -ApiVersion $script:graphApiVersion `
        -Class "deviceManagement/managedDevices" `
        -OData "?`$select=$staleSelect" `
        -Value)

    $cutoff = (Get-Date).AddDays(-30)
    [array]$staleList = @($allForStale | Where-Object {
        $null -ne $_ -and
        -not [string]::IsNullOrEmpty($_.lastSyncDateTime) -and
        ([DateTime]::Parse($_.lastSyncDateTime)) -le $cutoff
    })

    if ($Document) {
        Add-WordText -WordDocument $WordDocument -Text 'Stale Devices (No Sync > 30 Days)' -HeadingType Heading1 -Supress $True
        Add-WordText -WordDocument $WordDocument -Text '' -Supress $True

        if ($staleList.Count -eq 0) {
            Add-WordTable -WordDocument $WordDocument `
                -DataTable ([ordered]@{ 'Result' = 'No stale devices found.' }) `
                -Design LightGridAccent1 -AutoFit Contents -Supress $True
        } else {
            Add-WordTable -WordDocument $WordDocument `
                -DataTable ([ordered]@{ 'Stale Devices (> 30 days without sync)' = [string]$staleList.Count }) `
                -Design LightGridAccent1 -AutoFit Contents -Supress $True
            Add-WordText -WordDocument $WordDocument -Text '' -Supress $True

            $staleTable = $staleList | Sort-Object lastSyncDateTime |
                Select-Object `
                    @{ Name='Device Name'; Expression={ Format-DataToString $_.deviceName } },
                    @{ Name='User (UPN)';  Expression={ Format-DataToString $_.userPrincipalName } },
                    @{ Name='OS';          Expression={ Format-DataToString $_.operatingSystem } },
                    @{ Name='Version';     Expression={ Format-DataToString $_.osVersion } },
                    @{ Name='Last Sync';   Expression={
                        $raw = $_.lastSyncDateTime
                        if ($raw) { try { ([DateTime]::Parse($raw)).ToString('yyyy-MM-dd') } catch { Format-DataToString $raw } } else { '' }
                    }}
            Add-WordTable -WordDocument $WordDocument -DataTable $staleTable -Design LightGridAccent1 -AutoFit Window -Supress $True
        }
    }
}
#endregion staleDevices

#region nonCompliantDevices
if ($ProcessnonCompliantDevices) {
    $ncSelect = "deviceName,userPrincipalName,operatingSystem,osVersion,complianceState,lastSyncDateTime"
    [array]$nonCompliant = @(Get-GraphUri -ApiVersion $script:graphApiVersion `
        -Class "deviceManagement/managedDevices" `
        -OData "?`$filter=complianceState eq 'noncompliant'&`$select=$ncSelect" `
        -Value)

    if ($Document) {
        Add-WordText -WordDocument $WordDocument -Text 'Non-Compliant Devices' -HeadingType Heading1 -Supress $True
        Add-WordText -WordDocument $WordDocument -Text '' -Supress $True

        if ($nonCompliant.Count -eq 0) {
            Add-WordTable -WordDocument $WordDocument `
                -DataTable ([ordered]@{ 'Result' = 'No non-compliant devices found.' }) `
                -Design LightGridAccent1 -AutoFit Contents -Supress $True
        } else {
            Add-WordTable -WordDocument $WordDocument `
                -DataTable ([ordered]@{ 'Non-Compliant Devices' = [string]$nonCompliant.Count }) `
                -Design LightGridAccent1 -AutoFit Contents -Supress $True
            Add-WordText -WordDocument $WordDocument -Text '' -Supress $True

            $ncTable = $nonCompliant | Where-Object { $null -ne $_ } | Sort-Object deviceName |
                Select-Object `
                    @{ Name='Device Name'; Expression={ Format-DataToString $_.deviceName } },
                    @{ Name='User (UPN)';  Expression={ Format-DataToString $_.userPrincipalName } },
                    @{ Name='OS';          Expression={ Format-DataToString $_.operatingSystem } },
                    @{ Name='Version';     Expression={ Format-DataToString $_.osVersion } },
                    @{ Name='Last Sync';   Expression={
                        $raw = $_.lastSyncDateTime
                        if ($raw) { try { ([DateTime]::Parse($raw)).ToString('yyyy-MM-dd') } catch { Format-DataToString $raw } } else { '' }
                    }}
            Add-WordTable -WordDocument $WordDocument -DataTable $ncTable -Design LightGridAccent1 -AutoFit Window -Supress $True
        }
    }
}
#endregion nonCompliantDevices

#region osVersionSummary
if ($ProcessosVersionSummary) {
    $osSelect = "operatingSystem,osVersion"
    [array]$allForOS = @(Get-GraphUri -ApiVersion $script:graphApiVersion `
        -Class "deviceManagement/managedDevices" `
        -OData "?`$select=$osSelect" `
        -Value)

    if ($Document) {
        Add-WordText -WordDocument $WordDocument -Text 'OS Version Summary' -HeadingType Heading1 -Supress $True
        Add-WordText -WordDocument $WordDocument -Text '' -Supress $True

        $grouped = $allForOS | Where-Object { $null -ne $_ } |
            Group-Object operatingSystem, osVersion | Sort-Object Name |
            Select-Object `
                @{ Name='OS';      Expression={ ($_.Name -split ', ')[0] } },
                @{ Name='Version'; Expression={ ($_.Name -split ', ')[1] } },
                @{ Name='Devices'; Expression={ [string]$_.Count } }

        if ($grouped.Count -ge 1) {
            Add-WordTable -WordDocument $WordDocument -DataTable $grouped -Design LightGridAccent1 -AutoFit Contents -Supress $True
        }
    }
}
#endregion osVersionSummary

#region Standard classes
if ($ProcesstermsAndConditions)             { Invoke-GraphClass        -Class "deviceManagement/termsAndConditions"                -Title 'Terms and Conditions'                  -PropForFileName "@odata.type" -Value }
if ($ProcessdeviceCompliancePolicies)       { Invoke-GraphClassExpand  -Class "deviceManagement/deviceCompliancePolicies"          -Title 'Device Compliance Policies'             -Value -GetLastChange:$DocumentLastChange }
if ($ProcessdeviceEnrollmentConfigurations) { Invoke-GraphClass        -Class "deviceManagement/deviceEnrollmentConfigurations"    -Title 'Device Enrollment Configurations'       -PropForFileName "@odata.type" -Value -GetLastChange:$DocumentLastChange }
if ($ProcessdeviceConfigurations)           { Invoke-GraphClassExpand  -Class "deviceManagement/deviceConfigurations"              -Title 'Device Configurations'                  -Properties "displayName","id","lastModifiedDateTime","description" -Value -GetLastChange:$DocumentLastChange }
if ($ProcesswindowsAutopilotDeploymentProfiles) { Invoke-GraphClassExpand -Class "deviceManagement/windowsAutopilotDeploymentProfiles" -Title 'Windows Autopilot Deployment Profiles' -Value -GetLastChange:$DocumentLastChange }
if ($ProcessapplePushNotificationCertificate)   { Invoke-GraphClass    -Class "deviceManagement/applePushNotificationCertificate" -Title 'Apple Push Notification Certificate'    -PropForFileName "@odata.type" -Value }
if ($ProcessvppTokens)                      { Invoke-GraphClass        -Class "deviceAppManagement/vppTokens"                     -Title 'VPP Tokens'                             -Value }
#endregion

#region settingsCatalog
if ($ProcesssettingsCatalog) {
    $class = "deviceManagement/configurationPolicies"
    [array]$scPolicies = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Value
    $classpath = $class -replace "/", "\"

    if ($Document) { Add-WordText -WordDocument $WordDocument -Text 'Settings Catalog' -HeadingType Heading1 -Supress $True }

    foreach ($policy in $scPolicies) {
        if ($null -eq $policy) { continue }

        if ($Export) {
            Export-JSONData -JSON $policy -ExportPath "$ExportPath\$classpath" `
                -FileName ($policy.name -replace '[\/:*?"<>|]', "_") -Force | Out-Null
        }

        if ($Document) {
            $ht = [ordered]@{
                'Name'          = Format-DataToString $policy.name
                'Description'   = Format-DataToString $policy.description
                'Platforms'     = Format-DataToString $policy.platforms
                'Technologies'  = Format-DataToString $policy.technologies
                'Setting Count' = Format-DataToString $policy.settingCount
                'Last Modified' = Format-DataToString $policy.lastModifiedDateTime
                'Created'       = Format-DataToString $policy.createdDateTime
            }
            Add-WordText  -WordDocument $WordDocument -Text '' -Supress $True
            Add-WordText  -WordDocument $WordDocument -Text (Format-DataToString $policy.name) -HeadingType Heading3 -Supress $True
            Add-WordTable -WordDocument $WordDocument -DataTable $ht -Design LightGridAccent1 -AutoFit Contents -Supress $True

            $expanded = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Id $policy.id -OData '?$expand=assignments'
            if ($expanded -and $expanded.assignments.Count -ge 1) {
                Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                Add-AssignmentToDocument -Assignments $expanded.assignments -Groups $Groups
            }
        }
    }
}
#endregion settingsCatalog

#region mobileApps
if ($ProcessmobileApps) {
    $class = "deviceAppManagement/mobileApps"
    [array]$apps = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Value

    if ($Document) { Add-WordText -WordDocument $WordDocument -Text 'Applications' -HeadingType Heading1 -Supress $True }

    foreach ($app in $apps) {
        if ($app.appAvailability -eq "global") { continue }
        if ($app.'@odata.type' -eq "#microsoft.graph.microsoftStoreForBusinessApp" -and !$ProcessSfBApps) { continue }

        $classpath   = $class -replace "/", "\"
        $expandedApp = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Id $app.id -OData '?$expand=assignments'

        $JSONFileName = $null
        if ($Export) { $JSONFileName = Export-JSONData -JSON $expandedApp -ExportPath "$ExportPath\$classpath" -Force }

        if ($Document) {
            $ht = [ordered]@{}
            foreach ($p in $app.psobject.properties) { $ht[(Format-PropertyName $p.Name)] = Format-DataToString $p.Value }
            foreach ($kv in (Get-AuditInfo -ResourceId $app.id).GetEnumerator()) { $ht[$kv.Key] = $kv.Value }

            Add-WordText  -WordDocument $WordDocument -Text '' -Supress $True
            Add-WordText  -WordDocument $WordDocument -Text $app.displayName -HeadingType Heading3 -Supress $True
            Add-WordTable -WordDocument $WordDocument -DataTable $ht -Design LightGridAccent1 -AutoFit Contents -Supress $True

            if ($expandedApp.assignments.Count -ge 1) {
                Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                Add-AssignmentToDocument -Assignments $expandedApp.assignments -Groups $Groups
            }
        }
    }
}
#endregion mobileApps

#region appProtectionPolicies
if ($ProcessappProtectionPolicies) {
    $class = "deviceAppManagement/managedAppPolicies"
    [array]$mamPolicies = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Value
    $classpath = $class -replace "/", "\"

    if ($Document) { Add-WordText -WordDocument $WordDocument -Text 'App Protection Policies (MAM)' -HeadingType Heading1 -Supress $True }

    foreach ($policy in $mamPolicies) {
        if ($null -eq $policy) { continue }

        if ($Export) {
            Export-JSONData -JSON $policy -ExportPath "$ExportPath\$classpath" `
                -FileName ($policy.displayName -replace '[\/:*?"<>|]', "_") -Force | Out-Null
        }

        if ($Document) {
            $policyType = ($policy.'@odata.type' -replace '#microsoft\.graph\.', '')
            $ht = [ordered]@{
                'Display Name'  = Format-DataToString $policy.displayName
                'Policy Type'   = Format-DataToString $policyType
                'Description'   = Format-DataToString $policy.description
                'Last Modified' = Format-DataToString $policy.lastModifiedDateTime
                'Created'       = Format-DataToString $policy.createdDateTime
            }
            Add-WordText  -WordDocument $WordDocument -Text '' -Supress $True
            Add-WordText  -WordDocument $WordDocument -Text (Format-DataToString $policy.displayName) -HeadingType Heading3 -Supress $True
            Add-WordTable -WordDocument $WordDocument -DataTable $ht -Design LightGridAccent1 -AutoFit Contents -Supress $True
        }
    }
}
#endregion appProtectionPolicies

#region policySets
if ($Processpolicysets) {
    $class = "deviceAppManagement/policysets"
    [array]$sets = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Value

    if ($Document) { Add-WordText -WordDocument $WordDocument -Text 'Policy Sets' -HeadingType Heading1 -Supress $True }

    foreach ($set in $sets) {
        $classpath   = $class -replace "/", "\"
        $expandedSet = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Id $set.id -OData '?$expand=assignments,items'

        $JSONFileName = $null
        if ($Export) { $JSONFileName = Export-JSONData -JSON $expandedSet -ExportPath "$ExportPath\$classpath" -Force }

        if ($Document) {
            $ht = [ordered]@{}
            foreach ($p in $set.psobject.properties) { $ht[(Format-PropertyName $p.Name)] = Format-DataToString $p.Value }
            foreach ($kv in (Get-AuditInfo -ResourceId $set.id).GetEnumerator()) { $ht[$kv.Key] = $kv.Value }

            Add-WordText  -WordDocument $WordDocument -Text '' -Supress $True
            Add-WordText  -WordDocument $WordDocument -Text $set.displayName -HeadingType Heading3 -Supress $True
            Add-WordTable -WordDocument $WordDocument -DataTable $ht -Design LightGridAccent1 -AutoFit Contents -Supress $True

            if ($expandedSet.items.Count -ge 1) {
                Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                Add-WordText -WordDocument $WordDocument -Text 'This set includes the following items:' -Supress $True
                Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $expandedSet.items.displayName -Supress $True
            }

            if ($expandedSet.assignments.Count -ge 1) {
                Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                Add-AssignmentToDocument -Assignments $expandedSet.assignments -Groups $Groups
            }
        }
    }
}
#endregion policySets

#region groupPolicyConfigurations
if ($ProcessgroupPolicyConfigurations) {
    $class = "deviceManagement/groupPolicyConfigurations"
    [array]$gpcs = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Value
    $classpath    = $class -replace "/", "\"

    if ($Document) { Add-WordText -WordDocument $WordDocument -Text 'Group Policy Configurations' -HeadingType Heading1 -Supress $True }

    foreach ($gpc in $gpcs) {
        $expandedGpc = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Id $gpc.id -OData '?$expand=assignments'
        $folderName  = $gpc.displayName -replace '\<|\>|:|"|/|\\|\||\?|\*', "_"
        $gpcFiles    = @()

        $JSONFileName = $null
        if ($Export) { $JSONFileName = Export-JSONData -JSON $expandedGpc -ExportPath "$ExportPath\$classpath" -FileName "$($gpc.displayName)_assignments" -Force }

        if ($Document) { Add-WordText -WordDocument $WordDocument -Text $gpc.displayName -HeadingType Heading3 -Supress $True }

        $ht         = [ordered]@{}
        $defValues  = Get-GraphUri -ApiVersion $script:graphApiVersion -Class "deviceManagement/groupPolicyConfigurations/$($gpc.id)/definitionValues" -Value

        foreach ($dv in $defValues) {
            $pvArr     = @(Get-GraphUri -ApiVersion $script:graphApiVersion -Class "deviceManagement/groupPolicyConfigurations/$($gpc.id)/definitionValues/$($dv.id)/presentationValues" -Value)
            $definition= Get-GraphUri -ApiVersion $script:graphApiVersion -Class "deviceManagement/groupPolicyConfigurations/$($gpc.id)/definitionValues/$($dv.id)/definition"
            $value     = $null

            if ($Export) {
                $expObj = [PSCustomObject]@{
                    "definition@odata.bind" = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyDefinitions('$($definition.id)')"
                    "enabled"               = $dv.enabled.ToString().ToLower()
                }
            }

            if ($pvArr) {
                $presentationValues = @()
                foreach ($pv in $pvArr) {
                    $presentation = Get-GraphUri -ApiVersion $script:graphApiVersion -Class "deviceManagement/groupPolicyConfigurations/$($gpc.id)/definitionValues/$($dv.id)/presentationValues/$($pv.id)/presentation"

                    if ($Export) {
                        $pvObj = $pv | Select-Object -Property * -ExcludeProperty id, createdDateTime, lastModifiedDateTime, version
                        $pvObj | Add-Member -MemberType NoteProperty -Name "presentation@odata.bind" -Value "https://graph.microsoft.com/beta/deviceManagement/groupPolicyDefinitions('$($definition.id)')/presentations('$($presentation.id)')"
                        $presentationValues += $pvObj
                    }

                    $value = if ($null -ne $pv.values) { $pv.values } elseif ($null -ne $pv.value) { $pv.value } else { $null }
                    if ($null -ne $presentation) {
                        $ht[(Format-DataToString $presentation.label)] = Format-DataToString $value
                    }
                }
                if ($Export) {
                    $expObj | Add-Member -MemberType NoteProperty -Name "presentationValues" -Value $presentationValues
                    $gpcFiles += Export-JSONData -JSON $expObj -ExportPath "$ExportPath\$classpath\$folderName" -FileName $definition.displayName -Force
                }
            }
            else {
                if ($null -ne $definition) {
                    if ($Export) { $gpcFiles += Export-JSONData -JSON $expObj -ExportPath "$ExportPath\$classpath\$folderName" -FileName $definition.displayName -Force }
                    $ht[(Format-DataToString $definition.displayName)] = Format-DataToString "enabled=$($dv.enabled)"
                }
            }
        }

        if ($Document) {
            $auditInfo = Get-AuditInfo -ResourceId $gpc.id
            if ($null -ne $auditInfo) {
                foreach ($kv in $auditInfo.GetEnumerator()) { $ht[$kv.Key] = $kv.Value }
            }
            if ($ht.Count -gt 0) {
                Add-WordTable -WordDocument $WordDocument -DataTable $ht -Design LightGridAccent1 -AutoFit Contents -Supress $True
            }

            if ($expandedGpc.assignments.Count -ge 1) {
                Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                Add-AssignmentToDocument -Assignments $expandedGpc.assignments -Groups $Groups
            }
        }
    }
}
#endregion groupPolicyConfigurations

#region deviceManagementScripts
if ($ProcessdeviceManagementScripts) {
    $class = "deviceManagement/deviceManagementScripts"
    [array]$scripts = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Value

    if ($Document) { Add-WordText -WordDocument $WordDocument -Text 'Device Management Scripts' -HeadingType Heading1 -Supress $True }

    foreach ($scr in $scripts) {
        $classpath     = $class -replace "/", "\"
        $expandedScr   = Get-GraphUri -ApiVersion $script:graphApiVersion -Class $class -Id $scr.id -OData '?$expand=assignments'

        $JSONFileName = $null
        if ($Export) {
            $JSONFileName = Export-JSONData -JSON $scr -ExportPath "$ExportPath\$classpath" -Force
            $scriptContent= [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($expandedScr.scriptContent))
            $scriptContent | Out-File -FilePath (Join-Path "$ExportPath\$classpath" $scr.fileName) -Encoding UTF8
        }

        if ($Document) {
            $ht = [ordered]@{}
            foreach ($p in ($scr | Select-Object -ExcludeProperty scriptContent).psobject.properties) {
                $ht[(Format-PropertyName $p.Name)] = Format-DataToString $p.Value
            }
            foreach ($kv in (Get-AuditInfo -ResourceId $scr.id).GetEnumerator()) { $ht[$kv.Key] = $kv.Value }

            Add-WordText  -WordDocument $WordDocument -Text '' -Supress $True
            Add-WordText  -WordDocument $WordDocument -Text $scr.displayName -HeadingType Heading3 -Supress $True
            Add-WordTable -WordDocument $WordDocument -DataTable $ht -Design LightGridAccent1 -AutoFit Contents -Supress $True

            if ($expandedScr.assignments.Count -ge 1) {
                Add-WordText -WordDocument $WordDocument -Text '' -Supress $True
                Add-AssignmentToDocument -Assignments $expandedScr.assignments -Groups $Groups
            }
        }
    }
}
#endregion deviceManagementScripts

#region conditionalAccess
if ($ProcessconditionalAccess) {
    Invoke-GraphClass -Class "identity/conditionalAccess/policies" `
        -Title 'Conditional Access Policies' `
        -Properties "displayName","state","createdDateTime","modifiedDateTime","id" `
        -PropForFileName "displayName" -Value
}
#endregion conditionalAccess

#region roleDefinitions
if ($ProcessroleDefinitions) {
    Invoke-GraphClass -Class "deviceManagement/roleDefinitions" `
        -Title 'Role Definitions (RBAC)' `
        -PropForFileName "displayName" -Value
}
#endregion roleDefinitions

#region roleAssignments
if ($ProcessroleAssignments) {
    Invoke-GraphClass -Class "deviceManagement/roleAssignments" `
        -Title 'Role Assignments (RBAC)' `
        -PropForFileName "displayName" -Value
}
#endregion roleAssignments

#region deviceCategories
if ($ProcessdeviceCategories) {
    Invoke-GraphClass -Class "deviceManagement/deviceCategories" `
        -Title 'Device Categories' `
        -PropForFileName "displayName" -Value
}
#endregion deviceCategories

#region assignmentFilters
if ($ProcessassignmentFilters) {
    Invoke-GraphClass -Class "deviceManagement/assignmentFilters" `
        -Title 'Assignment Filters' `
        -Properties "displayName","description","platform","rule","createdDateTime","lastModifiedDateTime","id" `
        -PropForFileName "displayName" -Value
}
#endregion assignmentFilters

#region deviceHealthScripts
if ($ProcesshealthScripts) {
    Invoke-GraphClassExpand -Class "deviceManagement/deviceHealthScripts" `
        -Title 'Device Health Scripts (Endpoint Analytics)' `
        -PropForFileName "displayName" -Value -GetLastChange:$DocumentLastChange
}
#endregion deviceHealthScripts

#region windowsQualityUpdateProfiles
if ($ProcesswindowsQualityUpdates) {
    Invoke-GraphClassExpand -Class "deviceManagement/windowsQualityUpdateProfiles" `
        -Title 'Windows Quality Update Profiles' `
        -PropForFileName "displayName" -Value
}
#endregion windowsQualityUpdateProfiles

#region windowsFeatureUpdateProfiles
if ($ProcesswindowsFeatureUpdates) {
    Invoke-GraphClassExpand -Class "deviceManagement/windowsFeatureUpdateProfiles" `
        -Title 'Windows Feature Update Profiles' `
        -PropForFileName "displayName" -Value
}
#endregion windowsFeatureUpdateProfiles

#region notificationMessageTemplates
if ($ProcessnotificationTemplates) {
    Invoke-GraphClass -Class "deviceManagement/notificationMessageTemplates" `
        -Title 'Notification Message Templates' `
        -PropForFileName "displayName" -Value
}
#endregion notificationMessageTemplates

#region roleScopeTags
if ($ProcessscopeTags) {
    Invoke-GraphClass -Class "deviceManagement/roleScopeTags" `
        -Title 'Scope Tags' `
        -PropForFileName "displayName" -Value
}
#endregion roleScopeTags

#region endpointSecurity
if ($ProcessendpointSecurity) {
    Invoke-GraphClassExpand -Class "deviceManagement/intents" `
        -Title 'Endpoint Security Policies' `
        -Properties "displayName","description","isAssigned","lastModifiedDateTime","id" `
        -PropForFileName "displayName" -Value
}
#endregion endpointSecurity

#region groups
if ($ProcessGroups) {
    if ($Export) {
        foreach ($group in $Groups) {
            Export-JSONData -JSON $group -ExportPath "$ExportPath\Groups" -Force | Out-Null
        }
    }
    if ($Document) {
        $groupTable = $Groups | Sort-Object displayName | Select-Object displayName,
            @{Name='groupTypes'; Expression={ $_.groupTypes -join "," }},
            renewedDateTime

        Add-WordText  -WordDocument $WordDocument -Text 'Groups' -HeadingType Heading1 -Supress $True
        Add-WordText  -WordDocument $WordDocument -Text '' -Supress $True
        Add-WordTable -WordDocument $WordDocument -DataTable $groupTable -Design LightGridAccent1 -AutoFit Window -Supress $True
    }
}
#endregion groups

if ($Document) {
    if ($null -ne $WordDocument) {
        try {
            Save-WordDocument -WordDocument $WordDocument -FilePath $FullDocumentationPath -Supress $True
            Write-Host "Document saved: $FullDocumentationPath" -ForegroundColor Green
            Write-Log  "Document saved: $FullDocumentationPath"
            # Abrir el documento automaticamente
            if (Test-Path $FullDocumentationPath) {
                Start-Process $FullDocumentationPath
                Write-Host "Document opened: $FullDocumentationPath" -ForegroundColor Green
            }
        } catch {
            Write-Host "Error saving document: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log  "Error saving document: $($_.Exception.Message)" -LogLevel 3
        }
    } else {
        Write-Host "WordDocument object is null -- cannot save." -ForegroundColor Red
        Write-Log  "WordDocument is null at save time." -LogLevel 3
    }
}

Write-Log "---------- Script Completed ----------"
Write-Host "Script completed." -ForegroundColor Green
#endregion Execution

