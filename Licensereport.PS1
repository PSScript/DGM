<#
.SYNOPSIS
    Microsoft 365 License Report Generator with Active Directory Integration

.DESCRIPTION
    This script generates comprehensive license reports by combining data from:
    - On-premises Active Directory
    - Microsoft Graph API (Azure AD/Entra ID)
    - Microsoft 365 License assignments
    
    The report includes user details, license assignments, mailbox types, and organizational structure.

.PARAMETER OutputPath
    Base directory for report output. Defaults to C:\Data

.PARAMETER TenantConfigPath
    Path to secure configuration file containing tenant credentials

.PARAMETER RoutingDomain
    The routing domain for mailbox addresses (e.g., tenant.mail.onmicrosoft.com)

.PARAMETER SendEmail
    Switch to enable email delivery of reports

.EXAMPLE
    .\LicenseReport-Refactored.ps1 -OutputPath "C:\Reports" -SendEmail

.NOTES
    Author: [Jan Hübener]
    Version: 2.0
    Last Modified: 2025-10-22
    
    SECURITY REQUIREMENTS:
    - Credentials must be stored in secure configuration file (JSON with encrypted secrets)
    - Service account needs: AD read permissions, Graph API permissions (User.Read.All, Directory.Read.All)
    - NEVER commit credentials to source control

.LINK
    https://learn.microsoft.com/en-us/graph/api/resources/users
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "C:\Data",
    
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_})]
    [string]$TenantConfigPath,
    
    [Parameter(Mandatory = $false)]
    [string]$RoutingDomain = "elkw.mail.onmicrosoft.com",
    
    [Parameter(Mandatory = $false)]
    [switch]$SendEmail
)

#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Configuration and Constants

# License type mapping for mailbox types
$script:MailboxTypeMapping = @{
    "1"             = "Onprem_UserMailbox"
    "2"             = "Onprem_LinkedMailbox"
    "4"             = "Onprem_SharedMailbox"
    "16"            = "Onprem_RoomMailbox"
    "32"            = "Onprem_EquipmentMailbox"
    "128"           = "Onprem_Mailuser"
    "2147483648"    = "Cloud_UserMailbox"
    "8589934592"    = "Cloud_RoomMailbox"
    "17179869184"   = "Cloud_EquipmentMailbox"
    "34359738368"   = "Cloud_SharedMailbox"
    "8388608"       = "Onprem_SystemMailbox"
    "549755813888"  = "Onprem_HealthMailbox"
}

# Recipient type detection based on RemoteRecipientType
$script:RecipientTypeRegex = @{
    'UserMailbox'      = '^1$|^2$|^3$|^4$'
    'RoomMailbox'      = '33|35|36|38|49|52'
    'EquipmentMailbox' = '67|68|70|81|84'
    'SharedMailbox'    = '97|99|100|102|116'
}

# Server location mapping
$script:ServerLocationMapping = @{
    "1"             = "Onprem"
    "4"             = "Onprem"
    "2147483648"    = "Cloud"
    "34359738368"   = "Cloud"
}

# AD search base
$script:ADSearchBase = "DC=elkw,DC=local"

# License reference URL
$script:LicenseTableURL = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'

#endregion

#region Helper Functions

function Write-Log {
    <#
    .SYNOPSIS
        Writes timestamped log messages to console and file
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        'Error'   { Write-Host $logMessage -ForegroundColor Red }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Success' { Write-Host $logMessage -ForegroundColor Green }
        default   { Write-Host $logMessage }
    }
    
    # Also write to log file if output folder exists
    if ($script:OutputFolder -and (Test-Path $script:OutputFolder)) {
        Add-Content -Path "$script:OutputFolder\execution.log" -Value $logMessage
    }
}

function Initialize-OutputFolder {
    <#
    .SYNOPSIS
        Creates timestamped output folder structure
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath
    )
    
    try {
        $timestamp = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"
        $folderPath = Join-Path -Path $BasePath -ChildPath "LicenseReport_$timestamp"
        
        $folder = New-Item -Path $folderPath -ItemType Directory -Force
        
        # Create subdirectories
        New-Item -Path "$folderPath\Units" -ItemType Directory -Force | Out-Null
        New-Item -Path "$folderPath\Units\Excel" -ItemType Directory -Force | Out-Null
        
        Write-Log -Message "Output folder created: $folderPath" -Level Success
        
        return $folder.FullName
    }
    catch {
        Write-Log -Message "Failed to create output folder: $_" -Level Error
        throw
    }
}

function Get-SecureConfiguration {
    <#
    .SYNOPSIS
        Loads secure configuration from encrypted file
    .DESCRIPTION
        Configuration file should be JSON format with:
        {
            "TenantId": "guid",
            "ClientId": "guid",
            "ClientSecret": "encrypted_string",
            "EmailFrom": "address",
            "EmailTo": ["addresses"],
            "SmtpServer": "server"
        }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            throw "Configuration file not found: $ConfigPath"
        }
        
        $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        
        # Validate required fields
        $requiredFields = @('TenantId', 'ClientId', 'ClientSecret')
        foreach ($field in $requiredFields) {
            if (-not $config.$field) {
                throw "Missing required configuration field: $field"
            }
        }
        
        Write-Log -Message "Configuration loaded successfully" -Level Success
        return $config
    }
    catch {
        Write-Log -Message "Failed to load configuration: $_" -Level Error
        throw
    }
}

function Get-GraphAPIToken {
    <#
    .SYNOPSIS
        Obtains OAuth token for Microsoft Graph API
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientSecret
    )
    
    try {
        $tokenBody = @{
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            Client_Id     = $ClientId
            Client_Secret = $ClientSecret
        }
        
        $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        
        Write-Log -Message "Requesting Graph API token..."
        $tokenResponse = Invoke-RestMethod -Uri $tokenUri -Method POST -Body $tokenBody
        
        Write-Log -Message "Token obtained successfully" -Level Success
        
        return @{
            "Authorization"     = "Bearer $($tokenResponse.access_token)"
            "Content-type"      = "application/json"
            "ConsistencyLevel"  = "eventual"
        }
    }
    catch {
        Write-Log -Message "Failed to obtain Graph API token: $_" -Level Error
        throw
    }
}

function Get-MicrosoftLicenseTable {
    <#
    .SYNOPSIS
        Downloads or loads Microsoft license reference data
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [switch]$ForceDownload
    )
    
    try {
        $licenseFilePath = Join-Path -Path $OutputPath -ChildPath "Product_names_and_service_plan_identifiers.csv"
        
        # Try to download if forced or file doesn't exist
        if ($ForceDownload -or -not (Test-Path $licenseFilePath)) {
            Write-Log -Message "Downloading Microsoft license reference data..."
            
            try {
                $webContent = Invoke-WebRequest -Uri $script:LicenseTableURL -UseBasicParsing
                $webContent.Content | Out-File -FilePath $licenseFilePath -Force
                Write-Log -Message "License reference downloaded successfully" -Level Success
            }
            catch {
                Write-Log -Message "Download failed (may be blocked by firewall). Using cached version if available." -Level Warning
            }
        }
        
        if (Test-Path $licenseFilePath) {
            $licenseTable = Import-Csv -Path $licenseFilePath -Delimiter "," -Encoding UTF8
            Write-Log -Message "Loaded $($licenseTable.Count) license entries"
            return $licenseTable
        }
        else {
            throw "License reference file not found and download failed"
        }
    }
    catch {
        Write-Log -Message "Failed to load license table: $_" -Level Error
        throw
    }
}

function New-LicenseHashtables {
    <#
    .SYNOPSIS
        Creates lookup hashtables from license table for fast access
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$LicenseTable
    )
    
    try {
        # Filter to valid GUIDs only
        $validLicenses = $LicenseTable | Where-Object {
            [guid]::TryParse($_.GUID, [ref][guid]::Empty) -eq $true
        }
        
        # Remove duplicates by GUID
        $uniqueLicenses = $validLicenses | Sort-Object GUID -Unique
        
        # Create hashtables for lookups
        $guidToDisplayName = @{}
        $guidToStringId = @{}
        $stringIdToDisplayName = @{}
        
        foreach ($license in $uniqueLicenses) {
            $guid = [Guid]$license.GUID
            $guidToDisplayName[$guid] = $license.Product_Display_Name
            $guidToStringId[$guid] = $license.String_Id
            
            if ($license.String_Id) {
                $stringIdToDisplayName[$license.String_Id] = $license.Product_Display_Name
            }
        }
        
        Write-Log -Message "Created license lookup tables: $($guidToDisplayName.Count) entries"
        
        return @{
            GuidToDisplayName      = $guidToDisplayName
            GuidToStringId         = $guidToStringId
            StringIdToDisplayName  = $stringIdToDisplayName
        }
    }
    catch {
        Write-Log -Message "Failed to create license hashtables: $_" -Level Error
        throw
    }
}

function Get-ADUserData {
    <#
    .SYNOPSIS
        Retrieves comprehensive user data from Active Directory
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$SearchBase = $script:ADSearchBase
    )
    
    try {
        Write-Log -Message "Querying Active Directory users..."
        
        $properties = @(
            'DistinguishedName', 'SamAccountName', 'UserPrincipalName', 
            'mail', 'mailNickname', 'proxyAddresses', 'title', 'manager', 
            'employeeID', 'l', 'company', 'postalCode', 'streetAddress', 
            'telephoneNumber', 'countryCode', 'departmentNumber', 'memberOf',
            'TargetAddress', 'msExchRemoteRecipientType', 'msExchRecipientDisplayType',
            'msExchRecipientTypeDetails', 'msExchMailboxGuid', 'extensionAttribute4',
            'Enabled', 'LastLogonDate', 'pwdLastSet', 'accountExpires',
            'msDS-ExternalDirectoryObjectId'
        )
        
        $users = Get-ADUser -Filter * -SearchBase $SearchBase -Properties $properties
        
        Write-Log -Message "Retrieved $($users.Count) AD users" -Level Success
        
        return $users
    }
    catch {
        Write-Log -Message "Failed to query AD: $_" -Level Error
        throw
    }
}

function ConvertTo-EnrichedADUser {
    <#
    .SYNOPSIS
        Enriches AD user object with calculated properties
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$UserHashtableByDN
    )
    
    process {
        try {
            # Extract organizational units
            $ouParts = ($User.DistinguishedName -split "," | Where-Object { $_ -match "OU=" }) -replace "OU="
            $ouPath = ($ouParts[($ouParts.Count)..0]) -join "/"
            
            # Determine top OU based on structure
            $topOU = if ($User.DistinguishedName -match "365") {
                ($ouParts[1..0]) -join "/"
            }
            else {
                $ouParts[0]
            }
            
            # Get manager information if exists
            $managerUPN = $null
            $managerMail = $null
            $managerOU = $null
            
            if ($User.manager) {
                $managerObj = $UserHashtableByDN[[string]$User.manager]
                if ($managerObj) {
                    $managerUPN = $managerObj.UserPrincipalName
                    $managerMail = $managerObj.mail
                    
                    $managerOUParts = ($User.manager -split "," | Where-Object { $_ -match "OU=" }) -replace "OU="
                    $managerOU = ($managerOUParts[($managerOUParts.Count)..0]) -join "/"
                }
            }
            
            # Extract proxy addresses
            $proxyAddresses = $User.proxyAddresses
            $primarySMTP = ($proxyAddresses | Where-Object { $_ -cmatch "^SMTP:" }) -replace "SMTP:"
            $allSmtpAddresses = ($proxyAddresses | Where-Object { $_ -cmatch "^smtp:" }) -replace "smtp:" | Where-Object { $_ -notmatch $RoutingDomain -and $_ -notmatch "datagroup.local" }
            $x500Addresses = ($proxyAddresses | Where-Object { $_ -match "^X500:" }) -join '|'
            
            # Determine mailbox type and server location
            $recipientTypeDetails = $User.msExchRecipientTypeDetails
            $mailboxType = $script:MailboxTypeMapping[[string]$recipientTypeDetails]
            $serverLocation = $script:ServerLocationMapping[[string]$recipientTypeDetails]
            
            # Determine recipient type
            $recipientType = $null
            $remoteRecipientType = [string]$User.msExchRemoteRecipientType
            foreach ($type in $script:RecipientTypeRegex.Keys) {
                if ($remoteRecipientType -match $script:RecipientTypeRegex[$type]) {
                    $recipientType = $type
                    break
                }
            }
            
            # Extract license groups (M365 groups)
            $licenseGroups = (($User.memberOf | Where-Object { $_ -match "M365" }) -split "," | 
                Where-Object { $_ -match "CN=" }) -replace "CN=" | Join-Object -Separator "|"
            
            # Create enriched object
            return [PSCustomObject]@{
                SamAccountName                = $User.SamAccountName
                UserPrincipalName             = $User.UserPrincipalName
                Mail                          = $User.mail
                MailNickname                  = $User.mailNickname
                msExchMailboxGuid             = if ($User.msExchMailboxGuid) { [guid]$User.msExchMailboxGuid } else { $null }
                Title                         = $User.title
                EmployeeID                    = $User.employeeID
                City                          = $User.l
                Company                       = $User.company
                PostalCode                    = $User.postalCode
                StreetAddress                 = $User.streetAddress
                TelephoneNumber               = ([string]$User.telephoneNumber).Trim() -replace "00%$"
                CountryCode                   = $User.countryCode
                Department                    = [string]$User.departmentNumber
                LicenseGroups                 = $licenseGroups
                OU                            = $ouPath
                TopOU                         = $topOU
                Manager                       = $User.manager
                ManagerUPN                    = $managerUPN
                ManagerMail                   = $managerMail
                ManagerOU                     = $managerOU
                ManagerOUMatch                = if ($managerOU -eq $ouPath) { "YES" } else { "NO" }
                PrimarySMTPAddress            = $primarySMTP
                ProxyAddresses                = ($allSmtpAddresses -join '|')
                X500Addresses                 = $x500Addresses
                TargetAddress                 = ([string]$User.TargetAddress) -replace "smtp:"
                TargetAddressSam              = "$($User.SamAccountName)@$RoutingDomain"
                TargetAddressNickname         = "$($User.mailNickname)@$RoutingDomain"
                MailboxType                   = $mailboxType
                RecipientType                 = $recipientType
                ServerLocation                = $serverLocation
                msExchRemoteRecipientType     = $User.msExchRemoteRecipientType
                msExchRecipientDisplayType    = $User.msExchRecipientDisplayType
                msExchRecipientTypeDetails    = $User.msExchRecipientTypeDetails
                ExtensionAttribute4           = $User.extensionAttribute4
                Enabled                       = $User.Enabled
                LastLogonDate                 = $User.LastLogonDate
                PwdLastSet                    = [datetime]::FromFileTime($User.pwdLastSet)
                AccountExpires                = [datetime]::FromFileTime($User.accountExpires)
                ExternalDirectoryObjectId     = ([string]$User."msDS-ExternalDirectoryObjectId") -replace "User_"
                DistinguishedName             = $User.DistinguishedName
            }
        }
        catch {
            Write-Log -Message "Failed to enrich user $($User.SamAccountName): $_" -Level Warning
            return $null
        }
    }
}

function Get-GraphAPIData {
    <#
    .SYNOPSIS
        Retrieves data from Microsoft Graph API with pagination support
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Headers,
        
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )
    
    try {
        $allResults = @()
        $nextLink = $Uri
        
        do {
            Write-Log -Message "Fetching: $nextLink"
            $response = Invoke-RestMethod -Uri $nextLink -Headers $Headers -Method Get
            
            if ($response.value) {
                $allResults += $response.value
            }
            
            $nextLink = $response.'@odata.nextLink'
            
        } while ($nextLink)
        
        Write-Log -Message "Retrieved $($allResults.Count) items" -Level Success
        return $allResults
    }
    catch {
        Write-Log -Message "Graph API request failed: $_" -Level Error
        throw
    }
}

#endregion

#region Main Execution

function Invoke-LicenseReport {
    <#
    .SYNOPSIS
        Main function orchestrating the license report generation
    #>
    [CmdletBinding()]
    param()
    
    try {
        Write-Log -Message "=== License Report Generation Started ===" -Level Info
        Write-Log -Message "PowerShell Version: $($PSVersionTable.PSVersion)"
        
        # Initialize output folder
        $script:OutputFolder = Initialize-OutputFolder -BasePath $OutputPath
        Invoke-Item $script:OutputFolder
        
        # Load secure configuration
        Write-Log -Message "Loading secure configuration..."
        $config = Get-SecureConfiguration -ConfigPath $TenantConfigPath
        
        # Get Graph API token
        $headers = Get-GraphAPIToken -TenantId $config.TenantId `
            -ClientId $config.ClientId `
            -ClientSecret $config.ClientSecret
        
        # Load Microsoft license reference data
        Write-Log -Message "Loading license reference data..."
        $licenseTable = Get-MicrosoftLicenseTable -OutputPath $OutputPath
        $licenseHashtables = New-LicenseHashtables -LicenseTable $licenseTable
        
        # Export license table to output folder
        $licenseTable | Export-Csv -Path "$script:OutputFolder\License_Reference.csv" `
            -Delimiter "," -Encoding UTF8 -NoTypeInformation -Force
        
        # Get AD user data
        Write-Log -Message "Retrieving Active Directory data..."
        $adUsers = Get-ADUserData
        
        # Create hashtable for fast DN lookups
        $userHashtableByDN = @{}
        foreach ($user in $adUsers) {
            $userHashtableByDN[[string]$user.DistinguishedName] = $user
        }
        
        # Enrich AD user data
        Write-Log -Message "Processing AD user data..."
        $enrichedADUsers = $adUsers | ForEach-Object {
            ConvertTo-EnrichedADUser -User $_ -UserHashtableByDN $userHashtableByDN
        } | Where-Object { $_ -ne $null }
        
        # Export AD data
        Write-Log -Message "Exporting AD user data..."
        $enrichedADUsers | Export-Csv -Path "$script:OutputFolder\AD_Users_Complete.csv" `
            -Delimiter ";" -Encoding UTF8 -NoTypeInformation -Force
        
        # Create UPN hashtable for quick lookups
        $userHashtableByUPN = @{}
        foreach ($user in $enrichedADUsers) {
            $userHashtableByUPN[$user.UserPrincipalName] = $user
        }
        
        # Get Graph API data
        Write-Log -Message "Retrieving Microsoft 365 user data..."
        $graphUsers = Get-GraphAPIData -Headers $headers `
            -Uri "https://graph.microsoft.com/v1.0/users?`$select=id,userPrincipalName,displayName,mail,userType,accountEnabled,createdDateTime,signInActivity,assignedLicenses,licenseAssignmentStates,onPremisesSyncEnabled"
        
        Write-Log -Message "Retrieving Microsoft 365 license data..."
        $subscribedSkus = Get-GraphAPIData -Headers $headers `
            -Uri "https://graph.microsoft.com/v1.0/subscribedSkus"
        
        # TODO: Combine AD and Graph data, generate reports
        # This is where you would merge the data and create your comprehensive reports
        
        Write-Log -Message "=== License Report Generation Completed ===" -Level Success
        Write-Log -Message "Reports saved to: $script:OutputFolder" -Level Success
        
        return $script:OutputFolder
    }
    catch {
        Write-Log -Message "CRITICAL ERROR: $_" -Level Error
        Write-Log -Message "Stack Trace: $($_.ScriptStackTrace)" -Level Error
        throw
    }
}

#endregion

# Execute main function
try {
    Invoke-LicenseReport
}
catch {
    Write-Host "`nScript execution failed. Check the log file for details." -ForegroundColor Red
    #exit 1
}