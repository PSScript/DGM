# Dgm.ContactTools.psm1
# PowerShell tools to normalize, match, report and relocate Active Directory mail contacts

Set-StrictMode -Version Latest

# --- Helper: normalize emails to lower, trim, remove SMTP: prefix ---
function Normalize-Mail {
    param([string]$Address)
    if (-not $Address) { return $null }
    $a = $Address.Trim()
    if ($a -match '^(?i)smtp:') { $a = $a.Substring(5) }
    return $a.ToLowerInvariant()
}

# --- Helper: shape an AD/Exchange object into a flat report row ---
function New-DgmContactRecord {
    param(
        [Parameter(Mandatory)]$AdObject,
        [string[]]$Matches = @()
    )
    [PSCustomObject]@{
        Name              = $AdObject.Name
        DistinguishedName = $AdObject.DistinguishedName
        Mail              = $AdObject.mail
        TargetAddress     = $AdObject.targetAddress
        ProxyAddresses    = ($AdObject.proxyAddresses -join ' | ')
        Matches           = ($Matches -join ' | ')
        MatchCount        = $Matches.Count
        CurrentOU         = ($AdObject.DistinguishedName -replace '^CN=[^,]+,')
        PSTypeName        = 'Dgm.Contact'
    }
}

# --- Helper: robustly normalize arbitrary lists, arrays, here-strings, multi-line input ---
function Resolve-DgmEmailList {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$EmailList
    )
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($raw in $EmailList) {
        if ($null -eq $raw) { continue }
        $s = [string]$raw
        $parts = $s -split "(`r`n|`n|,|;|`t)"
        foreach ($p in $parts) {
            $t = ($p -as [string]).Trim()
            if ([string]::IsNullOrWhiteSpace($t)) { continue }
            $items.Add($t)
        }
    }
    # Deduplicate (case-insensitive)
    $hs = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($i in $items) { [void]$hs.Add($i) }
    return [string[]]$hs
}

# --- Core: build index of all contacts in OU (AD default, or Exchange with -FromExchange) ---
function Get-DgmContactIndex {
<#
.SYNOPSIS
Builds a lookup hash (mail keys -> contact object) from AD (default) or Exchange.

.PARAMETER SearchBase
OU DN to scan (subtree). Required for AD provider.

.PARAMETER FromExchange
Enumerate contacts via Get-MailContact (org-wide). Ignores -SearchBase.
#>
    [CmdletBinding(DefaultParameterSetName='AD')]
    param(
        [Parameter(ParameterSetName='AD', Mandatory)]
        [string]$SearchBase,

        [Parameter(ParameterSetName='EX')]
        [switch]$FromExchange
    )

    # Collect contacts
    $contacts = @()

    if ($PSCmdlet.ParameterSetName -eq 'EX') {
        # Exchange provider
        $contacts = Get-MailContact -ResultSize Unlimited -ErrorAction Stop |
          Select-Object Name, DistinguishedName,
                        @{n='mail';e={$_.WindowsEmailAddress}},
                        @{n='targetAddress';e={$_.ExternalEmailAddress}},
                        @{n='proxyAddresses';e={$_.EmailAddresses}}
    }
    else {
        # AD provider
        if (-not (Get-Module ActiveDirectory -ListAvailable)) {
            throw "ActiveDirectory module not found on this host."
        }
        Import-Module ActiveDirectory -ErrorAction Stop | Out-Null

        $contacts = Get-ADObject -Filter 'objectClass -eq "contact"' `
                    -SearchBase $SearchBase -SearchScope Subtree `
                    -Properties name, distinguishedName, mail, targetAddress, proxyAddresses
    }

    if (-not $contacts -or $contacts.Count -eq 0) {
        throw "No contacts found by the selected provider."
    }

    # Build index: normalize mail/target/proxies
    $index = @{}
    foreach ($c in $contacts) {
        $keys = New-Object System.Collections.Generic.HashSet[string]
        if ($c.mail)          { $null = $keys.Add( (Normalize-Mail $c.mail) ) }
        if ($c.targetAddress) { $null = $keys.Add( (Normalize-Mail $c.targetAddress) ) }
        if ($c.proxyAddresses){
            foreach ($p in $c.proxyAddresses) { $null = $keys.Add( (Normalize-Mail $p) ) }
        }
        foreach ($k in $keys) {
            if ($k -and -not $index.ContainsKey($k)) { $index[$k] = $c }
        }
    }

    if ($index.Count -eq 0) { throw "Index is empty (no usable mail keys)." }

    [PSCustomObject]@{
        Contacts = $contacts
        Index    = $index
        Provider = if ($PSCmdlet.ParameterSetName -eq 'EX') { 'Exchange' } else { 'AD' }
    }
}

# --- Main: match email list (csv or inline) against the index, returns flat report rows ---
function Find-DgmContacts {
<#
.SYNOPSIS
Matches email addresses (from CSV or inline list) to contacts.

.DESCRIPTION
- Input via CSV (-CsvPath/-EmailColumn) OR via inline list (-EmailList).
- Enumerates contacts from AD (default) or Exchange (-FromExchange).
- Emits one row per matched contact; use -PassthruAll to also see unmatched.

.PARAMETER CsvPath
UTF8 CSV with ';' delimiter.

.PARAMETER EmailColumn
CSV column containing emails (default: 'email').

.PARAMETER EmailList
Inline list: array, here-string, or mixed. Commas/semicolons/newlines OK.

.PARAMETER SearchBase
AD OU DN (required for AD provider). Ignored with -FromExchange.

.PARAMETER FromExchange
Enumerate contacts via Get-MailContact (organization-wide).

.PARAMETER PassthruAll
Also emit synthetic rows for unmatched keys.
#>
    [CmdletBinding(DefaultParameterSetName='Csv')]
    param(
        [Parameter(Mandatory, ParameterSetName='Csv')]
        [ValidateScript({ Test-Path $_ })]
        [string]$CsvPath,

        [Parameter(ParameterSetName='Csv')]
        [string]$EmailColumn = 'email',

        [Parameter(Mandatory, ParameterSetName='List')]
        [object[]]$EmailList,

        [Parameter(ParameterSetName='Csv', Mandatory)]
        [Parameter(ParameterSetName='List', Mandatory)]
        [string]$SearchBase,

        [switch]$FromExchange,

        [switch]$PassthruAll
    )

    # 1) Load/clean emails
    $emails = @()
    if ($PSCmdlet.ParameterSetName -eq 'Csv') {
        $csv = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding UTF8
        if (-not ($csv | Get-Member -Name $EmailColumn -MemberType NoteProperty)) {
            throw "CSV does not contain a column named '$EmailColumn'."
        }
        $emails = $csv | ForEach-Object { $_.$EmailColumn }
    } else {
        $emails = Resolve-DgmEmailList -EmailList $EmailList
    }

    # 2) Build index (AD default; Exchange on request)
    $idx = if ($FromExchange) {
        Get-DgmContactIndex -FromExchange
    } else {
        Get-DgmContactIndex -SearchBase $SearchBase
    }
    if (-not $idx -or -not $idx.Index) { throw "Contact index is null/empty." }

    # 3) Unique wanted keys (normalized)
    $wanted = @{}
    foreach ($addr in $emails) {
        $k = Normalize-Mail $addr
        if ($k) { $wanted[$k] = $true }
    }

    # 4) Match
    $seenDN = New-Object System.Collections.Generic.HashSet[string]
    $matches = New-Object System.Collections.Generic.List[object]

    foreach ($key in $wanted.Keys) {
        if ($idx.Index.ContainsKey($key)) {
            $ad = $idx.Index[$key]
            if ($seenDN.Add($ad.DistinguishedName)) {
                $keys = @()
                $m = Normalize-Mail $ad.mail;           if ($m) { $keys += $m }
                $t = Normalize-Mail $ad.targetAddress;  if ($t) { $keys += $t }
                foreach ($p in $ad.proxyAddresses) {
                    $pp = Normalize-Mail $p; if ($pp) { $keys += $pp }
                }
                $hits = $keys | Where-Object { $_ -and $wanted.ContainsKey($_) }
                $matches.Add( (New-DgmContactRecord -AdObject $ad -Matches $hits) )
            }
        }
        elseif ($PassthruAll) {
            $matches.Add([PSCustomObject]@{
                Name              = $null
                DistinguishedName = $null
                Mail              = $null
                TargetAddress     = $null
                ProxyAddresses    = $null
                Matches           = $key
                MatchCount        = 0
                CurrentOU         = $null
                PSTypeName        = 'Dgm.Contact'
            })
        }
    }

    return $matches
}

# --- Move: for each matched contact, move (with WhatIf/dry-run support), emit flat report ---
function Move-DgmContacts {
<#
.SYNOPSIS
Moves contacts (from CSV or inline list) to a target OU. Supports -WhatIf.

.PARAMETER CsvPath / EmailList
Choose one input method.

.PARAMETER SearchBase
AD OU DN (required for AD provider). Ignored with -FromExchange.

.PARAMETER FromExchange
Enumerate via Exchange (Get-MailContact).

.PARAMETER TargetOU
Destination OU DN.

.PARAMETER ReportPath
Optional CSV report path.
#>
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName='Csv')]
    param(
        [Parameter(Mandatory, ParameterSetName='Csv')]
        [ValidateScript({ Test-Path $_ })]
        [string]$CsvPath,

        [Parameter(ParameterSetName='Csv')]
        [string]$EmailColumn = 'email',

        [Parameter(Mandatory, ParameterSetName='List')]
        [object[]]$EmailList,

        [Parameter(ParameterSetName='Csv', Mandatory)]
        [Parameter(ParameterSetName='List', Mandatory)]
        [string]$SearchBase,

        [switch]$FromExchange,

        [Parameter(Mandatory)]
        [string]$TargetOU,

        [string]$ReportPath
    )

    $matches = if ($PSCmdlet.ParameterSetName -eq 'Csv') {
        Find-DgmContacts -CsvPath $CsvPath -EmailColumn $EmailColumn -SearchBase $SearchBase -FromExchange:$FromExchange
    } else {
        Find-DgmContacts -EmailList $EmailList -SearchBase $SearchBase -FromExchange:$FromExchange
    }

    $results = foreach ($row in $matches) {
        $status = 'Skipped (no DN)'
        $err    = $null

        if ($row.DistinguishedName) {
            if ($PSCmdlet.ShouldProcess($row.Name, "Move to $TargetOU")) {
                try {
                    Move-ADObject -Identity $row.DistinguishedName -TargetPath $TargetOU -Confirm:$false
                    $status = 'Moved'
                } catch {
                    $status = 'Failed'
                    $err    = $_.Exception.Message
                }
            } else {
                $status = 'Dry-Run (would move)'
            }
        }

        [PSCustomObject]@{
            Name              = $row.Name
            Mail              = $row.Mail
            TargetAddress     = $row.TargetAddress
            Matches           = $row.Matches
            MatchCount        = $row.MatchCount
            CurrentOU         = $row.CurrentOU
            TargetOU          = $TargetOU
            Action            = $status
            Error             = $err
        }
    }

    if ($ReportPath) {
        $results | Sort-Object Action, Name |
          Export-Csv -Path $ReportPath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
    }

    return $results
}

# --- (Optional) report: dump all matches to CSV, no move ---
function Export-DgmContactsReport {
<#
.SYNOPSIS
Quickly produce a match-only report without moving anything.

.EXAMPLE
Export-DgmContactsReport -CsvPath ... -SearchBase ... -ReportPath ...
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ })]
        [string]$CsvPath,

        [Parameter()]
        [string]$EmailColumn = 'email',

        [Parameter(Mandatory)]
        [string]$SearchBase,

        [Parameter(Mandatory)]
        [string]$ReportPath
    )

    $rows = Find-DgmContacts -CsvPath $CsvPath -EmailColumn $EmailColumn -SearchBase $SearchBase -PassthruAll
    $rows | Export-Csv -Path $ReportPath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
    return $rows
}

if ($MyInvocation.PSScriptRoot -and $MyInvocation.InvocationName -eq '.') {
    # Dot-sourced, do NOT export
    # (Optional: write-host a friendly note if you want)
}
elseif ($MyInvocation.MyCommand.Module -or $PSCommandPath) {
    Export-ModuleMember -Function `
        Normalize-Mail,New-DgmContactRecord,Resolve-DgmEmailList,Get-DgmContactIndex,Find-DgmContacts,Move-DgmContacts,Export-DgmContactsReport
}
