function Find-DgmContacts {
<#
.SYNOPSIS
Matches email addresses (from CSV or list) to contacts under an OU.

.DESCRIPTION
You can provide either:
- CsvPath (with -EmailColumn)
- EmailList (array of raw addresses, or a here-string)

.PARAMETER CsvPath
Path to CSV with ';' delimiter and UTF8 encoding.

.PARAMETER EmailColumn
Column in CSV containing the email addresses.

.PARAMETER EmailList
Alternative to CSV: pass an array of addresses, or a here-string.

.PARAMETER SearchBase
OU DN to search.

.PARAMETER PassthruAll
Emit unmatched entries as well.
#>
    [CmdletBinding(DefaultParameterSetName='Csv')]
    param(
        [Parameter(Mandatory, ParameterSetName='Csv')]
        [ValidateScript({ Test-Path $_ })]
        [string]$CsvPath,

        [Parameter(ParameterSetName='Csv')]
        [string]$EmailColumn = 'email',

        [Parameter(Mandatory, ParameterSetName='List')]
        [string[]]$EmailList,

        [Parameter(Mandatory)]
        [string]$SearchBase,

        [switch]$PassthruAll
    )

    # Load emails
    $emails = @()
    if ($PSCmdlet.ParameterSetName -eq 'Csv') {
        $csv = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding UTF8
        if (-not ($csv | Get-Member -Name $EmailColumn -MemberType NoteProperty)) {
            throw "CSV does not contain a column named '$EmailColumn'."
        }
        $emails = $csv | ForEach-Object { $_.$EmailColumn }
    } else {
        $emails = $EmailList
    }

    $idx = Get-DgmContactIndex -SearchBase $SearchBase

    $wanted = @{}
    foreach ($addr in $emails) {
        $k = Normalize-Mail $addr
        if ($k) { $wanted[$k] = $true }
    }

    $seenDN = New-Object System.Collections.Generic.HashSet[string]
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
                New-DgmContactRecord -AdObject $ad -Matches $hits
            }
        }
        elseif ($PassthruAll) {
            [PSCustomObject]@{
                Name              = $null
                DistinguishedName = $null
                Mail              = $null
                TargetAddress     = $null
                ProxyAddresses    = $null
                Matches           = $key
                MatchCount        = 0
                CurrentOU         = $null
                PSTypeName        = 'Dgm.Contact'
            }
        }
    }
}


function Move-DgmContacts {
<#
.SYNOPSIS
Moves contacts (from CSV or list) to a new OU.

.DESCRIPTION
You can provide either:
- CsvPath (with -EmailColumn)
- EmailList (array of addresses)

.PARAMETER CsvPath
CSV path with ';' delimiter.

.PARAMETER EmailColumn
CSV column.

.PARAMETER EmailList
Array/list of addresses to match.

.PARAMETER SearchBase
OU DN where contacts currently live.

.PARAMETER TargetOU
OU DN where to move them.

.PARAMETER ReportPath
Optional path to export a report.
#>
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName='Csv')]
    param(
        [Parameter(Mandatory, ParameterSetName='Csv')]
        [ValidateScript({ Test-Path $_ })]
        [string]$CsvPath,

        [Parameter(ParameterSetName='Csv')]
        [string]$EmailColumn = 'email',

        [Parameter(Mandatory, ParameterSetName='List')]
        [string[]]$EmailList,

        [Parameter(Mandatory)]
        [string]$SearchBase,

        [Parameter(Mandatory)]
        [string]$TargetOU,

        [string]$ReportPath
    )

    $matches = if ($PSCmdlet.ParameterSetName -eq 'Csv') {
        Find-DgmContacts -CsvPath $CsvPath -EmailColumn $EmailColumn -SearchBase $SearchBase
    } else {
        Find-DgmContacts -EmailList $EmailList -SearchBase $SearchBase
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
