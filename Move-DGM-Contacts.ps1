#Requires -Modules ActiveDirectory

param(
    [string]$CsvPath      = 'C:\temp\Wave_W1-Filter.csv',
    [string]$EmailColumn  = 'email',
    [string]$SearchBase   = 'OU=DigitalesGemeindeManagement,OU=ELKW-Kontakte,DC=elkw,DC=local',
    [string]$TargetOU     = 'OU=DGM_Konktakte_W1,CN=LostAndFound,DC=elkw,DC=local',
    [switch]$Apply,                 # do the move; omit for dry-run
    [string]$ReportPath  = 'C:\temp\kontakte_move_report.csv'
)

function Normalize-Mail {
    param([string]$addr)
    if (-not $addr) { return $null }
    $a = $addr.Trim()
    if ($a -match '^(?i)smtp:') { $a = $a.Substring(5) }
    return $a.ToLowerInvariant()
}

# 1) Load CSV & build wanted-email set
if (-not (Test-Path $CsvPath)) { throw "CSV not found: $CsvPath" }
$raw = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding UTF8

if (-not ($raw | Get-Member -Name $EmailColumn -MemberType NoteProperty)) {
    throw "CSV does not contain a column named '$EmailColumn'."
}

$wantedEmails = @{}
foreach ($row in $raw) {
    $k = Normalize-Mail $row.$EmailColumn
    if ($k) { $wantedEmails[$k] = $true }
}

Write-Host "Loaded $($wantedEmails.Keys.Count) unique email(s) to find."

# 2) Pull contacts from AD (limit to SearchBase)
$adContacts = Get-ADObject `
    -Filter 'objectClass -eq "contact"' `
    -SearchBase $SearchBase `
    -SearchScope Subtree `
    -Properties name,distinguishedName,mail,targetAddress,proxyAddresses

Write-Host "Scanned $($adContacts.Count) contact object(s) under $SearchBase."

# 3) For each contact, compute all matchable keys
$toProcess = foreach ($c in $adContacts) {
    $keys = New-Object System.Collections.Generic.HashSet[string]
    if ($c.mail)            { $null = $keys.Add( (Normalize-Mail $c.mail) ) }
    if ($c.targetAddress)   { $null = $keys.Add( (Normalize-Mail $c.targetAddress) ) }
    if ($c.proxyAddresses)  {
        foreach ($p in $c.proxyAddresses) {
            $null = $keys.Add( (Normalize-Mail $p) )
        }
    }

    # Which (if any) of those keys are in the wanted set?
    $hits = @()
    foreach ($k in $keys) {
        if ($k -and $wantedEmails.ContainsKey($k)) { $hits += $k }
    }

    [PSCustomObject]@{
        Name               = $c.Name
        DistinguishedName  = $c.DistinguishedName
        Mail               = $c.mail
        TargetAddress      = $c.targetAddress
        ProxyAddresses     = ($c.proxyAddresses -join ' | ')
        Matches            = ($hits -join ' | ')
        MatchCount         = $hits.Count
        ShouldMove         = ($hits.Count -gt 0)
        CurrentOU          = ($c.DistinguishedName -replace '^CN=[^,]+,')
    }
}

# 4) Do (or simulate) the move and build a report
$results = foreach ($row in $toProcess) {
    $status = 'Skipped (no match)'
    $error  = $null

    if ($row.ShouldMove) {
        if ($Apply) {
            try {
                Move-ADObject -Identity $row.DistinguishedName -TargetPath $TargetOU -Confirm:$false
                $status = 'Moved'
            }
            catch {
                $status = 'Failed'
                $error  = $_.Exception.Message
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
        Error             = $error
    }
}

# 5) Export report
$results | Sort-Object Action, Name | Export-Csv -Path $ReportPath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
Write-Host "Report written to $ReportPath"

# 6) Summary
$summary = $results | Group-Object Action | Select-Object Name,Count
$summary | Format-Table -Auto
