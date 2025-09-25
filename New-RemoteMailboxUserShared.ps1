function ConvertTo-MailAlias {
    param([Parameter(Mandatory)][string]$InputString)

    # Replace German umlauts / ß
    $s = $InputString
    $s = $s -replace 'Ä','Ae' -replace 'Ö','Oe' -replace 'Ü','Ue'
    $s = $s -replace 'ä','ae' -replace 'ö','oe' -replace 'ü','ue' -replace 'ß','ss'

    # Lowercase & remove spaces
    $s = $s.ToLower()
    $s = $s -replace '\s+', ''

    # keep only a-z 0-9 . _ -
    $s = $s -replace '[^a-z0-9._-]', ''

    # collapse duplicates and trim edges
    $s = $s -replace '\.{2,}', '.' -replace '_{2,}', '_' -replace '-{2,}', '-'
    $s = $s -replace '^[\.\-_]+', '' -replace '[\.\-_]+$', ''

    if ([string]::IsNullOrWhiteSpace($s)) { $s = 'alias' }
    return $s
}

function Get-UniqueSamAccountName {
    param(
        [Parameter(Mandatory)][string]$Base20,      # already sanitized & <= 20
        [Parameter(Mandatory)][string]$SearchBase   # OU DN
    )
    # Try the base first
    $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$Base20)" -ErrorAction SilentlyContinue
    if (-not $exists) { return $Base20 }

    # Then try adding numeric tails 1..99, trimming the left side to fit 20 total
    for ($n=1; $n -le 99; $n++) {
        $tail = $n.ToString()
        $maxBaseLen = 20 - $tail.Length
        if ($maxBaseLen -lt 1) { continue } # pathological, but safe
        $proposal = $Base20.Substring(0, $maxBaseLen) + $tail
        $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$proposal)" -ErrorAction SilentlyContinue
        if (-not $exists) { return $proposal }
    }

    throw "Could not find a unique sAMAccountName for base '$Base20' within 99 attempts."
}

function New-RemoteMailboxUserShared {
    <#
    .SYNOPSIS
      Create a disabled AD stub for a cloud-first **shared** mailbox (hybrid),
      with strict alias rules (no spaces/umlauts; ≤20; consistent across sAM, mailNickname, routing alias).
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][string]$Alias,                      # pre-sanitized intent; we'll sanitize again just in case
        [Parameter(Mandatory)][string]$OU,                         # target OU DN
        [string]$PrimarySmtpAddress,                               # if empty: <finalAlias>@elkw.de
        [string]$RemoteRoutingDomain = "elkw.mail.onmicrosoft.com",
        [string]$RemoteSuffix = ".OKR",
        [string]$GivenName,
        [string]$Surname,
        [string]$DisplayNameOverride                              # optional, CN/DisplayName if you want a specific one
    )

    # 1) sanitize alias, enforce <= 20, warn on shorten
    $aliasSan = ConvertTo-MailAlias $Alias
    if ($aliasSan.Length -gt 20) {
        $old = $aliasSan
        $aliasSan = $aliasSan.Substring(0,20)
        Write-Warning "Alias > 20 chars. Shortened: '$old' -> '$aliasSan'"
    }

    # 2) find a unique sAMAccountName within the domain (<=20); if collision, add numeric tail
    $sam = Get-UniqueSamAccountName -Base20 $aliasSan -SearchBase $OU

    # 3) set mailNickname to the final alias (same as sAM), per your rule
    $mailNickname = $sam

    # 4) primary SMTP (default)
    if ($PrimarySmtpAddress) {
        $primarySmtp = $PrimarySmtpAddress.ToLower()
    } else {
        $primarySmtp = "$sam@elkw.de"
    }

    # 5) remote routing alias uses the SAME base (sam/final alias), plus suffix
    $routingAlias = "$sam$RemoteSuffix"
    $remoteRoutingAddress = "$routingAlias@$RemoteRoutingDomain"

    # 6) CN/DisplayName (human readable): either override, or GivenName+Surname, else fall back to final alias
    if ($DisplayNameOverride) {
        $displayName = $DisplayNameOverride
    } elseif ($GivenName -or $Surname) {
        $displayName = ($GivenName, $Surname -ne $null) -join ' '  # simple join
        $displayName = ($GivenName ? $GivenName : '') + ($(if($GivenName){' '}else{''})) + ($Surname ? $Surname : '')
    } else {
        $displayName = $sam
    }
    # CN should be safe (no weird LDAP specials)
    $cn = ($displayName -replace '[^\p{L}\p{M}\p{N} \-._]', '_')

    Write-Host "🔹 Creating shared stub '$sam' in '$OU' (mailNickname='$mailNickname', routing='$remoteRoutingAddress') ..." -ForegroundColor Cyan

    # 7) double-check CN uniqueness in OU
    $dupCN = Get-ADUser -SearchBase $OU -LDAPFilter "(cn=$cn)" -ErrorAction SilentlyContinue
    if ($dupCN) {
        # if CN clashes, append a numeric tail similar to sAM logic (but CN has no 20 limit; still keep it tidy)
        for ($n=1; $n -le 99; $n++) {
            $cnTry = "$cn $n"
            $dup = Get-ADUser -SearchBase $OU -LDAPFilter "(cn=$cnTry)" -ErrorAction SilentlyContinue
            if (-not $dup) { $cn = $cnTry; break }
        }
    }

    $userCreated = $false
    $attrsSet = $false

    # 8) Create disabled user (no password)
    try {
        $adUserParams = @{
            Name              = $cn
            SamAccountName    = $sam
            UserPrincipalName = $primarySmtp
            DisplayName       = $displayName
            Enabled           = $false
            Path              = $OU
            Verbose           = $true
        }
        if ($GivenName) { $adUserParams.GivenName = $GivenName }
        if ($Surname)   { $adUserParams.Surname   = $Surname }

        New-ADUser @adUserParams
        $userCreated = $true
        Start-Sleep -Seconds 1
    } catch {
        Write-Host "New-ADUser failed: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }

    # 9) Stamp shared cloud-first attributes
    if ($userCreated) {
        $replace = @{
            msExchRemoteRecipientType  = 97            # cloud-provisioned
            msExchRecipientDisplayType = -2147483642   # synced mailbox
            msExchRecipientTypeDetails = 34359738368   # RemoteSharedMailbox
            mailNickname               = $mailNickname
            targetAddress              = "SMTP:$remoteRoutingAddress"
            mail                       = $primarySmtp
        }
        try {
            Set-ADUser $sam -Replace $replace -Verbose
            # ProxyAddresses: primary uppercase SMTP, remote lowercase smtp
            Set-ADUser $sam -Replace @{ proxyAddresses = @("SMTP:$primarySmtp", "smtp:$remoteRoutingAddress") } -Verbose
            $attrsSet = $true
        } catch {
            Write-Host "Failed to stamp Exchange attributes: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    if ($userCreated -and $attrsSet) {
        Write-Host "Shared stub ready. sAM='$sam'  mailNickname='$mailNickname'  routing='$remoteRoutingAddress'" -ForegroundColor Green
        return Get-ADUser $sam -Properties mail,proxyAddresses,mailNickname,targetAddress
    } elseif ($userCreated) {
        Write-Host "User created but Exchange attrs incomplete." -ForegroundColor Yellow
        return Get-ADUser $sam -Properties *
    } else {
        Write-Host "Nothing created." -ForegroundColor Red
        return $null
    }
}
