<#
.SYNOPSIS
    PowerShell AD attribute management toolkit for extensionAttribute9
.DESCRIPTION
    Functions to resolve users by UPN/email, preview and report extensionAttribute9,
    and safely set or clear values, including WhatIf/Confirm support.
.LICENSE
    MIT License (see https://opensource.org/licenses/MIT)
    Copyright (c) 2025 Jan Hübener
#>

# Requires -Modules ActiveDirectory

function Resolve-AdPerson {
<#
.SYNOPSIS
    Resolve a UPN/email to a unique AD user object.
.DESCRIPTION
    Accepts pipeline or parameter input; returns the first match found.
.PARAMETER Identifier
    UPN, email, or samAccountName.
.EXAMPLE
    'user@domain.de' | Resolve-AdPerson
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Id','UPN','Mail')]
        [string]$Identifier
    )
    process {
        $u = Get-ADUser -LDAPFilter "(|(userPrincipalName=$Identifier)(mail=$Identifier)(samAccountName=$Identifier))" `
            -Properties mail,userPrincipalName,displayName,extensionAttribute9,samAccountName
        if ($u -is [array]) {
            Write-Warning "Multiple matches for '$Identifier' → using first: $($u[0].SamAccountName)"
            $u = $u[0]
        }
        if (-not $u) { Write-Warning "Not found: $Identifier" }
        $u
    }
}

function Get-ExtAttr9Report {
<#
.SYNOPSIS
    Report current extensionAttribute9 state for users.
.DESCRIPTION
    Accepts pipeline of identifiers (UPN/email/sam) and emits an object for each.
.EXAMPLE
    $list | Get-ExtAttr9Report | Format-Table
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Id','UPN','Mail')]
        [string]$Identifier
    )
    begin { $bag = New-Object System.Collections.Generic.List[object] }
    process {
        $u = $Identifier | Resolve-AdPerson
        if ($u) {
            $bag.Add([pscustomobject]@{
                InputId            = $Identifier
                Found              = $true
                SamAccountName     = $u.SamAccountName
                UserPrincipalName  = $u.UserPrincipalName
                Mail               = $u.Mail
                DisplayName        = $u.DisplayName
                ExtensionAttribute9= $u.extensionAttribute9
            })
        } else {
            $bag.Add([pscustomobject]@{
                InputId            = $Identifier
                Found              = $false
                SamAccountName     = $null
                UserPrincipalName  = $null
                Mail               = $null
                DisplayName        = $null
                ExtensionAttribute9= $null
            })
        }
    }
    end { $bag.ToArray() }
}

function Set-ExtAttr9 {
<#
.SYNOPSIS
    Set extensionAttribute9 for AD users.
.DESCRIPTION
    Accepts identifiers via pipeline. Honors -WhatIf and -Confirm.
.PARAMETER Identifier
    UPN, email, or samAccountName.
.PARAMETER TargetValue
    Value to set (default: RouteToKopano)
.PARAMETER OnlyWhenEmpty
    Only set if attribute is null/empty.
.EXAMPLE
    $list | Set-ExtAttr9 -TargetValue 'RouteToKopano' -WhatIf
    Import-Csv users.csv | Select-Object -ExpandProperty Mail | Set-ExtAttr9 -TargetValue 'RouteToKopano' -Confirm
#>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Id','UPN','Mail')]
        [string]$Identifier,

        [string]$TargetValue = 'RouteToKopano',

        [switch]$OnlyWhenEmpty
    )
    process {
        $u = $Identifier | Resolve-AdPerson
        if (-not $u) { return }
        $current = $u.extensionAttribute9
        if ($OnlyWhenEmpty -and -not [string]::IsNullOrWhiteSpace($current)) {
            Write-Verbose "Skip $($u.SamAccountName): value already set ('$current')."
            return
        }
        if (-not $OnlyWhenEmpty -and $current -eq $TargetValue) {
            Write-Verbose "Skip $($u.SamAccountName): already '$TargetValue'."
            return
        }
        $msg = "$($u.SamAccountName) ($($u.DisplayName)): '$current' -> '$TargetValue'"
        if ($PSCmdlet.ShouldProcess($u.SamAccountName, "Set extensionAttribute9 to '$TargetValue'")) {
            try {
                Set-ADUser -Identity $u.SamAccountName -Replace @{extensionAttribute9 = $TargetValue}
                Write-Host "Set $msg"
            } catch {
                Write-Warning "FAILED $msg : $($_.Exception.Message)"
            }
        }
    }
}

function Clear-ExtAttr9 {
<#
.SYNOPSIS
    Clear extensionAttribute9 for AD users.
.DESCRIPTION
    Accepts identifiers via pipeline. Honors -WhatIf and -Confirm.
.EXAMPLE
    $list | Clear-ExtAttr9 -WhatIf
#>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Id','UPN','Mail')]
        [string]$Identifier
    )
    process {
        $u = $Identifier | Resolve-AdPerson
        if (-not $u) { return }
        $current = $u.extensionAttribute9
        if ([string]::IsNullOrWhiteSpace($current)) {
            Write-Verbose "Skip $($u.SamAccountName): already empty."
            return
        }
        if ($PSCmdlet.ShouldProcess($u.SamAccountName, "Clear extensionAttribute9 (was '$current')")) {
            try {
                Set-ADUser -Identity $u.SamAccountName -Clear extensionAttribute9
                Write-Host "Cleared $($u.SamAccountName) ($($u.DisplayName))"
            } catch {
                Write-Warning "FAILED to clear $($u.SamAccountName): $($_.Exception.Message)"
            }
        }
    }
}
