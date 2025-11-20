# ===============================
# DGM Master Builder v4.4 - COMPLETE & CORRECTED
# Wave files filtered by MailboxMigration=YES, routing only for non-elkw.local
# ===============================
$ErrorActionPreference = 'Stop'
Set-StrictMode -Off

# --- Initialize global lookup tables early (PS5.1-safe) ---
$gemeinden_hash = @{}
$Welle_vorher_hash = @{}
$Welle_nachher_hash = @{}
$Welle_Dekanat_hash = @{}
$Welle_Kirchenbezirk_hash = @{}
$licenseByUPN = @{}
$ct_hash = @{}
$ctAdminsByInstanz = @{}
$ctStatsByInstanz = @{}
$passwordByEmail = @{}
$migrationStatusByEmail = @{}

# Pilot GKZ List & Hash
$PilotGKZ_List = @("420048", "323032", "425019", "444105", "44", "444108", "224026", "407008")
$PilotHash = @{}

# 1. --- PATH CONFIGURATION ---
$SessionRoot         = "C:\temp3"
$GemeindenCTPath     = Join-Path $SessionRoot "Gemeinden_Alle_CT.csv"
$WellenPath          = Join-Path $SessionRoot "wellen.csv"
$CancomPath          = Join-Path $SessionRoot "2025-11-17_benutzer_status_Cancom.csv"
$SharedMailboxPath   = Join-Path $SessionRoot "2025-11-17_funktionspostfaecher_status_Cancom_semicolon.csv"
$LizenzReportPath    = Join-Path $SessionRoot "Lizenzreport_ALL1.csv"
$ChurchtoolsPath     = Join-Path $SessionRoot "churchtools_organizations.csv"
$ElkwAdminsPath      = Join-Path $SessionRoot "elkwadmins-2025-09-15-krz.csv"
$ElkwStatsPath       = Join-Path $SessionRoot "elkwstats-2025-09-15-krz.csv"
$PasswordMasterPath  = "C:\temp\master_users_all_merged_MASTER_NEU2.csv"
$Quick2InputPath     = Join-Path $SessionRoot "MigrationUsers_Quick2.csv"

# Shared Mailbox Default Password
$SharedMailboxPassword = "tklT75Rc8bxEw4jhLf6A#"

# Current batch threshold (adjust as migration progresses)
$CurrentBatchThreshold = 10  # W10 = next batch to migrate

# Output dirs
$RunStamp            = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$RunRoot             = Join-Path $SessionRoot ("out_" + $RunStamp)
$OutDir_Master       = Join-Path $RunRoot "Master"
$OutDir_Waves        = Join-Path $RunRoot "Waves"
New-Item -ItemType Directory -Path $RunRoot,$OutDir_Master,$OutDir_Waves -Force | Out-Null

$Delimiter = ';'

# 2. --- DATE PARSING HELPERS ---
function Parse-DateFlexible {
    param([string]$s)
    if (-not $s) { return $null }
    $t = $s.Trim()
    if ($t -eq '' -or $t -eq 'NULL') { return $null }
    $de = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    $styles = [System.Globalization.DateTimeStyles]::None
    $fmts = @('dd.MM.yyyy','d.M.yyyy','dd.MM.yy','d.M.yy','yyyy-MM-dd','yyyy/MM/dd','yyyy.MM.dd','yyyy')
    foreach ($f in $fmts) { try { return [datetime]::ParseExact($t, $f, $de, $styles) } catch { } }
    try   { return [datetime]::Parse($t, $de, $styles) } catch { }
    try   { return [datetime]::Parse($t, [System.Globalization.CultureInfo]::InvariantCulture, $styles) } catch { }
    return $null
}

function To-IsoDateString { 
    param([string]$s)
    $dt = Parse-DateFlexible $s
    if ($dt) { return $dt.ToString('yyyy-MM-dd') }
    return '' 
}

function YearOrNull { 
    param($s)
    $dt = Parse-DateFlexible ($s -as [string])
    if ($dt) { return $dt.Year }
    return $null 
}

function Get-PropValue { 
    param($obj, [string]$prop)
    if (-not $obj) { return "" }
    if ($obj.PSObject.Properties.Match($prop).Count -gt 0) {
        $val = $obj.$prop
        if ($null -ne $val) { return "$val".Trim() }
    }
    return "" 
}

# 3. --- GENERAL HELPERS ---
function Strip-Quotes($path) { 
    if (Test-Path $path) { 
        (Get-Content $path) -replace '"','' | Set-Content $path -Encoding UTF8 
    } 
}

function Normalize-Key($s) { 
    $s = ($s -as [string])
    if (-not $s) { return "" }
    return ($s.Trim().ToLower()) -replace '\s+', ' '
}

function SplitEmails($arr) { 
    return ($arr -join ',') -split '[,\s]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ } 
}

function CanonKRO { 
    param($val, $pad = 3)
    $s = ($val -as [string]).Trim()
    if ([string]::IsNullOrEmpty($s) -or $s -eq 'NULL') { return $null }
    return $s.PadLeft($pad,'0')
}

function GetCTInstanzFromUrl($url) { 
    $u = ($url -as [string]).Trim()
    if ([string]::IsNullOrEmpty($u)) { return "" }
    if ($u -match 'https://(elkw\d{4})\.krz\.tools') { return $matches[1] }
    elseif ($u -match 'elkw(\d{4})') { return "elkw$($matches[1])" }
    return ""
}

function GetCTDesignations($g, $ct_hash) {
    $kros = @()
    if ($g.KRO) { $kros += ($g.KRO -as [string]).Trim() }
    if ($g.PSObject.Properties['kro_alt'] -and $g.kro_alt) {
        $kros += ($g.kro_alt -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    $list = @()
    foreach ($k in $kros) {
        if ($ct_hash.ContainsKey($k)) {
            $d = $ct_hash[$k].designation
            $d = if ($d) { $d.Trim() } else { "" }
            if ($d -and $list -notcontains $d) { $list += $d }
        }
    }
    return $list -join ', '
}

function GetCTInstanzListByKROs($kros, $ct_hash) {
    $insts = @()
    foreach ($k in $kros) {
        $id = ($k -as [string]).Trim()
        if ($id -and $ct_hash.ContainsKey($id)) {
            $val = $ct_hash[$id].CT_Instanz
            if ($val) { $insts += $val }
        }
    }
    return ($insts | Sort-Object -Unique) -join ', '
}

# ## --- 3-WAY WAVE MAPPING FUNCTION (USERS & SHARED) --- ##
function Get-WaveInfo3Way {
    param(
        [string]$KRO_GKZ,           # Main GKZ
        [string]$Vertrauensinstanz, # Vertrauensinstanz GKZ (users only)
        [string]$DekanatName        # Dekanat name for lookup
    )
    
    # 1. Wave via KRO → Dekanat
    $Wave_via_KRO_Dekanat = ""
    $keyDek = Normalize-Key $DekanatName
    if ($keyDek -and $script:Welle_Dekanat_hash.ContainsKey($keyDek)) { 
        $Wave_via_KRO_Dekanat = $script:Welle_Dekanat_hash[$keyDek].Wave 
    }

    # 2. Wave via Vertrauensinstanz → Dekanat (if provided)
    $Wave_via_Vertrauensinstanz = ""
    if ($Vertrauensinstanz) {
        $vCanon = CanonKRO $Vertrauensinstanz 3
        if ($vCanon -and $script:gemeinden_hash.ContainsKey($vCanon)) {
            $gv = $script:gemeinden_hash[$vCanon]
            $dekv = ($gv.ORENAMEEBENE4 -as [string]).Trim()
            $keyv = Normalize-Key $dekv
            if ($keyv -and $script:Welle_Dekanat_hash.ContainsKey($keyv)) { 
                $Wave_via_Vertrauensinstanz = $script:Welle_Dekanat_hash[$keyv].Wave 
            }
        }
    }

    # 3. Wave via Fallback (KRO direct lookup)
    $Wave_via_Fallback = ""
    $fallback_key = Normalize-Key $KRO_GKZ
    if ($fallback_key -and $script:Welle_vorher_hash.ContainsKey($fallback_key)) { 
        $Wave_via_Fallback = $script:Welle_vorher_hash[$fallback_key].Wave 
    }

    # Decision logic
    $Wave = ""
    $Wave_Source = ""
    $Wave_Warning = ""

    if ($Wave_via_KRO_Dekanat -and $Wave_via_Vertrauensinstanz -and $Wave_via_Fallback) {
        if (($Wave_via_KRO_Dekanat -eq $Wave_via_Vertrauensinstanz) -and ($Wave_via_KRO_Dekanat -eq $Wave_via_Fallback)) {
            $Wave = $Wave_via_KRO_Dekanat
            $Wave_Source = "All3(KRO->Dekanat & Vertrauensinstanz & Fallback)"
        } else {
            $Wave = $Wave_via_KRO_Dekanat
            $Wave_Source = "KRO->Dekanat"
            $Wave_Warning = "3-way mismatch: KRO=$Wave_via_KRO_Dekanat, Vertr=$Wave_via_Vertrauensinstanz, Fallback=$Wave_via_Fallback"
        }
    } elseif ($Wave_via_KRO_Dekanat -and $Wave_via_Vertrauensinstanz) {
        if ($Wave_via_KRO_Dekanat -eq $Wave_via_Vertrauensinstanz) {
            $Wave = $Wave_via_KRO_Dekanat
            $Wave_Source = "Both(KRO->Dekanat & Vertrauensinstanz)"
        } else {
            $Wave = $Wave_via_KRO_Dekanat
            $Wave_Source = "KRO->Dekanat"
            $Wave_Warning = "Mismatch: KRO=$Wave_via_KRO_Dekanat vs Vertr=$Wave_via_Vertrauensinstanz"
        }
    } elseif ($Wave_via_KRO_Dekanat -and $Wave_via_Fallback) {
        if ($Wave_via_KRO_Dekanat -eq $Wave_via_Fallback) {
            $Wave = $Wave_via_KRO_Dekanat
            $Wave_Source = "Both(KRO->Dekanat & Fallback)"
        } else {
            $Wave = $Wave_via_KRO_Dekanat
            $Wave_Source = "KRO->Dekanat"
            $Wave_Warning = "Mismatch: KRO=$Wave_via_KRO_Dekanat vs Fallback=$Wave_via_Fallback"
        }
    } elseif ($Wave_via_Vertrauensinstanz -and $Wave_via_Fallback) {
        if ($Wave_via_Vertrauensinstanz -eq $Wave_via_Fallback) {
            $Wave = $Wave_via_Vertrauensinstanz
            $Wave_Source = "Both(Vertrauensinstanz & Fallback)"
        } else {
            $Wave = $Wave_via_Vertrauensinstanz
            $Wave_Source = "Vertrauensinstanz"
            $Wave_Warning = "Mismatch: Vertr=$Wave_via_Vertrauensinstanz vs Fallback=$Wave_via_Fallback"
        }
    } elseif ($Wave_via_KRO_Dekanat) {
        $Wave = $Wave_via_KRO_Dekanat
        $Wave_Source = "KRO->Dekanat"
    } elseif ($Wave_via_Vertrauensinstanz) {
        $Wave = $Wave_via_Vertrauensinstanz
        $Wave_Source = "Vertrauensinstanz"
    } elseif ($Wave_via_Fallback) {
        $Wave = $Wave_via_Fallback
        $Wave_Source = "Fallback"
    }

    return [PSCustomObject]@{
        Wave = $Wave
        Source = $Wave_Source
        Warning = $Wave_Warning
    }
}

# 4. --- LOAD INPUT FILES ---
Write-Host "Loading input files..."
$gemeindenCT = Import-Csv $GemeindenCTPath -Delimiter $Delimiter -Encoding UTF8
$wellen      = Import-Csv $WellenPath -Delimiter $Delimiter -Encoding UTF8
$cancom      = Import-Csv $CancomPath -Delimiter ',' -Encoding UTF8
$lizenz      = Import-Csv $LizenzReportPath -Delimiter $Delimiter -Encoding UTF8
$churchtools = Import-Csv $ChurchtoolsPath -Delimiter $Delimiter -Encoding UTF8
$elkwAdmins  = Import-Csv $ElkwAdminsPath -Delimiter ';' -Encoding UTF8
$elkwStats   = Import-Csv $ElkwStatsPath  -Delimiter ';' -Encoding UTF8

# Load shared mailboxes if file exists (SEMICOLON SEPARATOR)
$sharedMailboxes = @()
if (Test-Path $SharedMailboxPath) {
    Write-Host "Loading shared mailboxes from $SharedMailboxPath"
    $sharedMailboxes = Import-Csv $SharedMailboxPath -Delimiter ';' -Encoding UTF8
}

# Load password master file
if (Test-Path $PasswordMasterPath) {
    Write-Host "Loading password master file from $PasswordMasterPath"
    $passwordData = Import-Csv $PasswordMasterPath -Delimiter ';' -Encoding UTF8
    foreach ($row in $passwordData) {
        $email = ($row.email -as [string]).ToLower().Trim()
        $password = ($row.password -as [string]).Trim()
        if ($email -and $password) {
            $passwordByEmail[$email] = $password
        }
    }
    Write-Host "  Loaded $($passwordByEmail.Count) passwords from master file"
} else {
    Write-Warning "Password master file not found: $PasswordMasterPath"
}

# Load Quick2 INPUT file (migration status from Get-MigrationUser)
if (Test-Path $Quick2InputPath) {
    Write-Host "Loading Quick2 migration status from $Quick2InputPath"
    $header = Get-Content $Quick2InputPath -TotalCount 1
    $delim = if ($header -match ';') { ';' } else { ',' }
    
    $quick2Data = Import-Csv $Quick2InputPath -Delimiter $delim -Encoding UTF8
    foreach ($row in $quick2Data) {
        $ea = Get-PropValue $row 'EmailAddress'
        if ($ea) {
            $email = $ea.ToLower().Trim()
            $status = Get-PropValue $row 'UserStatus'
            $migrationStatusByEmail[$email] = if ($status) { $status } else { "Unknown" }
        }
    }
    Write-Host "  Loaded migration status for $($migrationStatusByEmail.Count) accounts"
} else {
    Write-Warning "Quick2 input file not found: $Quick2InputPath (all users treated as unmigrated)"
}

# Build Pilot hash
foreach ($p in $PilotGKZ_List) { 
    $canon = CanonKRO $p 3
    if ($canon) { $PilotHash[$canon] = $true } 
}
Write-Host "  Pilot GKZ configured: $($PilotHash.Count) entries"

# 5. --- BUILD LOOKUP HASHES ---
Write-Host "Building lookup hashes..."

# Gemeinden hash
foreach ($g in $gemeindenCT) { 
    $k = ($g.KRO -as [string]).Trim()
    if ($k) { $gemeinden_hash[$k] = $g }
}

# Wellen hashes
foreach ($w in $wellen) {
    $v1 = Normalize-Key $w.Welle_vorher
    $v2 = Normalize-Key $w.Welle_nachher
    $v3 = Normalize-Key $w.Welle_Dekanat
    $v4 = Normalize-Key $w.Welle_Kirchenbezirk
    if ($v1) { $Welle_vorher_hash[$v1] = $w }
    if ($v2) { $Welle_nachher_hash[$v2] = $w }
    if ($v3) { $Welle_Dekanat_hash[$v3] = $w }
    if ($v4) { $Welle_Kirchenbezirk_hash[$v4] = $w }
}

# License hash - key by UserPrincipalName
foreach ($l in $lizenz) {
    $upn = Get-PropValue $l 'UserPrincipalName'
    if ($upn) { 
        $licenseByUPN[$upn.ToLower().Trim()] = $l 
    }
}

# ChurchTools hash
foreach ($ct in $churchtools) {
    $id  = ($ct.identifier -as [string]).Trim()
    $url = ($ct.url -as [string]).Trim()
    if ($id) {
        $ct_hash[$id] = @{
            CT_Instanz    = GetCTInstanzFromUrl $url
            designation   = ($ct.designation -as [string]).Trim()
            url           = $url
        }
    }
}

# CT Admins by Instanz
foreach ($row in $elkwAdmins) {
    $inst = ($row.Instanz -as [string]).Trim()
    if (!$inst) { continue }
    if (-not $ctAdminsByInstanz.ContainsKey($inst)) { $ctAdminsByInstanz[$inst] = @() }
    $mail = ($row.'E-Mail' -as [string]).Trim()
    if ($mail -and $ctAdminsByInstanz[$inst] -notcontains $mail) {
        $ctAdminsByInstanz[$inst] += $mail
    }
}

# CT Stats by Instanz
foreach ($row in $elkwStats) {
    $inst = ($row.Instanz -as [string]).Trim()
    if ($inst) { $ctStatsByInstanz[$inst] = $row }
}

# 6. --- ENRICH GEMEINDEN DATA ---
Write-Host "Enriching Gemeinden data..."

# Enrich All_Emails to always include EMAIL1 from 2025
foreach ($row in $gemeindenCT) {
    $email1_2025 = ""
    if ($row.PSObject.Properties.Match('Year_Source').Count -gt 0 -and $row.Year_Source -eq 2025) {
        $email1_2025 = ($row.EMAIL1 -as [string]).Trim()
    }
    $pip = $row.EMAIL_PIP_Merged
    $pf  = $row.PF_EMail1_Merged
    $all = SplitEmails @( $email1_2025, $pip, $pf ) | Sort-Object -Unique
    $row | Add-Member -NotePropertyName All_Emails -NotePropertyValue ($all -join ', ') -Force
    $row | Add-Member -NotePropertyName Emails_For_Migration -NotePropertyValue ($all -join ', ') -Force
}

# CT_Instanz_Gemeinde aggregate
foreach ($row in $gemeindenCT) {
    $kros = @()
    $main = ($row.KRO -as [string]).Trim()
    if ($main) { $kros += $main }
    if ($row.PSObject.Properties['kro_alt'] -and $row.kro_alt) {
        $kros += ($row.kro_alt -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    $cti = GetCTInstanzListByKROs $kros $ct_hash
    $row | Add-Member -NotePropertyName CT_Instanz_Gemeinde -NotePropertyValue $cti -Force
}

# 7. --- BUILD MASTER USERS LIST ---
Write-Host "Building master users list..."
$masterUsers = New-Object System.Collections.Generic.List[psobject]

foreach ($u in $cancom) {
    try {
        $email    = if ($u.PSObject.Properties.Match('email').Count -gt 0 -and $u.email) { ($u.email -as [string]).ToLower().Trim() } else { '' }
        $vertr    = if ($u.PSObject.Properties.Match('vertrauensinstanz_gkz').Count -gt 0 -and $u.vertrauensinstanz_gkz) { ($u.vertrauensinstanz_gkz -as [string]).Trim() } else { '' }
        $kro_old  = if ($u.PSObject.Properties.Match('kirchliche_stelle_gkz').Count -gt 0 -and $u.kirchliche_stelle_gkz) { ($u.kirchliche_stelle_gkz -as [string]).Trim() } else { '' }

        $DekanatName = ""
        $KirchenbezirkName = ""
        $GemeindeName = ""
        $kro_new = $kro_old
        $OREID = ""

        if ($kro_old -and $gemeinden_hash.ContainsKey($kro_old)) {
            $g = $gemeinden_hash[$kro_old]
            $DekanatName        = ($g.ORENAMEEBENE4 -as [string]).Trim()
            $KirchenbezirkName = ($g.ORENAMEEBENE3 -as [string]).Trim()
            $GemeindeName       = ($g.ORENAME       -as [string]).Trim()
            if ($g.PSObject.Properties.Match('kro_new').Count -gt 0 -and $g.kro_new) { $kro_new = $g.kro_new }
            $OREID = ($g.OREID -as [string]).Trim()
        }

        # Use 3-way wave function
        $waveInfo = Get-WaveInfo3Way -KRO_GKZ $kro_old -Vertrauensinstanz $vertr -DekanatName $DekanatName

        # PILOT OVERRIDE
        $kro_old_canon = CanonKRO $kro_old 3
        if ($kro_old_canon -and $PilotHash.ContainsKey($kro_old_canon)) { 
            $waveInfo.Wave = "Pilot"
            $waveInfo.Source = "Pilot_Override"
            $waveInfo.Warning = ""
        }

        # License/Cloud info lookup
        $licFound = $null
        if ($email) { 
            $licFound = $licenseByUPN[$email.ToLower()] 
        }
        
        # Extract cloud properties from Lizenzreport
        $Cloud_AD_Domain = ""
        $Cloud_OU = ""
        $Cloud_LizenzName = ""
        $Cloud_Company = ""
        $Cloud_ObjectID = ""
        $Cloud_Displayname = ""
        $Cloud_DisplaynameEmail = ""
        $Cloud_mailnickname = ""
        $Cloud_samaccountname = ""
        $Cloud_extensionattribute4 = ""
        $Cloud_AP_Typ = ""
        $Cloud_Mailbox_Typ = ""
        $Cloud_Routing = ""
        $Cloud_Dekanat = ""
        
        if ($licFound) {
            $Cloud_AD_Domain = Get-PropValue $licFound 'AD_Domain'
            $Cloud_OU = Get-PropValue $licFound 'OU_AD'
            $Cloud_LizenzName = Get-PropValue $licFound 'LizenzName_Graph'
            $Cloud_Company = Get-PropValue $licFound 'Company_AD'
            $Cloud_ObjectID = Get-PropValue $licFound 'Az_ObjectID'
            $Cloud_Displayname = Get-PropValue $licFound 'Name'
            $Cloud_DisplaynameEmail = Get-PropValue $licFound 'Mail'
            $Cloud_mailnickname = Get-PropValue $licFound 'mailNickname'
            $Cloud_samaccountname = Get-PropValue $licFound 'Samaccountname'
            $Cloud_extensionattribute4 = Get-PropValue $licFound 'employeeID'
            $Cloud_AP_Typ = Get-PropValue $licFound 'AP_Typ'
            $Cloud_Mailbox_Typ = Get-PropValue $licFound 'Mailbox_Typ'
            $Cloud_Routing = Get-PropValue $licFound 'Routing'
            $Cloud_Dekanat = Get-PropValue $licFound 'Dekanat'
        }

        # Lookup password
        $password = ""
        if ($email -and $passwordByEmail.ContainsKey($email.ToLower())) {
            $password = $passwordByEmail[$email.ToLower()]
        }

        # Check if AUSGESCHIEDEN (deactivated)
        $isAusgeschieden = $false
        if ($Cloud_OU -match '_Deaktivierte-Benutzer' -or $Cloud_OU -match 'ELKW-Benutzer/_Deaktivierte') {
            $isAusgeschieden = $true
        }

        # === BUSINESS LOGIC: MIGRATION STATUS & SEAFILE ===
        $MailboxMigration_Internal = $false
        $SEAFILE = $false
        $MigrationStatus = ""
        
        # Check last login from Cancom
        $lastLoginRaw = Get-PropValue $u 'letzte_anmeldung'
        $isInactive12Plus = $false
        if ($lastLoginRaw -and $lastLoginRaw -ne 'NULL') {
            try {
                $lastLogin = [datetime]$lastLoginRaw
                if (((Get-Date) - $lastLogin).TotalDays -ge 365) {
                    $isInactive12Plus = $true
                }
            } catch { }
        }

        # Business customer detection
        $isBusiness = $false
        if ($Cloud_AP_Typ -and $Cloud_AP_Typ -ne '' -and $Cloud_AP_Typ -ne 'NULL') {
            if ($Cloud_AP_Typ -notmatch 'DGM') {
                $isBusiness = $true
            }
        }

        # Check AD_Domain for elkw.local (deactivated but not in OU yet)
        $isElkwLocal = ($Cloud_AD_Domain -eq 'elkw.local')

        # Decision logic
        if ($isAusgeschieden) {
            # Deactivated user
            $MailboxMigration_Internal = $false
            $SEAFILE = $false
            $MigrationStatus = "AUSGESCHIEDEN"
        } elseif ($isElkwLocal) {
            # elkw.local domain → no migration (likely deactivated)
            $MailboxMigration_Internal = $false
            $SEAFILE = $false
            $MigrationStatus = "ELKW_LOCAL_DOMAIN"
        } elseif ($isBusiness) {
            # Business customer → SEAFILE ONLY
            $MailboxMigration_Internal = $false
            $SEAFILE = $true
            $MigrationStatus = "BUSINESS_SEAFILE_ONLY"
        } else {
            # Non-business → check activity
            if ($isInactive12Plus) {
                $MailboxMigration_Internal = $false
                $SEAFILE = $false
                $MigrationStatus = "INACTIVE_12M"
            } else {
                $MailboxMigration_Internal = $true
                $SEAFILE = $true
                $MigrationStatus = "ACTIVE"
            }
        }

        # PROVISIONING STATUS LOGIC
        $provStatus = "IGNORE"
        
        if ($MailboxMigration_Internal) {
            # Check Quick2 status first
            $quick2Status = $migrationStatusByEmail[$email.ToLower()]
            
            if ($quick2Status -eq 'Synced') {
                # Already migrated successfully
                $provStatus = "MIGRATED"
            } elseif ($quick2Status -eq 'Failed' -or (-not $quick2Status)) {
                # Failed or not in Quick2 → check if needs provisioning
                if (-not [string]::IsNullOrWhiteSpace($Cloud_AP_Typ) -and $Cloud_AP_Typ -match 'DGM') {
                    # Has DGM license → ready for migration
                    if ([string]::IsNullOrWhiteSpace($password)) { 
                        $provStatus = "MISSING_PW" 
                    } else { 
                        # Check wave to determine status
                        $wRaw = $waveInfo.Wave -replace 'W|Pilot',''
                        if ($wRaw -match '^\d+$') {
                            $waveNum = [int]$wRaw
                            if ($waveNum -lt $CurrentBatchThreshold) {
                                # Past wave → should be migrated
                                $provStatus = if ($quick2Status -eq 'Failed') { "NACHZUEGLER_FAILED" } else { "NACHZUEGLER_PENDING" }
                            } else {
                                # Future wave → OK for later
                                $provStatus = "OK_FUTURE_WAVE"
                            }
                        } elseif ($waveInfo.Wave -eq 'Pilot') {
                            $provStatus = if ($quick2Status -eq 'Failed') { "NACHZUEGLER_FAILED" } else { "NACHZUEGLER_PENDING" }
                        } else {
                            $provStatus = "OK"
                        }
                    }
                } else {
                    # No DGM license → needs AD account creation
                    $provStatus = "UNPROVISIONED"
                }
            }
        }

        # ROUTING VALIDATION (ONLY if domain NOT elkw.local AND mailbox exists)
        $ExpectedRouting = ""
        $RoutingStatus = ""
        
        if ($MailboxMigration_Internal -and 
            $provStatus -ne "MIGRATED" -and 
            $Cloud_AD_Domain -ne 'elkw.local' -and
            $licFound) {
            
            $wRaw = $waveInfo.Wave -replace 'W|Pilot',''
            if ($wRaw -match '^\d+$') {
                $waveNum = [int]$wRaw
                if ($waveNum -lt $CurrentBatchThreshold) {
                    $ExpectedRouting = "NOT RouteToKopano"
                    if ($Cloud_Routing -eq "RouteToKopano") {
                        $RoutingStatus = "ROUTING_FAILURE_SHOULD_NOT_BE_KOPANO"
                    } else {
                        $RoutingStatus = "ROUTING_OK"
                    }
                } elseif ($waveNum -ge $CurrentBatchThreshold) {
                    $ExpectedRouting = "RouteToKopano"
                    if ($Cloud_Routing -ne "RouteToKopano") {
                        $RoutingStatus = "ROUTING_FAILURE_SHOULD_BE_KOPANO"
                    } else {
                        $RoutingStatus = "ROUTING_OK"
                    }
                }
            } elseif ($waveInfo.Wave -eq "Pilot") {
                $ExpectedRouting = "NOT RouteToKopano"
                if ($Cloud_Routing -eq "RouteToKopano") {
                    $RoutingStatus = "ROUTING_FAILURE_SHOULD_NOT_BE_KOPANO"
                } else {
                    $RoutingStatus = "ROUTING_OK"
                }
            }
        }

        # DEKANAT MISMATCH CHECK (for OKR.ELK-WUE.DE domain only)
        $DekanatMismatch = ""
        if ($Cloud_AD_Domain -eq "OKR.ELK-WUE.DE") {
            if ($Cloud_Dekanat -and $DekanatName -and (Normalize-Key $Cloud_Dekanat) -ne (Normalize-Key $DekanatName)) {
                $DekanatMismatch = "MISMATCH: Cloud=$Cloud_Dekanat vs KRO=$DekanatName"
            }
        }

        # Build the user object
        $MailboxMigration_Output = if ($MailboxMigration_Internal) { "YES" } else { "NO" }
        
        $row = [PSCustomObject]@{
            Wave                   = $waveInfo.Wave
            Wave_Source            = $waveInfo.Source
            Wave_Warning           = $waveInfo.Warning
            MigrationStatus        = $MigrationStatus
            ProvisioningStatus     = $provStatus
            RoutingStatus          = $RoutingStatus
            Routing_Current        = $Cloud_Routing
            Routing_Expected       = $ExpectedRouting
            DekanatMismatch        = $DekanatMismatch
            email                  = $email
            Password               = $password
            Vorname                = Get-PropValue $u 'vorname'
            Nachname               = Get-PropValue $u 'nachname'
            Displayname            = if ($Cloud_Displayname) { $Cloud_Displayname } else { Get-PropValue $u 'displayname' }
            Displayname_Email      = if ($Cloud_DisplaynameEmail) { $Cloud_DisplaynameEmail } else { Get-PropValue $u 'displayname_email' }
            mailnickname           = if ($Cloud_mailnickname) { $Cloud_mailnickname } else { Get-PropValue $u 'mailnickname' }
            samaccountname         = if ($Cloud_samaccountname) { $Cloud_samaccountname } else { Get-PropValue $u 'samaccountname' }
            Gemeinde               = $GemeindeName
            Dekanat                = $DekanatName
            Dekanat_Cloud          = $Cloud_Dekanat
            Kirchenbezirk          = $KirchenbezirkName
            ObjectID               = if ($Cloud_ObjectID) { $Cloud_ObjectID } else { Get-PropValue $u 'objectid' }
            OU                     = if ($Cloud_OU) { $Cloud_OU } else { Get-PropValue $u 'ou' }
            extensionattribute4    = if ($Cloud_extensionattribute4) { $Cloud_extensionattribute4 } else { Get-PropValue $u 'extensionattribute4' }
            AP_Typ                 = if ($Cloud_AP_Typ) { $Cloud_AP_Typ } else { Get-PropValue $u 'arbeitsplatztyp' }
            AD_Domain              = if ($Cloud_AD_Domain) { $Cloud_AD_Domain } else { Get-PropValue $u 'ad_domain' }
            MailboxMigration       = $MailboxMigration_Output
            SEAFILE                = if ($SEAFILE) { "TRUE" } else { "FALSE" }
            vertrauensinstanz_gkz  = $vertr
            gkz                    = $kro_new
            gkz_old                = $kro_old
            OREID                  = $OREID
            LicenseAssigned        = $Cloud_LizenzName
            Mailbox_Typ            = $Cloud_Mailbox_Typ
            Company                = $Cloud_Company
            letzte_anmeldung       = Get-PropValue $u 'letzte_anmeldung'
            letzte_anmeldung_recent = -not $isInactive12Plus
        }
        $masterUsers.Add($row) | Out-Null
    }
    catch {
        Write-Warning "User enrichment error [$($u.email)]: $_"
    }
}

# 8. --- BUILD MASTER SHARED MAILBOXES LIST WITH FULL KRO DATA ---
$masterSharedMailboxes = New-Object System.Collections.Generic.List[psobject]

if ($sharedMailboxes.Count -gt 0) {
    Write-Host "Processing $($sharedMailboxes.Count) shared mailboxes..."
    
    foreach ($smb in $sharedMailboxes) {
        try {
            $email = if ($smb.PSObject.Properties.Match('funktionspostfach').Count -gt 0 -and $smb.funktionspostfach) { 
                ($smb.funktionspostfach -as [string]).ToLower().Trim() 
            } else { '' }
            
            if (-not $email) { continue }

            # Map kirchliche_stelle_gkz to get ORENAME and other data
            $gkz_raw = if ($smb.PSObject.Properties.Match('kirchliche_stelle_gkz').Count -gt 0 -and $smb.kirchliche_stelle_gkz) { 
                ($smb.kirchliche_stelle_gkz -as [string]).Trim() 
            } else { '' }

            $gkz = CanonKRO $gkz_raw 3
            $DekanatName = ""
            $KirchenbezirkName = ""
            $GemeindeName = ""
            $OREID = ""

            if ($gkz -and $gemeinden_hash.ContainsKey($gkz)) {
                $g = $gemeinden_hash[$gkz]
                $DekanatName        = ($g.ORENAMEEBENE4 -as [string]).Trim()
                $KirchenbezirkName = ($g.ORENAMEEBENE3 -as [string]).Trim()
                $GemeindeName       = ($g.ORENAME       -as [string]).Trim()
                $OREID              = ($g.OREID -as [string]).Trim()
            }

            # Use 3-way wave function (no Vertrauensinstanz for shared mailboxes)
            $waveInfo = Get-WaveInfo3Way -KRO_GKZ $gkz_raw -Vertrauensinstanz "" -DekanatName $DekanatName
            
            # PILOT OVERRIDE for Shared Mailboxes
            if ((CanonKRO $gkz 3) -and $PilotHash.ContainsKey((CanonKRO $gkz 3))) { 
                $waveInfo.Wave = "Pilot"
                $waveInfo.Source = "Pilot_Override"
                $waveInfo.Warning = ""
            }

            $row = [PSCustomObject]@{
                funktionspostfach   = $email
                Password            = $SharedMailboxPassword
                gkz                 = $gkz
                Wave                = $waveInfo.Wave
                Wave_Source         = $waveInfo.Source
                Wave_Warning        = $waveInfo.Warning
                Gemeinde            = $GemeindeName
                Dekanat             = $DekanatName
                Kirchenbezirk       = $KirchenbezirkName
                OREID               = $OREID
                ticket_status       = Get-PropValue $smb 'ticket_status'
                besitzermailplus    = Get-PropValue $smb 'besitzermailplus'
            }
            $masterSharedMailboxes.Add($row) | Out-Null
        } 
        catch { 
            Write-Warning "Shared Mailbox enrichment error [$($smb.funktionspostfach)]: $_" 
        }
    }
}

# 9. --- BUILD KRO → ACTIVE USERS INDEX ---
Write-Host "Building KRO to active users index..."
$usersByKRO = @{}
foreach ($u in $masterUsers) {
    if ($u.MailboxMigration -ne 'YES') { continue }
    $keys = @()
    if ($u.PSObject.Properties.Match('gkz').Count -gt 0 -and $u.gkz) { 
        $keys += (CanonKRO $u.gkz 3) 
    }
    if ($u.PSObject.Properties.Match('gkz_old').Count -gt 0 -and $u.gkz_old) { 
        $keys += ($u.gkz_old -split ',' | ForEach-Object { CanonKRO $_ 3 } | Where-Object { $_ }) 
    }
    foreach ($k in ($keys | Sort-Object -Unique)) {
        if (-not $usersByKRO.ContainsKey($k)) { 
            $usersByKRO[$k] = New-Object System.Collections.ArrayList 
        }
        [void]$usersByKRO[$k].Add($u)
    }
}

# 10. --- BUILD KRO TO GEMEINDE MULTI-MAP ---
Write-Host "Building KRO to Gemeinde mapping..."
$kroToGemeindeMulti = @{}
foreach ($g in $gemeindenCT) {
    $allKeys = @()
    if ($g.KRO) { $allKeys += (CanonKRO $g.KRO 3) }
    if ($g.PSObject.Properties.Match('kro_alt').Count -gt 0 -and $g.kro_alt) {
        $allKeys += ($g.kro_alt -split ',' | ForEach-Object { CanonKRO $_ 3 } | Where-Object { $_ })
    }
    $allKeys = $allKeys | Where-Object { $_ } | Sort-Object -Unique
    
    foreach ($k in $allKeys) {
        if (-not $kroToGemeindeMulti.ContainsKey($k)) { 
            $kroToGemeindeMulti[$k] = @() 
        }
        $kroToGemeindeMulti[$k] += $g
    }
}

# 11. --- EXPORT MASTER USER FILES ---
Write-Host "Exporting master user files..."
$MasterUsersCsv = Join-Path $OutDir_Master "Master_User_Migration.csv"
$masterUsers | Export-Csv -Path $MasterUsersCsv -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation
Strip-Quotes $MasterUsersCsv

# Export AUSGESCHIEDEN users separately
$ausgeschiedene = @($masterUsers | Where-Object { $_.MigrationStatus -eq 'AUSGESCHIEDEN' })
if ($ausgeschiedene.Count -gt 0) {
    $ausFile = Join-Path $OutDir_Master "Ausgeschiedene_Users.csv"
    $ausgeschiedene | Export-Csv -Path $ausFile -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation
    Strip-Quotes $ausFile
    Write-Host "  Exported $($ausgeschiedene.Count) AUSGESCHIEDEN users"
}

# Export Dekanat mismatch users (OKR.ELK-WUE.DE only)
$dekanatMismatches = @($masterUsers | Where-Object { $_.DekanatMismatch })
if ($dekanatMismatches.Count -gt 0) {
    $dmFile = Join-Path $OutDir_Master "Dekanat_Mismatch_OKR.csv"
    $dekanatMismatches | Select-Object email, Displayname, Dekanat, Dekanat_Cloud, DekanatMismatch, AD_Domain | 
        Export-Csv -Path $dmFile -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation
    Strip-Quotes $dmFile
    Write-Host "  Exported $($dekanatMismatches.Count) Dekanat mismatches (OKR domain)"
}

# Export master shared mailboxes if any exist
if ($masterSharedMailboxes.Count -gt 0) {
    Write-Host "Exporting master shared mailboxes file..."
    $MasterSharedMailboxesCsv = Join-Path $OutDir_Master "Master_SharedMailboxes_Migration.csv"
    $masterSharedMailboxes | Export-Csv -Path $MasterSharedMailboxesCsv -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation
    Strip-Quotes $MasterSharedMailboxesCsv
}

# --- PROVISIONING & MIGRATION STATUS ANALYSIS ---
Write-Host "Analyzing provisioning and migration status..."

$provisionCols = @(
    'Wave','email','Password','Vorname','Nachname','Displayname','mailnickname','samaccountname',
    'Gemeinde','Dekanat','Kirchenbezirk','ObjectID','AP_Typ','AD_Domain','MailboxMigration','MigrationStatus','gkz','gkz_old'
)

# === UNPROVISIONED USERS (Need AD Account Creation) ===
$createCandidates = @($masterUsers | Where-Object { 
    $_.ProvisioningStatus -eq 'UNPROVISIONED' -and 
    $_.MigrationStatus -ne 'AUSGESCHIEDEN' 
})
Write-Host "  Unprovisioned users (need AD creation): $($createCandidates.Count)"

# Nachzügler (W1-W9) - FAILED or NOT IN QUICK2
$nachzueglerUnprov = @($createCandidates | Where-Object { 
    $wRaw = $_.Wave -replace 'W|Pilot',''
    if ($wRaw -match '^\d+$') { [int]$wRaw -lt $CurrentBatchThreshold } else { $false }
})
if ($nachzueglerUnprov.Count -gt 0) {
    $nFile = Join-Path $OutDir_Master "Unprovisioned_Users_AD_Create_Nachzuegler.csv"
    $nachzueglerUnprov | Select-Object $provisionCols | Export-Csv -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation -Path $nFile
    Strip-Quotes $nFile
    Write-Host "    └─ Nachzügler (W1-W9): $($nachzueglerUnprov.Count)"
}

# Current Waves (W10-W16)
foreach ($i in $CurrentBatchThreshold..16) {
    $w = "W$i"
    $wUsers = @($createCandidates | Where-Object { $_.Wave -eq $w })
    if ($wUsers.Count -gt 0) {
        $wFile = Join-Path $OutDir_Master "Unprovisioned_Users_AD_Create_$w.csv"
        $wUsers | Select-Object $provisionCols | Export-Csv -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation -Path $wFile
        Strip-Quotes $wFile
        Write-Host "    └─ $w`: $($wUsers.Count)"
    }
}

# Pilot
$wPilot = @($createCandidates | Where-Object { $_.Wave -eq 'Pilot' })
if ($wPilot.Count -gt 0) {
    $pFile = Join-Path $OutDir_Master "Unprovisioned_Users_AD_Create_Pilot.csv"
    $wPilot | Select-Object $provisionCols | Export-Csv -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation -Path $pFile
    Strip-Quotes $pFile
    Write-Host "    └─ Pilot: $($wPilot.Count)"
}

# === NACHZÜGLER (Past Wave, Has AD Account, NOT MIGRATED) ===
$nachzuegler = @($masterUsers | Where-Object {
    $wRaw = $_.Wave -replace 'W|Pilot',''
    $isPastWave = if ($wRaw -match '^\d+$') { [int]$wRaw -lt $CurrentBatchThreshold } elseif ($_.Wave -eq 'Pilot') { $true } else { $false }
    
    $isPastWave -and 
    ($_.ProvisioningStatus -match 'NACHZUEGLER') -and
    ($_.MigrationStatus -ne 'AUSGESCHIEDEN')
})

if ($nachzuegler.Count -gt 0) {
    Write-Host "  Nachzügler (W1-9/Pilot, ready for batch): $($nachzuegler.Count)"
    
    # Full detail export
    $nzFile = Join-Path $OutDir_Master "Nachzuegler_W1-9_Ready.csv"
    $nachzuegler | Select-Object $provisionCols | 
        Export-Csv -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation -Path $nzFile
    Strip-Quotes $nzFile
    
    # Quick2 Format (Email-Only for Batch Import)
    $qFile = Join-Path $OutDir_Master "Nachzuegler_Quick2_Batch.csv"
    $nachzuegler | Select-Object @{N='EmailAddress';E={$_.email}} | 
        Export-Csv -Delimiter ',' -Encoding UTF8 -NoTypeInformation -Path $qFile
    Strip-Quotes $qFile
    Write-Host "    └─ Quick2 CSV generated (email-only batch format)"
}

# === ROUTING FAILURES REPORT (USERS + SHARED MAILBOXES) ===
Write-Host "Analyzing routing configuration..."

# User routing failures (ONLY MailboxMigration=YES AND not yet migrated AND domain NOT elkw.local AND mailbox exists)
$routingFailuresUsers = @($masterUsers | Where-Object { 
    ($_.MailboxMigration -eq 'YES') -and
    ($_.ProvisioningStatus -ne 'MIGRATED') -and
    ($_.AD_Domain -ne 'elkw.local') -and
    ($licenseByUPN.ContainsKey($_.email.ToLower())) -and
    ($_.RoutingStatus -match 'ROUTING_FAILURE')
})

# Shared mailbox routing failures - Check Lizenzreport for routing (ONLY if NOT elkw.local AND mailbox exists)
$routingFailuresShared = @()
foreach ($smb in $masterSharedMailboxes) {
    $email = $smb.funktionspostfach.ToLower().Trim()
    $licFound = $licenseByUPN[$email]
    
    if ($licFound) {
        $ad_domain = Get-PropValue $licFound 'AD_Domain'
        
        # Skip if elkw.local
        if ($ad_domain -eq 'elkw.local') { continue }
        
        $routing = Get-PropValue $licFound 'Routing'
        $wRaw = $smb.Wave -replace 'W|Pilot',''
        
        $expectedRouting = ""
        if ($wRaw -match '^\d+$') {
            $waveNum = [int]$wRaw
            if ($waveNum -lt $CurrentBatchThreshold) {
                $expectedRouting = "NOT RouteToKopano"
                if ($routing -eq "RouteToKopano") {
                    $routingFailuresShared += [PSCustomObject]@{
                        Wave             = $smb.Wave
                        email            = $smb.funktionspostfach
                        Routing_Current  = $routing
                        Routing_Expected = $expectedRouting
                        Gemeinde         = $smb.Gemeinde
                        AD_Domain        = $ad_domain
                        Type             = "SharedMailbox"
                    }
                }
            } elseif ($waveNum -ge $CurrentBatchThreshold) {
                $expectedRouting = "RouteToKopano"
                if ($routing -ne "RouteToKopano") {
                    $routingFailuresShared += [PSCustomObject]@{
                        Wave             = $smb.Wave
                        email            = $smb.funktionspostfach
                        Routing_Current  = $routing
                        Routing_Expected = $expectedRouting
                        Gemeinde         = $smb.Gemeinde
                        AD_Domain        = $ad_domain
                        Type             = "SharedMailbox"
                    }
                }
            }
        } elseif ($smb.Wave -eq "Pilot") {
            $expectedRouting = "NOT RouteToKopano"
            if ($routing -eq "RouteToKopano") {
                $routingFailuresShared += [PSCustomObject]@{
                    Wave             = $smb.Wave
                    email            = $smb.funktionspostfach
                    Routing_Current  = $routing
                    Routing_Expected = $expectedRouting
                    Gemeinde         = $smb.Gemeinde
                    AD_Domain        = $ad_domain
                    Type             = "SharedMailbox"
                }
            }
        }
    }
}

# Combine user and shared mailbox routing failures
$routingFailuresAll = @()
foreach ($u in $routingFailuresUsers) {
    $routingFailuresAll += [PSCustomObject]@{
        Wave             = $u.Wave
        email            = $u.email
        Routing_Current  = $u.Routing_Current
        Routing_Expected = $u.Routing_Expected
        Gemeinde         = $u.Gemeinde
        AD_Domain        = $u.AD_Domain
        Type             = "User"
    }
}
$routingFailuresAll += $routingFailuresShared

if ($routingFailuresAll.Count -gt 0) {
    $rfFile = Join-Path $OutDir_Master "Routing_Failures.csv"
    $routingFailuresAll | Export-Csv -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation -Path $rfFile
    Strip-Quotes $rfFile
    Write-Host "  ⚠️  ROUTING FAILURES: $($routingFailuresAll.Count) total ($($routingFailuresUsers.Count) users + $($routingFailuresShared.Count) shared mailboxes)"
    
    # Split by category
    $pastWaveRoutingIssues = @($routingFailuresAll | Where-Object {
        $wRaw = $_.Wave -replace 'W|Pilot',''
        if ($wRaw -match '^\d+$') { [int]$wRaw -lt $CurrentBatchThreshold } else { $_.Wave -eq 'Pilot' }
    })
    $futureWaveRoutingIssues = @($routingFailuresAll | Where-Object {
        $wRaw = $_.Wave -replace 'W|Pilot',''
        if ($wRaw -match '^\d+$') { [int]$wRaw -ge $CurrentBatchThreshold } else { $false }
    })
    
    if ($pastWaveRoutingIssues.Count -gt 0) {
        Write-Host "    └─ W1-W9/Pilot still routing to Kopano: $($pastWaveRoutingIssues.Count)"
    }
    if ($futureWaveRoutingIssues.Count -gt 0) {
        Write-Host "    └─ W10+ NOT routing to Kopano: $($futureWaveRoutingIssues.Count)"
    }
} else {
    Write-Host "  ✅ All routing configurations are correct"
}

# 12. --- WAVE EXPORTS (FILTERED BY MailboxMigration=YES) ---
Write-Host "Creating wave-specific exports..."
$wavesOrdered = (1..14 | ForEach-Object { "W$_" }) + "Pilot"
$joinedRows = @()
$activeRows = @()
$inactiveRows = @()

foreach ($w in $wavesOrdered) {
    $waveNum = if ($w -eq "Pilot") { "Pilot" } else { $w.Substring(1) }
    Write-Host "  Processing $w..."
    
    # --- Export Wave Users (ALL USERS for full CSV) ---
    $usersWaveAll = @($masterUsers | Where-Object { 
        $_.Wave -eq $w -and 
        $_.MigrationStatus -ne 'AUSGESCHIEDEN'
    })
    
    # --- Export Wave Users (MIGRATION FILES: ONLY MailboxMigration=YES) ---
    $usersWaveMigration = @($usersWaveAll | Where-Object { $_.MailboxMigration -eq 'YES' })
    
    # Full wave CSV (all users)
    if ($usersWaveAll.Count -gt 0) {
        $usersCsv = Join-Path $OutDir_Waves ("Wave_{0}_Users.csv" -f $w)
        $usersWaveAll | Export-Csv -Path $usersCsv -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation
        Strip-Quotes $usersCsv
    }

    # Migration-specific files (ONLY MailboxMigration=YES)
    if ($usersWaveMigration.Count -gt 0) {
        $fSuffix = if ($w -eq "Pilot") { 
            "Pilot" 
        } else { 
            "{0:D2}" -f [int]$w.Substring(1) 
        }
        
        # 1. Initial file with Password - semicolon separated
        $initialFile = Join-Path $OutDir_Waves ("Wave_{0}_initial_with_PW.csv" -f $fSuffix)
        $usersWaveMigration | Select-Object email, displayname_email, Password | 
            Export-Csv -Path $initialFile -Delimiter ';' -Encoding UTF8 -NoTypeInformation
        Strip-Quotes $initialFile

        # 2. Cancom file - semicolon separated
        $cancomFile = Join-Path $OutDir_Waves ("Wave_{0}_Cancom.csv" -f $fSuffix)
        $usersWaveMigration | Select-Object email, Password | 
            Export-Csv -Path $cancomFile -Delimiter ';' -Encoding UTF8 -NoTypeInformation
        Strip-Quotes $cancomFile

        # 3. Migration batch file - comma separated
        $migrationBatchFile = Join-Path $OutDir_Waves ("Wave_{0}_migrationbatch.csv" -f $fSuffix)
        $migrationBatchData = $usersWaveMigration | ForEach-Object {
            [PSCustomObject]@{
                emailaddress = $_.email
                username     = $_.email
                Password     = $_.Password
            }
        }
        $migrationBatchData | Export-Csv -Path $migrationBatchFile -Delimiter ',' -Encoding UTF8 -NoTypeInformation
        Strip-Quotes $migrationBatchFile
    }

    # --- Export Wave Shared Mailboxes ---
    if ($masterSharedMailboxes.Count -gt 0) {
        $sharedMailboxesWave = @($masterSharedMailboxes | Where-Object { $_.Wave -eq $w })
        if ($sharedMailboxesWave.Count -gt 0) {
            $sharedCsvPath = Join-Path $OutDir_Waves ("Wave_{0}_SharedMailboxes.csv" -f $w)
            $sharedMailboxesWave | Export-Csv -Path $sharedCsvPath -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8
            Strip-Quotes $sharedCsvPath

            $fSuffix = if ($w -eq "Pilot") { 
                "Pilot" 
            } else { 
                "{0:D2}" -f [int]$w.Substring(1) 
            }
            
            # 1. Initial file with Password - semicolon separated
            $sharedInitialFile = Join-Path $OutDir_Waves ("Wave_{0}_SharedMailbox_initial_with_PW.csv" -f $fSuffix)
            $sharedMailboxesWave | Select-Object funktionspostfach, @{N='displayname_email';E={$_.funktionspostfach}}, Password | 
                Export-Csv -Path $sharedInitialFile -Delimiter ';' -Encoding UTF8 -NoTypeInformation
            Strip-Quotes $sharedInitialFile

            # 2. Cancom file - semicolon separated
            $sharedCancomFile = Join-Path $OutDir_Waves ("Wave_{0}_SharedMailbox_Cancom.csv" -f $fSuffix)
            $sharedMailboxesWave | Select-Object @{N='email';E={$_.funktionspostfach}}, Password | 
                Export-Csv -Path $sharedCancomFile -Delimiter ';' -Encoding UTF8 -NoTypeInformation
            Strip-Quotes $sharedCancomFile

            # 3. Migration batch file - comma separated
            $sharedMigrationBatchFile = Join-Path $OutDir_Waves ("Wave_{0}_SharedMailbox_migrationbatch.csv" -f $fSuffix)
            $sharedMigrationBatchData = $sharedMailboxesWave | ForEach-Object {
                [PSCustomObject]@{
                    emailaddress = $_.funktionspostfach
                    username     = $_.funktionspostfach
                    Password     = $_.Password
                }
            }
            $sharedMigrationBatchData | Export-Csv -Path $sharedMigrationBatchFile -Delimiter ',' -Encoding UTF8 -NoTypeInformation
            Strip-Quotes $sharedMigrationBatchFile
        }
    }

    # --- Create Gemeinden Export for Wave (based on ALL users, not just migration) ---
    $kroSet = @($usersWaveAll | ForEach-Object {
        $tmp = @()
        if ($_.gkz) { $tmp += (CanonKRO $_.gkz 3) }
        if ($_.gkz_old) { 
            $tmp += ($_.gkz_old -split ',' | ForEach-Object { CanonKRO $_ 3 }) 
        }
        $tmp
    }) | Where-Object { $_ } | Sort-Object -Unique

    $gemeindenExport = @()
    foreach ($k in $kroSet) {
        if (-not $kroToGemeindeMulti.ContainsKey($k)) { continue }
        
        foreach ($g in $kroToGemeindeMulti[$k]) {
            $allKeys = @()
            if ($g.KRO) { $allKeys += (CanonKRO $g.KRO 3) }
            if ($g.PSObject.Properties.Match('kro_alt').Count -gt 0 -and $g.kro_alt) {
                $allKeys += ($g.kro_alt -split ',' | ForEach-Object { CanonKRO $_ 3 })
            }
            $allKeys = $allKeys | Where-Object { $_ } | Sort-Object -Unique

            # Check if there are active users
            $hasActive = $false
            foreach ($kk in $allKeys) {
                if ($usersByKRO.ContainsKey($kk) -and $usersByKRO[$kk].Count -gt 0) {
                    $hasActive = $true
                    break
                }
            }
            $aktive = if ($hasActive) { 'JA' } else { 'NEIN' }

            # Get planned week and date
            $plannedWeek = ""
            $datum = ""
            $dek = $g.ORENAMEEBENE4
            $key = ""
            if ($dek) { $key = $dek.Trim().ToLower() }
            if ($key -and $Welle_Dekanat_hash.ContainsKey($key)) {
                $plannedWeek = $Welle_Dekanat_hash[$key].Planned_Week
                $datum = $Welle_Dekanat_hash[$key].Datum
            }

            # Get CT instances and designations
            $ctInstances = @()
            $ctDesignations = @()
            foreach ($lookupKey in $allKeys) {
                if ($ct_hash.ContainsKey($lookupKey)) {
                    $ctEntry = $ct_hash[$lookupKey]
                    $inst = $ctEntry.CT_Instanz
                    $des = $ctEntry.designation
                    if ($inst -and $ctInstances -notcontains $inst) { $ctInstances += $inst }
                    if ($des -and $ctDesignations -notcontains $des) { $ctDesignations += $des }
                }
            }
            $ctinst = $ctInstances -join ', '
            $ctdes = $ctDesignations -join ', '

            $gemeindenExport += [PSCustomObject]@{
                KRO              = (CanonKRO $g.KRO 3)
                OREID            = $g.OREID
                OREART           = $g.OREART
                ORENAME          = $g.ORENAME
                Gkz_alt          = $g.kro_alt
                Email_all        = $g.All_Emails
                Kirchenbezirk    = $g.ORENAMEEBENE3
                Dekanat          = $g.ORENAMEEBENE4
                Wave             = $w
                Planned_Week     = $plannedWeek
                Datum            = $datum
                CTInstanz        = $ctinst
                CT_Designation   = $ctdes
                OREID_Aktive     = $aktive
                ValidFrom        = $g.OEVGUELTIGAB
                ValidUntil       = $g.OEVGUELTIGBIS
            }
        }
    }

    # Export Gemeinden CSV for wave
    if ($gemeindenExport.Count -gt 0) {
        $gemeindenCsv = Join-Path $OutDir_Waves ("Gemeinden_{0}.csv" -f $w)
        $gemeindenExport | Export-Csv -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -Path $gemeindenCsv
        Strip-Quotes $gemeindenCsv

        # --- Enrich with CT admin and stats ---
        $gRows = Import-Csv $gemeindenCsv -Delimiter $Delimiter -Encoding UTF8
        foreach ($row in $gRows) {
            $cti = ($row.CTInstanz -as [string]).Trim()
            
            # Add CT admin
            $ct_admin = ""
            if ($cti -and $ctAdminsByInstanz.ContainsKey($cti)) {
                $ct_admin = $ctAdminsByInstanz[$cti] -join ', '
            }
            $row | Add-Member -NotePropertyName CT_admin -NotePropertyValue $ct_admin -Force

            # Add CT stats
            if ($cti -and $ctStatsByInstanz.ContainsKey($cti)) {
                $stats = $ctStatsByInstanz[$cti]
                $row | Add-Member -NotePropertyName CT_Gemeindename -NotePropertyValue ($stats.Gemeindename) -Force
                $row | Add-Member -NotePropertyName CT_Personen -NotePropertyValue ($stats.Personen) -Force
                $row | Add-Member -NotePropertyName CT_Aktive -NotePropertyValue ($stats.Aktiv30Tage) -Force
                $row | Add-Member -NotePropertyName CT_Dienste -NotePropertyValue ($stats.Dienste) -Force
                $row | Add-Member -NotePropertyName CT_Dienste_Neu -NotePropertyValue (@($stats.NeueDienste, $stats.NeueBuchungen, $stats.NeueTermine) -join ', ') -Force
            } else {
                $row | Add-Member -NotePropertyName CT_Gemeindename -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName CT_Personen -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName CT_Aktive -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName CT_Dienste -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName CT_Dienste_Neu -NotePropertyValue "" -Force
            }
        }

        # Re-export with ordered columns
        $colorder = @(
            'KRO','OREID','OREART','ORENAME','Gkz_alt','Email_all','Kirchenbezirk','Dekanat',
            'Wave','Planned_Week','Datum','CTInstanz','CT_Designation','OREID_Aktive',
            'ValidFrom','ValidUntil','CT_admin','CT_Gemeindename','CT_Personen','CT_Aktive','CT_Dienste','CT_Dienste_Neu'
        )
        $gRows | Select-Object $colorder | Export-Csv -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -Path $gemeindenCsv
        Strip-Quotes $gemeindenCsv

        # Add to global collections
        $joinedRows += $gRows
        $activeRows += ($gRows | Where-Object { $_.OREID_Aktive -eq "JA" })
        $inactiveRows += ($gRows | Where-Object { $_.OREID_Aktive -ne "JA" })
    }
}

# 13. --- EXPORT JOINED/GLOBAL/ACTIVE/INACTIVE GEMEINDEN FILES ---
Write-Host "Exporting combined Gemeinden files..."
$joinedCsv   = Join-Path $OutDir_Master "Gemeinden_W1-16_ALL.csv"
$activeCsv   = Join-Path $OutDir_Master "Gemeinden_W1-16_AKTIV.csv"
$inactiveCsv = Join-Path $OutDir_Master "Gemeinden_W1-16_INAKTIV.csv"

if ($joinedRows.Count -gt 0) {
    $joinedRows | Export-Csv -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -Path $joinedCsv
    Strip-Quotes $joinedCsv
}

if ($activeRows.Count -gt 0) {
    $activeRows | Export-Csv -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -Path $activeCsv
    Strip-Quotes $activeCsv
}

if ($inactiveRows.Count -gt 0) {
    $inactiveRows | Export-Csv -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8 -Path $inactiveCsv
    Strip-Quotes $inactiveCsv
}

# 14. --- OPTIONAL EXCEL EXPORTS ---
try {
    Import-Module ImportExcel -ErrorAction Stop
    Write-Host "Creating Excel workbooks..."
    
    $MigrationGemeindenPath = Join-Path $RunRoot "Migration_Gemeinden.xlsx"
    $MigrationSeafilePath   = Join-Path $RunRoot "Migration_Seafile.xlsx"
    $MigrationUserPath      = Join-Path $RunRoot "Migration_User.xlsx"
    
    Remove-Item $MigrationGemeindenPath,$MigrationSeafilePath,$MigrationUserPath -ErrorAction SilentlyContinue

    # Migration_Gemeinden.xlsx (active Gemeinden only)
    foreach ($i in 1..14) {
        $w = "W$i"
        $sheetName = "Welle" + "{0:D2}" -f $i
        $gemeindenCsv = Join-Path $OutDir_Waves ("Gemeinden_{0}.csv" -f $w)
        
        if (Test-Path $gemeindenCsv) {
            $gRows = Import-Csv $gemeindenCsv -Delimiter $Delimiter
            $active = $gRows | Where-Object {
                ($_.CT_Aktive -match '^\d+$' -and [int]$_.CT_Aktive -gt 0) -or
                ($_.Email_all -ne $null -and $_.Email_all -ne "")
            }
            if ($active -and $active.Count -gt 0) {
                $active | Export-Excel -Path $MigrationGemeindenPath -WorksheetName $sheetName -AutoSize -TableName $sheetName -NoNumberConversion:$true -TableStyle Medium2 -Append
            }
        }
    }
    
    # Add Pilot sheet
    $gemeindenCsvPilot = Join-Path $OutDir_Waves "Gemeinden_Pilot.csv"
    if (Test-Path $gemeindenCsvPilot) {
        $gRows = Import-Csv $gemeindenCsvPilot -Delimiter $Delimiter
        $active = $gRows | Where-Object {
            ($_.CT_Aktive -match '^\d+$' -and [int]$_.CT_Aktive -gt 0) -or
            ($_.Email_all -ne $null -and $_.Email_all -ne "")
        }
        if ($active -and $active.Count -gt 0) {
            $active | Export-Excel -Path $MigrationGemeindenPath -WorksheetName "Pilot" -AutoSize -TableName "Pilot" -NoNumberConversion:$true -TableStyle Medium2 -Append
        }
    }

    # Migration_Seafile.xlsx (ONLY users with SEAFILE=TRUE and MailboxMigration=YES)
    foreach ($i in 1..14) {
        $w = "W$i"
        $sheetName = "Welle" + "{0:D2}" -f $i
        $usersCsv = Join-Path $OutDir_Waves ("Wave_{0}_Users.csv" -f $w)
        
        if (Test-Path $usersCsv) {
            $users = Import-Csv $usersCsv -Delimiter $Delimiter
            $seafile = $users | Where-Object { 
                ($_.SEAFILE -eq "TRUE") -and
                ($_.MailboxMigration -eq 'YES')
            }
            if ($seafile -and $seafile.Count -gt 0) {
                $seafile | Export-Excel -Path $MigrationSeafilePath -WorksheetName $sheetName -AutoSize -TableName $sheetName -NoNumberConversion:$true -TableStyle Medium2 -Append
            }
        }
    }
    
    # Add Pilot sheet for Seafile
    $usersCsvPilot = Join-Path $OutDir_Waves "Wave_Pilot_Users.csv"
    if (Test-Path $usersCsvPilot) {
        $users = Import-Csv $usersCsvPilot -Delimiter $Delimiter
        $seafile = $users | Where-Object { 
            ($_.SEAFILE -eq "TRUE") -and
            ($_.MailboxMigration -eq 'YES')
        }
        if ($seafile -and $seafile.Count -gt 0) {
            $seafile | Export-Excel -Path $MigrationSeafilePath -WorksheetName "Pilot" -AutoSize -TableName "Pilot" -NoNumberConversion:$true -TableStyle Medium2 -Append
        }
    }

    # Migration_User.xlsx
    foreach ($i in 1..14) {
        $w = "W$i"
        $sheetName = "Welle" + "{0:D2}" -f $i
        $usersCsv = Join-Path $OutDir_Waves ("Wave_{0}_Users.csv" -f $w)
        
        if (Test-Path $usersCsv) {
            $users = Import-Csv $usersCsv -Delimiter $Delimiter
            if ($users -and $users.Count -gt 0) {
                $users | Export-Excel -Path $MigrationUserPath -WorksheetName $sheetName -AutoSize -TableName $sheetName -NoNumberConversion:$true -TableStyle Medium2 -Append
            }
        }
    }
    
    # Add Pilot sheet
    if (Test-Path $usersCsvPilot) {
        $users = Import-Csv $usersCsvPilot -Delimiter $Delimiter
        if ($users -and $users.Count -gt 0) {
            $users | Export-Excel -Path $MigrationUserPath -WorksheetName "Pilot" -AutoSize -TableName "Pilot" -NoNumberConversion:$true -TableStyle Medium2 -Append
        }
    }
    
    Write-Host "  Excel workbooks created successfully"
} catch {
    Write-Warning "Excel export skipped (ImportExcel not available): $_"
}

# 15. --- SUMMARY ---
Write-Host "`n========================================="
Write-Host "✅ DGM Master Builder v4.4 - COMPLETE"
Write-Host "========================================="
Write-Host "`nOutput directory: $RunRoot"
Write-Host "`n📋 Master Files:"
Write-Host "  Core Files:"
Write-Host "    - Master_User_Migration.csv (full detail + statuses)"
Write-Host "    - Master_SharedMailboxes_Migration.csv (with KRO data)" -NoNewline
if ($masterSharedMailboxes.Count -gt 0) { Write-Host " ✓" } else { Write-Host " (empty)" }
Write-Host "    - Ausgeschiedene_Users.csv" -NoNewline
if ($ausgeschiedene.Count -gt 0) { Write-Host " [$($ausgeschiedene.Count) users]" } else { Write-Host " (none)" }
Write-Host "    - Dekanat_Mismatch_OKR.csv" -NoNewline
if ($dekanatMismatches.Count -gt 0) { Write-Host " [$($dekanatMismatches.Count) mismatches]" } else { Write-Host " (none)" }
Write-Host "`n  Provisioning Files (AD Account Creation Needed):"
Write-Host "    - Unprovisioned_Users_AD_Create_Nachzuegler.csv (W1-W9)" -NoNewline
if ($nachzueglerUnprov.Count -gt 0) { Write-Host " [$($nachzueglerUnprov.Count) users]" } else { Write-Host " (none)" }
for ($i = $CurrentBatchThreshold; $i -le 16; $i++) {
    $wf = Join-Path $OutDir_Master "Unprovisioned_Users_AD_Create_W$i.csv"
    if (Test-Path $wf) {
        $cnt = (Import-Csv $wf -Delimiter $Delimiter).Count
        Write-Host "    - Unprovisioned_Users_AD_Create_W$i.csv [$cnt users]"
    }
}
if ($wPilot.Count -gt 0) {
    Write-Host "    - Unprovisioned_Users_AD_Create_Pilot.csv [$($wPilot.Count) users]"
}
Write-Host "`n  Migration Files (Ready for Mailbox Migration):"
Write-Host "    - Nachzuegler_W1-9_Ready.csv" -NoNewline
if ($nachzuegler.Count -gt 0) { 
    Write-Host " [$($nachzuegler.Count) users ready]" 
    Write-Host "    - Nachzuegler_Quick2_Batch.csv (email-only batch format)"
} else { 
    Write-Host " (all up-to-date)" 
}
Write-Host "`n  Status & Validation Reports:"
Write-Host "    - Routing_Failures.csv" -NoNewline
if ($routingFailuresAll.Count -gt 0) { 
    Write-Host " ⚠️  [$($routingFailuresAll.Count) failures: $($routingFailuresUsers.Count) users + $($routingFailuresShared.Count) shared]" 
} else { 
    Write-Host " ✓ (all correct)" 
}
Write-Host "`n  Gemeinden Files:"
Write-Host "    - Gemeinden_W1-16_ALL.csv"
Write-Host "    - Gemeinden_W1-16_AKTIV.csv"
Write-Host "    - Gemeinden_W1-16_INAKTIV.csv"
Write-Host "`n📦 Per Wave Files (W01-W16 + Pilot):"
Write-Host "  User Migration Files (FILTERED BY MailboxMigration=YES):"
Write-Host "    - Wave_WX_Users.csv (full detail - ALL users)"
Write-Host "    - Wave_XX_initial_with_PW.csv (3-column: email;displayname;password) ✓ FILTERED"
Write-Host "    - Wave_XX_Cancom.csv (2-column: email;password) ✓ FILTERED"
Write-Host "    - Wave_XX_migrationbatch.csv (3-column: emailaddress,username,password) ✓ FILTERED"
if ($masterSharedMailboxes.Count -gt 0) {
    Write-Host "  Shared Mailbox Files:"
    Write-Host "    - Wave_WX_SharedMailboxes.csv (with KRO/Gemeinde data)"
    Write-Host "    - Wave_XX_SharedMailbox_[initial|Cancom|batch].csv"
}
Write-Host "  Gemeinden Files:"
Write-Host "    - Gemeinden_WX.csv (CT enriched)"
Write-Host "`n📊 Excel Workbooks (if ImportExcel available):"
Write-Host "  - Migration_Gemeinden.xlsx"
Write-Host "  - Migration_Seafile.xlsx (ONLY MailboxMigration=YES + SEAFILE=TRUE)"
Write-Host "  - Migration_User.xlsx"
Write-Host "`n========================================="
Write-Host "📈 Statistics:"
Write-Host "========================================="
Write-Host "Quick2 Migration Status:"
Write-Host "  - Loaded from Quick2: $($migrationStatusByEmail.Count) accounts"
$syncedCount = ($migrationStatusByEmail.Values | Where-Object { $_ -eq 'Synced' }).Count
$failedCount = ($migrationStatusByEmail.Values | Where-Object { $_ -eq 'Failed' }).Count
Write-Host "  - Synced: $syncedCount"
Write-Host "  - Failed: $failedCount"
Write-Host "`nPasswords:"
Write-Host "  - Loaded from master: $($passwordByEmail.Count)"
Write-Host "  - Users with passwords: $(($masterUsers | Where-Object { $_.Password }).Count)"
Write-Host "`nMigration Status Categories:"
$activeCount = ($masterUsers | Where-Object { $_.MigrationStatus -eq 'ACTIVE' }).Count
$businessCount = ($masterUsers | Where-Object { $_.MigrationStatus -eq 'BUSINESS_SEAFILE_ONLY' }).Count
$inactiveCount = ($masterUsers | Where-Object { $_.MigrationStatus -eq 'INACTIVE_12M' }).Count
$elkwLocalCount = ($masterUsers | Where-Object { $_.MigrationStatus -eq 'ELKW_LOCAL_DOMAIN' }).Count
Write-Host "  ✓ ACTIVE (migrating): $activeCount"
Write-Host "  - BUSINESS (Seafile only): $businessCount"
Write-Host "  - INACTIVE (12+ months): $inactiveCount"
Write-Host "  - ELKW_LOCAL_DOMAIN: $elkwLocalCount"
Write-Host "  - AUSGESCHIEDEN: $($ausgeschiedene.Count)"
Write-Host "`nProvisioning Status:"
$migratedCount = ($masterUsers | Where-Object { $_.ProvisioningStatus -eq 'MIGRATED' }).Count
$okCount = ($masterUsers | Where-Object { $_.ProvisioningStatus -match 'OK' -and $_.ProvisioningStatus -ne 'ROUTING_OK' }).Count
$unprovCount = ($masterUsers | Where-Object { $_.ProvisioningStatus -eq 'UNPROVISIONED' }).Count
$nachzueglerCount = ($masterUsers | Where-Object { $_.ProvisioningStatus -match 'NACHZUEGLER' }).Count
$missingPwCount = ($masterUsers | Where-Object { $_.ProvisioningStatus -eq 'MISSING_PW' }).Count
$ignoreCount = ($masterUsers | Where-Object { $_.ProvisioningStatus -eq 'IGNORE' }).Count
Write-Host "  ✓ MIGRATED (complete): $migratedCount"
Write-Host "  ✓ OK (ready): $okCount"
Write-Host "  ⚠ NACHZUEGLER (past wave, needs migration): $nachzueglerCount"
Write-Host "  ⚠ UNPROVISIONED (need AD): $unprovCount"
Write-Host "  ⚠ MISSING_PW: $missingPwCount"
Write-Host "  - IGNORE (not migrating): $ignoreCount"
Write-Host "`nRouting Status:"
$routingOkCount = ($masterUsers | Where-Object { $_.RoutingStatus -eq 'ROUTING_OK' }).Count
Write-Host "  ✓ Correct routing (non-elkw.local users with mailbox): $routingOkCount"
if ($routingFailuresAll.Count -gt 0) {
    Write-Host "  ⚠️  ROUTING FAILURES: $($routingFailuresAll.Count) ($($routingFailuresUsers.Count) users + $($routingFailuresShared.Count) shared)" -ForegroundColor Red
} else {
    Write-Host "  ✓ No routing failures" -ForegroundColor Green
}
Write-Host "`nDekanat Validation:"
if ($dekanatMismatches.Count -gt 0) {
    Write-Host "  ⚠️  Mismatches (OKR domain): $($dekanatMismatches.Count)" -ForegroundColor Yellow
} else {
    Write-Host "  ✓ All Dekanat assignments correct"
}
Write-Host "`nWave Distribution (MailboxMigration=YES):"
Write-Host "  - Pilot accounts: $(($masterUsers | Where-Object { $_.Wave -eq 'Pilot' -and $_.MailboxMigration -eq 'YES' -and $_.MigrationStatus -ne 'AUSGESCHIEDEN' }).Count) users"
for ($i = 1; $i -le 16; $i++) {
    $cnt = ($masterUsers | Where-Object { $_.Wave -eq "W$i" -and $_.MailboxMigration -eq 'YES' -and $_.MigrationStatus -ne 'AUSGESCHIEDEN' }).Count
    if ($cnt -gt 0) {
        $marker = if ($i -lt $CurrentBatchThreshold) { "✓" } elseif ($i -eq $CurrentBatchThreshold) { "→" } else { " " }
        Write-Host "  $marker W$($i.ToString('D2')): $cnt users"
    }
}
Write-Host "`nShared Mailboxes: $($masterSharedMailboxes.Count) (with full KRO/Gemeinde data)"
Write-Host "========================================="
Write-Host "`n🚀 Opening output folder..."
Invoke-Item $RunRoot