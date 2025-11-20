#Requires -Version 5.1
<#
.SYNOPSIS
    DGM Master Builder v5.1 - Enhanced Debug Mode
.DESCRIPTION
    Refactored migration orchestrator with:
    - Comprehensive debug logging system
    - Empty variable detection and warnings
    - Output validation with diagnostics
    - Configurable verbosity levels
    - Debug artifact preservation
.NOTES
    Author: Migration Team
    Version: 5.1 - Debug Mode Enhancement
#>

[CmdletBinding()]
param(
    [switch]$DebugMode,
    [ValidateSet('Silent', 'Normal', 'Verbose', 'Diagnostic')]
    [string]$DebugLevel = 'Normal'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Off

#region Debug Infrastructure
$Global:DebugContext = @{
    Enabled = $DebugMode.IsPresent
    Level = $DebugLevel
    Warnings = New-Object System.Collections.ArrayList
    Metrics = @{}
    StartTime = Get-Date
}

function Write-DebugLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Diagnostic')]
        [string]$Level = 'Info',
        
        [string]$Component = 'General',
        
        [hashtable]$Data
    )
    
    if (-not $Global:DebugContext.Enabled -and $Level -ne 'Error') { return }
    
    # Skip diagnostic unless in diagnostic mode
    if ($Level -eq 'Diagnostic' -and $Global:DebugContext.Level -ne 'Diagnostic') { return }
    
    $timestamp = Get-Date -Format 'HH:mm:ss.fff'
    $prefix = switch ($Level) {
        'Info' { '   ' }
        'Warning' { ' ⚠️ ' }
        'Error' { ' ❌' }
        'Success' { ' ✅' }
        'Diagnostic' { ' 🔍' }
    }
    
    $color = switch ($Level) {
        'Info' { 'White' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Success' { 'Green' }
        'Diagnostic' { 'Cyan' }
    }
    
    $logMessage = "[$timestamp]$prefix[$Component] $Message"
    Write-Host $logMessage -ForegroundColor $color
    
    # Track warnings
    if ($Level -eq 'Warning') {
        [void]$Global:DebugContext.Warnings.Add(@{
            Timestamp = $timestamp
            Component = $Component
            Message = $Message
            Data = $Data
        })
    }
}

function Test-EmptyVariable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [AllowEmptyString()]
        $Value,
        
        [Parameter(Mandatory)]
        [string]$Name,
        
        [string]$Component = 'Validation',
        
        [switch]$Critical,
        
        [hashtable]$Context
    )
    
    $isEmpty = $false
    $type = 'Unknown'
    $count = 0
    
    if ($null -eq $Value) {
        $isEmpty = $true
        $type = 'Null'
    }
    elseif ($Value -is [string]) {
        $type = 'String'
        $isEmpty = [string]::IsNullOrWhiteSpace($Value)
    }
    elseif ($Value -is [array] -or $Value -is [System.Collections.ICollection]) {
        $type = 'Collection'
        $count = $Value.Count
        $isEmpty = ($count -eq 0)
    }
    elseif ($Value -is [hashtable]) {
        $type = 'Hashtable'
        $count = $Value.Count
        $isEmpty = ($count -eq 0)
    }
    
    if ($isEmpty) {
        $level = if ($Critical) { 'Error' } else { 'Warning' }
        $data = @{
            Variable = $Name
            Type = $type
            Count = $count
        }
        if ($Context) { $data += $Context }
        
        Write-DebugLog -Message "Variable '$Name' is empty ($type)" -Level $level -Component $Component -Data $data
        
        if ($Critical) {
            throw "Critical variable '$Name' is empty"
        }
    }
    else {
        if ($Global:DebugContext.Level -eq 'Diagnostic') {
            Write-DebugLog -Message "Variable '$Name' validated: $type$(if($count -gt 0){" (Count: $count)"})" `
                -Level 'Diagnostic' -Component $Component
        }
    }
    
    return @{
        IsEmpty = $isEmpty
        Type = $type
        Count = $count
    }
}

function Start-DebugTimer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Operation
    )
    
    if (-not $Global:DebugContext.Enabled) { return $null }
    
    $timer = @{
        Operation = $Operation
        StartTime = Get-Date
    }
    
    Write-DebugLog -Message "Starting: $Operation" -Level 'Diagnostic' -Component 'Timer'
    return $timer
}

function Stop-DebugTimer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Timer,
        
        [int]$ItemCount = 0
    )
    
    if (-not $Timer -or -not $Global:DebugContext.Enabled) { return }
    
    $elapsed = (Get-Date) - $Timer.StartTime
    $operation = $Timer.Operation
    
    if (-not $Global:DebugContext.Metrics.ContainsKey($operation)) {
        $Global:DebugContext.Metrics[$operation] = New-Object System.Collections.ArrayList
    }
    
    [void]$Global:DebugContext.Metrics[$operation].Add(@{
        Duration = $elapsed
        ItemCount = $ItemCount
    })
    
    $msg = "Completed: $operation (Duration: $($elapsed.TotalSeconds.ToString('F2'))s"
    if ($ItemCount -gt 0) {
        $msg += ", Items: $ItemCount, Rate: $([math]::Round($ItemCount / [math]::Max($elapsed.TotalSeconds, 0.001), 2))/s"
    }
    $msg += ")"
    
    Write-DebugLog -Message $msg -Level 'Success' -Component 'Timer'
}

function Save-DebugArtifacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OutputPath
    )
    
    if (-not $Global:DebugContext.Enabled) { return }
    
    $debugDir = Join-Path $OutputPath "_Debug"
    New-Item -ItemType Directory -Path $debugDir -Force | Out-Null
    
    # Export warnings
    if ($Global:DebugContext.Warnings.Count -gt 0) {
        $warningsPath = Join-Path $debugDir "Warnings.json"
        $Global:DebugContext.Warnings | ConvertTo-Json -Depth 5 | Set-Content $warningsPath -Encoding UTF8
        Write-DebugLog -Message "Saved $($Global:DebugContext.Warnings.Count) warnings to $warningsPath" -Component 'Debug'
    }
    
    # Export metrics
    if ($Global:DebugContext.Metrics.Count -gt 0) {
        $metricsPath = Join-Path $debugDir "Metrics.json"
        $Global:DebugContext.Metrics | ConvertTo-Json -Depth 5 | Set-Content $metricsPath -Encoding UTF8
        Write-DebugLog -Message "Saved performance metrics to $metricsPath" -Component 'Debug'
    }
    
    # Summary report
    $summary = @{
        StartTime = $Global:DebugContext.StartTime
        EndTime = Get-Date
        TotalDuration = (Get-Date) - $Global:DebugContext.StartTime
        WarningCount = $Global:DebugContext.Warnings.Count
        MetricsCount = $Global:DebugContext.Metrics.Count
    }
    
    $summaryPath = Join-Path $debugDir "Summary.json"
    $summary | ConvertTo-Json -Depth 3 | Set-Content $summaryPath -Encoding UTF8
}
#endregion

#region Configuration
function Initialize-Configuration {
    [CmdletBinding()]
    param()
    
    $timer = Start-DebugTimer -Operation 'Initialize-Configuration'
    
    $config = @{
        MaxWave = 14
        CurrentBatchThreshold = 10
        SharedMailboxPassword = "tklT75Rc8bxEw4jhLf6A#"
        PilotGKZ = @("420048", "323032", "425019", "444105", "44", "444108", "224026", "407008")
        
        Paths = @{
            SessionRoot = "C:\temp3"
            PasswordMaster = "C:\temp\master_users_all_merged_MASTER_NEU2.csv"
        }
        
        InputFiles = @{
            GemeindenCT = "Gemeinden_Alle_CT.csv"
            Wellen = "wellen.csv"
            Cancom = "2025-11-17_benutzer_status_Cancom.csv"
            SharedMailbox = "2025-11-17_funktionspostfaecher_status_Cancom_semicolon.csv"
            LizenzReport = "Lizenzreport_ALL1.csv"
            Churchtools = "churchtools_organizations.csv"
            ElkwAdmins = "elkwadmins-2025-09-15-krz.csv"
            ElkwStats = "elkwstats-2025-09-15-krz.csv"
        }
        
        Delimiters = @{
            Default = ';'
            Cancom = ','
            SharedMailbox = ';'
        }
        
        OutputDirs = @{
            RunStamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
        }
    }
    
    # Validate configuration
    Test-EmptyVariable -Value $config.PilotGKZ -Name 'PilotGKZ' -Component 'Config' -Critical
    Test-EmptyVariable -Value $config.Paths.SessionRoot -Name 'SessionRoot' -Component 'Config' -Critical
    
    Stop-DebugTimer -Timer $timer
    return $config
}
#endregion

#region CSV Helpers
function Import-CsvSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        
        [string]$Delimiter = ';',
        
        [switch]$Required
    )
    
    $timer = Start-DebugTimer -Operation "Import-CsvSafe: $(Split-Path $Path -Leaf)"
    
    if (-not (Test-Path $Path)) {
        if ($Required) {
            Write-DebugLog -Message "Required file missing: $Path" -Level 'Error' -Component 'CSV'
            throw "Required file missing: $Path"
        }
        Write-DebugLog -Message "Optional file not found: $Path" -Level 'Warning' -Component 'CSV'
        return @()
    }
    
    try {
        $data = Import-Csv -Path $Path -Delimiter $Delimiter -Encoding UTF8
        
        # Validate imported data
        $validation = Test-EmptyVariable -Value $data -Name (Split-Path $Path -Leaf) -Component 'CSV'
        
        if ($validation.IsEmpty -and $Required) {
            throw "Required file is empty: $Path"
        }
        
        Write-DebugLog -Message "Loaded $($data.Count) rows from $(Split-Path $Path -Leaf)" `
            -Level 'Success' -Component 'CSV'
        
        Stop-DebugTimer -Timer $timer -ItemCount $data.Count
        return $data
    }
    catch {
        Write-DebugLog -Message "Failed to import CSV from ${Path}: $_" -Level 'Error' -Component 'CSV'
        throw
    }
}

function Export-CsvClean {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Data,
        
        [Parameter(Mandatory)]
        [string]$Path,
        
        [string]$Delimiter = ';',
        
        [string[]]$Select
    )
    
    $timer = Start-DebugTimer -Operation "Export-CsvClean: $(Split-Path $Path -Leaf)"
    
    # Validate input data
    $validation = Test-EmptyVariable -Value $Data -Name "ExportData:$(Split-Path $Path -Leaf)" -Component 'CSV'
    
    if ($validation.IsEmpty) {
        Write-DebugLog -Message "Attempting to export empty dataset to: $Path" -Level 'Warning' -Component 'CSV'
    }
    
    if ($Select) {
        $Data = $Data | Select-Object $Select
    }
    
    $Data | Export-Csv -Path $Path -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation
    Strip-Quotes $Path
    
    Write-DebugLog -Message "Exported $($Data.Count) rows to $(Split-Path $Path -Leaf)" `
        -Level 'Success' -Component 'CSV'
    
    Stop-DebugTimer -Timer $timer -ItemCount $Data.Count
}

function Strip-Quotes {
    [CmdletBinding()]
    param([string]$Path)
    
    if (Test-Path $Path) {
        (Get-Content $Path) -replace '"', '' | Set-Content $Path -Encoding UTF8
    }
}
#endregion

#region Core Helpers
function Get-PropValue {
    [CmdletBinding()]
    param($obj, [string]$prop)
    
    if (-not $obj) {
        if ($Global:DebugContext.Level -eq 'Diagnostic') {
            Write-DebugLog -Message "Get-PropValue: Object is null for property '$prop'" `
                -Level 'Diagnostic' -Component 'Property'
        }
        return ""
    }
    
    if ($obj.PSObject.Properties.Match($prop).Count -gt 0) {
        $val = $obj.$prop
        if ($null -ne $val) {
            return "$val".Trim()
        }
        else {
            if ($Global:DebugContext.Level -eq 'Diagnostic') {
                Write-DebugLog -Message "Property '$prop' exists but is null" `
                    -Level 'Diagnostic' -Component 'Property'
            }
        }
    }
    else {
        if ($Global:DebugContext.Level -eq 'Diagnostic') {
            Write-DebugLog -Message "Property '$prop' not found on object" `
                -Level 'Diagnostic' -Component 'Property'
        }
    }
    
    return ""
}

function Normalize-Key {
    [CmdletBinding()]
    param([string]$s)
    
    $s = ($s -as [string])
    if (-not $s) { return "" }
    return ($s.Trim().ToLower()) -replace '\s+', ' '
}

function CanonKRO {
    [CmdletBinding()]
    param($val, $pad = 3)
    
    $s = ($val -as [string]).Trim()
    if ([string]::IsNullOrEmpty($s) -or $s -eq 'NULL') {
        return $null
    }
    return $s.PadLeft($pad, '0')
}

function Parse-DateFlexible {
    [CmdletBinding()]
    param([string]$s)
    
    if (-not $s) { return $null }
    $t = $s.Trim()
    if ($t -eq '' -or $t -eq 'NULL') { return $null }
    
    $de = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    $styles = [System.Globalization.DateTimeStyles]::None
    $fmts = @('dd.MM.yyyy', 'd.M.yyyy', 'dd.MM.yy', 'd.M.yy', 'yyyy-MM-dd', 'yyyy/MM/dd', 'yyyy.MM.dd', 'yyyy')
    
    foreach ($f in $fmts) {
        try { return [datetime]::ParseExact($t, $f, $de, $styles) }
        catch { }
    }
    
    try { return [datetime]::Parse($t, $de, $styles) }
    catch { }
    try { return [datetime]::Parse($t, [System.Globalization.CultureInfo]::InvariantCulture, $styles) }
    catch { }
    
    return $null
}

function SplitEmails {
    [CmdletBinding()]
    param([string[]]$arr)
    
    return ($arr -join ',') -split '[,\s]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

function GetCTInstanzFromUrl {
    [CmdletBinding()]
    param([string]$url)
    
    $u = ($url -as [string]).Trim()
    if ([string]::IsNullOrEmpty($u)) { return "" }
    
    if ($u -match 'https://(elkw\d{4})\.krz\.tools') { return $matches[1] }
    elseif ($u -match 'elkw(\d{4})') { return "elkw$($matches[1])" }
    
    return ""
}
#endregion

#region Lookup Builders
function Build-LookupHash {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Data,
        
        [Parameter(Mandatory)]
        [string]$KeyProperty,
        
        [scriptblock]$KeyTransform
    )
    
    $timer = Start-DebugTimer -Operation "Build-LookupHash: $KeyProperty"
    
    # Validate input
    Test-EmptyVariable -Value $Data -Name "LookupData:$KeyProperty" -Component 'Lookup' -Critical
    
    $hash = @{}
    $emptyKeyCount = 0
    $duplicateCount = 0
    
    foreach ($item in $Data) {
        $key = Get-PropValue $item $KeyProperty
        if ($KeyTransform) { $key = & $KeyTransform $key }
        
        if ([string]::IsNullOrWhiteSpace($key)) {
            $emptyKeyCount++
            continue
        }
        
        if ($hash.ContainsKey($key)) {
            $duplicateCount++
            if ($Global:DebugContext.Level -eq 'Diagnostic') {
                Write-DebugLog -Message "Duplicate key '$key' in lookup for property '$KeyProperty'" `
                    -Level 'Diagnostic' -Component 'Lookup'
            }
        }
        
        $hash[$key] = $item
    }
    
    if ($emptyKeyCount -gt 0) {
        Write-DebugLog -Message "Skipped $emptyKeyCount items with empty keys in lookup '$KeyProperty'" `
            -Level 'Warning' -Component 'Lookup'
    }
    
    if ($duplicateCount -gt 0) {
        Write-DebugLog -Message "Found $duplicateCount duplicate keys in lookup '$KeyProperty' (last value wins)" `
            -Level 'Warning' -Component 'Lookup'
    }
    
    Test-EmptyVariable -Value $hash -Name "LookupHash:$KeyProperty" -Component 'Lookup'
    
    Write-DebugLog -Message "Built lookup hash with $($hash.Count) entries from property '$KeyProperty'" `
        -Level 'Success' -Component 'Lookup'
    
    Stop-DebugTimer -Timer $timer -ItemCount $hash.Count
    return $hash
}

function Build-MultiValueLookup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Data,
        
        [Parameter(Mandatory)]
        [string]$KeyProperty,
        
        [scriptblock]$KeyTransform
    )
    
    $timer = Start-DebugTimer -Operation "Build-MultiValueLookup: $KeyProperty"
    
    Test-EmptyVariable -Value $Data -Name "MultiValueData:$KeyProperty" -Component 'Lookup' -Critical
    
    $hash = @{}
    $emptyKeyCount = 0
    
    foreach ($item in $Data) {
        $key = Get-PropValue $item $KeyProperty
        if ($KeyTransform) { $key = & $KeyTransform $key }
        
        if ([string]::IsNullOrWhiteSpace($key)) {
            $emptyKeyCount++
            continue
        }
        
        if (-not $hash.ContainsKey($key)) {
            $hash[$key] = New-Object System.Collections.ArrayList
        }
        [void]$hash[$key].Add($item)
    }
    
    if ($emptyKeyCount -gt 0) {
        Write-DebugLog -Message "Skipped $emptyKeyCount items with empty keys in multi-value lookup '$KeyProperty'" `
            -Level 'Warning' -Component 'Lookup'
    }
    
    Test-EmptyVariable -Value $hash -Name "MultiValueLookup:$KeyProperty" -Component 'Lookup'
    
    Write-DebugLog -Message "Built multi-value lookup with $($hash.Count) keys" `
        -Level 'Success' -Component 'Lookup'
    
    Stop-DebugTimer -Timer $timer -ItemCount $hash.Count
    return $hash
}

function Build-PilotHash {
    [CmdletBinding()]
    param([string[]]$PilotGKZ)
    
    $timer = Start-DebugTimer -Operation 'Build-PilotHash'
    
    Test-EmptyVariable -Value $PilotGKZ -Name 'PilotGKZ' -Component 'Lookup' -Critical
    
    $hash = @{}
    foreach ($gkz in $PilotGKZ) {
        $canon = CanonKRO $gkz 3
        if ($canon) { $hash[$canon] = $true }
    }
    
    Test-EmptyVariable -Value $hash -Name 'PilotHash' -Component 'Lookup'
    
    Write-DebugLog -Message "Built Pilot hash with $($hash.Count) entries" `
        -Level 'Success' -Component 'Lookup'
    
    Stop-DebugTimer -Timer $timer -ItemCount $hash.Count
    return $hash
}
#endregion

#region Data Import
function Import-AllData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    
    $timer = Start-DebugTimer -Operation 'Import-AllData'
    
    Write-Host "`n========================================="
    Write-Host "📥 Loading Input Files"
    Write-Host "========================================="
    
    $sessionRoot = $Config.Paths.SessionRoot
    $inputFiles = $Config.InputFiles
    $delimiters = $Config.Delimiters
    
    Write-DebugLog -Message "Loading input files from $sessionRoot" -Component 'Import'
    
    $data = @{
        GemeindenCT = Import-CsvSafe -Path (Join-Path $sessionRoot $inputFiles.GemeindenCT) -Delimiter $delimiters.Default -Required
        Wellen = Import-CsvSafe -Path (Join-Path $sessionRoot $inputFiles.Wellen) -Delimiter $delimiters.Default -Required
        Cancom = Import-CsvSafe -Path (Join-Path $sessionRoot $inputFiles.Cancom) -Delimiter $delimiters.Cancom -Required
        LizenzReport = Import-CsvSafe -Path (Join-Path $sessionRoot $inputFiles.LizenzReport) -Delimiter $delimiters.Default -Required
        Churchtools = Import-CsvSafe -Path (Join-Path $sessionRoot $inputFiles.Churchtools) -Delimiter $delimiters.Default -Required
        ElkwAdmins = Import-CsvSafe -Path (Join-Path $sessionRoot $inputFiles.ElkwAdmins) -Delimiter $delimiters.Default -Required
        ElkwStats = Import-CsvSafe -Path (Join-Path $sessionRoot $inputFiles.ElkwStats) -Delimiter $delimiters.Default -Required
        SharedMailbox = Import-CsvSafe -Path (Join-Path $sessionRoot $inputFiles.SharedMailbox) -Delimiter $delimiters.SharedMailbox
        Passwords = Import-CsvSafe -Path $Config.Paths.PasswordMaster -Delimiter $delimiters.Default
    }
    
    # Validate critical data loaded
    $critical = @('GemeindenCT', 'Wellen', 'Cancom', 'LizenzReport')
    foreach ($key in $critical) {
        Test-EmptyVariable -Value $data[$key] -Name $key -Component 'Import' -Critical
    }
    
    Write-DebugLog -Message "All input files loaded successfully" -Level 'Success' -Component 'Import'
    Stop-DebugTimer -Timer $timer
    
    return $data
}
#endregion

#region Lookup Building
function Build-AllLookups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Data,
        
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    
    $timer = Start-DebugTimer -Operation 'Build-AllLookups'
    
    Write-Host "`n========================================="
    Write-Host "🗂️  Building Lookup Tables"
    Write-Host "========================================="
    
    $lookups = @{
        Gemeinden = Build-LookupHash -Data $Data.GemeindenCT -KeyProperty 'KRO' -KeyTransform { param($k) $k.Trim() }
        Welle_Vorher = Build-LookupHash -Data $Data.Wellen -KeyProperty 'Welle_vorher' -KeyTransform { param($k) Normalize-Key $k }
        Welle_Nachher = Build-LookupHash -Data $Data.Wellen -KeyProperty 'Welle_nachher' -KeyTransform { param($k) Normalize-Key $k }
        Welle_Dekanat = Build-LookupHash -Data $Data.Wellen -KeyProperty 'Welle_Dekanat' -KeyTransform { param($k) Normalize-Key $k }
        Welle_Kirchenbezirk = Build-LookupHash -Data $Data.Wellen -KeyProperty 'Welle_Kirchenbezirk' -KeyTransform { param($k) Normalize-Key $k }
        License = Build-LookupHash -Data $Data.LizenzReport -KeyProperty 'UserPrincipalName' -KeyTransform { param($k) $k.ToLower().Trim() }
        ChurchTools = @{}
        CTAdmins = @{}
        CTStats = Build-LookupHash -Data $Data.ElkwStats -KeyProperty 'Instanz' -KeyTransform { param($k) $k.Trim() }
        Passwords = @{}
        Pilot = Build-PilotHash -PilotGKZ $Config.PilotGKZ
        KROToGemeinde = @{}
        UsersByKRO = @{}
    }
    
    # Build ChurchTools hash
    Write-DebugLog -Message "Building ChurchTools lookup..." -Component 'Lookup'
    $ctCount = 0
    foreach ($ct in $Data.Churchtools) {
        $id = ($ct.identifier -as [string]).Trim()
        $url = ($ct.url -as [string]).Trim()
        if ($id) {
            $lookups.ChurchTools[$id] = @{
                CT_Instanz = GetCTInstanzFromUrl $url
                designation = ($ct.designation -as [string]).Trim()
                url = $url
            }
            $ctCount++
        }
    }
    Write-DebugLog -Message "ChurchTools lookup complete: $ctCount entries" -Level 'Success' -Component 'Lookup'
    Test-EmptyVariable -Value $lookups.ChurchTools -Name 'ChurchTools' -Component 'Lookup'
    
    # Build CT Admins by Instanz
    Write-DebugLog -Message "Building CT Admins lookup..." -Component 'Lookup'
    $adminCount = 0
    foreach ($row in $Data.ElkwAdmins) {
        $inst = ($row.Instanz -as [string]).Trim()
        if (-not $inst) { continue }
        
        if (-not $lookups.CTAdmins.ContainsKey($inst)) {
            $lookups.CTAdmins[$inst] = @()
        }
        
        $mail = ($row.'E-Mail' -as [string]).Trim()
        if ($mail -and $lookups.CTAdmins[$inst] -notcontains $mail) {
            $lookups.CTAdmins[$inst] += $mail
            $adminCount++
        }
    }
    Write-DebugLog -Message "CT Admins lookup complete: $adminCount admins across $($lookups.CTAdmins.Count) instances" `
        -Level 'Success' -Component 'Lookup'
    Test-EmptyVariable -Value $lookups.CTAdmins -Name 'CTAdmins' -Component 'Lookup'
    
    # Build Password lookup
    Write-DebugLog -Message "Building Password lookup..." -Component 'Lookup'
    $pwCount = 0
    foreach ($row in $Data.Passwords) {
        $email = ($row.email -as [string]).ToLower().Trim()
        $password = ($row.password -as [string]).Trim()
        if ($email -and $password) {
            $lookups.Passwords[$email] = $password
            $pwCount++
        }
    }
    Write-DebugLog -Message "Password lookup complete: $pwCount passwords" -Level 'Success' -Component 'Lookup'
    Test-EmptyVariable -Value $lookups.Passwords -Name 'Passwords' -Component 'Lookup'
    
    # Build KRO to Gemeinde multi-map
    Write-DebugLog -Message "Building KRO to Gemeinde multi-map..." -Component 'Lookup'
    $kroCount = 0
    foreach ($g in $Data.GemeindenCT) {
        $allKeys = @()
        if ($g.KRO) { $allKeys += (CanonKRO $g.KRO 3) }
        if ($g.PSObject.Properties.Match('kro_alt').Count -gt 0 -and $g.kro_alt) {
            $allKeys += ($g.kro_alt -split ',' | ForEach-Object { CanonKRO $_ 3 } | Where-Object { $_ })
        }
        $allKeys = $allKeys | Where-Object { $_ } | Sort-Object -Unique
        
        foreach ($k in $allKeys) {
            if (-not $lookups.KROToGemeinde.ContainsKey($k)) {
                $lookups.KROToGemeinde[$k] = @()
            }
            $lookups.KROToGemeinde[$k] += $g
            $kroCount++
        }
    }
    Write-DebugLog -Message "KRO multi-map complete: $kroCount mappings" -Level 'Success' -Component 'Lookup'
    Test-EmptyVariable -Value $lookups.KROToGemeinde -Name 'KROToGemeinde' -Component 'Lookup'
    
    Write-Host "`n📊 Lookup Summary:"
    Write-Host "  - Gemeinden: $($lookups.Gemeinden.Count)"
    Write-Host "  - Licenses: $($lookups.License.Count)"
    Write-Host "  - Passwords: $($lookups.Passwords.Count)"
    Write-Host "  - Pilot GKZ: $($lookups.Pilot.Count)"
    Write-Host "  - ChurchTools: $($lookups.ChurchTools.Count)"
    Write-Host "  - CT Admins: $($lookups.CTAdmins.Count) instances"
    
    Stop-DebugTimer -Timer $timer
    return $lookups
}
#endregion

#region Wave Assignment
function Get-WaveInfo {
    [CmdletBinding()]
    param(
        [string]$DekanatName,
        [string]$FallbackKey,
        [hashtable]$Welle_Dekanat_Hash,
        [hashtable]$Welle_Vorher_Hash
    )
    
    $Wave_via_KRO_Dekanat = ""
    $keyDek = Normalize-Key $DekanatName
    if ($keyDek -and $Welle_Dekanat_Hash.ContainsKey($keyDek)) {
        $Wave_via_KRO_Dekanat = $Welle_Dekanat_Hash[$keyDek].Wave
    }
    
    $Wave_via_Fallback = ""
    $fallback_norm_key = Normalize-Key $FallbackKey
    if ($fallback_norm_key -and $Welle_Vorher_Hash.ContainsKey($fallback_norm_key)) {
        $Wave_via_Fallback = $Welle_Vorher_Hash[$fallback_norm_key].Wave
    }
    
    $Wave = ""
    $Wave_Source = ""
    $Wave_Warning = ""
    
    if ($Wave_via_KRO_Dekanat -and $Wave_via_Fallback) {
        if ($Wave_via_KRO_Dekanat -eq $Wave_via_Fallback) {
            $Wave = $Wave_via_KRO_Dekanat
            $Wave_Source = "Both(KRO->Dekanat & Fallback)"
        }
        else {
            $Wave = $Wave_via_KRO_Dekanat
            $Wave_Source = "KRO->Dekanat"
            $Wave_Warning = "Mismatch: Dekanat=$Wave_via_KRO_Dekanat vs Fallback=$Wave_via_Fallback"
        }
    }
    elseif ($Wave_via_KRO_Dekanat) {
        $Wave = $Wave_via_KRO_Dekanat
        $Wave_Source = "KRO->Dekanat"
    }
    elseif ($Wave_via_Fallback) {
        $Wave = $Wave_via_Fallback
        $Wave_Source = "Fallback"
    }
    
    # Warn if wave couldn't be determined
    if (-not $Wave) {
        if ($Global:DebugContext.Level -eq 'Diagnostic') {
            Write-DebugLog -Message "Could not determine wave for Dekanat='$DekanatName', Fallback='$FallbackKey'" `
                -Level 'Diagnostic' -Component 'Wave'
        }
    }
    
    return [PSCustomObject]@{
        Wave = $Wave
        Source = $Wave_Source
        Warning = $Wave_Warning
    }
}

function Test-PilotOverride {
    [CmdletBinding()]
    param(
        [string]$KRO,
        [hashtable]$PilotHash
    )
    
    $canon = CanonKRO $KRO 3
    return ($canon -and $PilotHash.ContainsKey($canon))
}
#endregion

#region User Enrichment
function Add-WaveAssignments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Users,
        
        [Parameter(Mandatory)]
        [hashtable]$Lookups
    )
    
    $timer = Start-DebugTimer -Operation 'Add-WaveAssignments'
    
    Write-Host "`n========================================="
    Write-Host "🌊 Adding Wave Assignments"
    Write-Host "========================================="
    
    Test-EmptyVariable -Value $Users -Name 'Users' -Component 'Enrichment' -Critical
    
    $noWaveCount = 0
    $pilotCount = 0
    
    foreach ($u in $Users) {
        $kro_old = Get-PropValue $u 'kirchliche_stelle_gkz'
        
        # Get Dekanat from KRO
        $DekanatName = ""
        if ($kro_old -and $Lookups.Gemeinden.ContainsKey($kro_old)) {
            $g = $Lookups.Gemeinden[$kro_old]
            $DekanatName = (Get-PropValue $g 'ORENAMEEBENE4')
        }
        
        # Get wave via multiple paths
        $waveInfo = Get-WaveInfo -DekanatName $DekanatName -FallbackKey $kro_old `
            -Welle_Dekanat_Hash $Lookups.Welle_Dekanat -Welle_Vorher_Hash $Lookups.Welle_Vorher
        
        # Pilot override
        if (Test-PilotOverride -KRO $kro_old -PilotHash $Lookups.Pilot) {
            $waveInfo.Wave = "Pilot"
            $waveInfo.Source = "Pilot_Override"
            $waveInfo.Warning = ""
            $pilotCount++
        }
        
        if (-not $waveInfo.Wave) {
            $noWaveCount++
        }
        
        $u | Add-Member -NotePropertyName Wave -NotePropertyValue $waveInfo.Wave -Force
        $u | Add-Member -NotePropertyName Wave_Source -NotePropertyValue $waveInfo.Source -Force
        $u | Add-Member -NotePropertyName Wave_Warning -NotePropertyValue $waveInfo.Warning -Force
    }
    
    if ($noWaveCount -gt 0) {
        Write-DebugLog -Message "$noWaveCount users could not be assigned to a wave" `
            -Level 'Warning' -Component 'Enrichment'
    }
    
    Write-DebugLog -Message "Wave assignments complete (Pilot: $pilotCount, No Wave: $noWaveCount)" `
        -Level 'Success' -Component 'Enrichment'
    
    Stop-DebugTimer -Timer $timer -ItemCount $Users.Count
    return $Users
}

function Add-PasswordData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Users,
        
        [Parameter(Mandatory)]
        [hashtable]$PasswordLookup
    )
    
    $timer = Start-DebugTimer -Operation 'Add-PasswordData'
    
    Write-Host "`n========================================="
    Write-Host "🔐 Adding Password Data"
    Write-Host "========================================="
    
    Test-EmptyVariable -Value $Users -Name 'Users' -Component 'Enrichment' -Critical
    Test-EmptyVariable -Value $PasswordLookup -Name 'PasswordLookup' -Component 'Enrichment'
    
    $foundCount = 0
    $missingCount = 0
    
    foreach ($u in $Users) {
        $email = (Get-PropValue $u 'email').ToLower().Trim()
        $password = ""
        
        if ($email -and $PasswordLookup.ContainsKey($email)) {
            $password = $PasswordLookup[$email]
            $foundCount++
        }
        else {
            if ($email) {
                $missingCount++
            }
        }
        
        $u | Add-Member -NotePropertyName Password -NotePropertyValue $password -Force
    }
    
    if ($missingCount -gt 0) {
        Write-DebugLog -Message "$missingCount users missing passwords" `
            -Level 'Warning' -Component 'Enrichment'
    }
    
    Write-DebugLog -Message "Passwords added: $foundCount / $($Users.Count)" `
        -Level 'Success' -Component 'Enrichment'
    
    Stop-DebugTimer -Timer $timer -ItemCount $foundCount
    return $Users
}

function Add-CloudState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Users,
        
        [Parameter(Mandatory)]
        [hashtable]$LicenseLookup
    )
    
    $timer = Start-DebugTimer -Operation 'Add-CloudState'
    
    Write-Host "`n========================================="
    Write-Host "☁️  Adding Cloud State Data"
    Write-Host "========================================="
    
    Test-EmptyVariable -Value $Users -Name 'Users' -Component 'Enrichment' -Critical
    Test-EmptyVariable -Value $LicenseLookup -Name 'LicenseLookup' -Component 'Enrichment'
    
    $foundCount = 0
    
    foreach ($u in $Users) {
        $email = (Get-PropValue $u 'email').ToLower().Trim()
        $licFound = $null
        
        if ($email) {
            $licFound = $LicenseLookup[$email]
            if ($licFound) { $foundCount++ }
        }
        
        # Extract cloud properties
        $u | Add-Member -NotePropertyName Cloud_AD_Domain -NotePropertyValue (Get-PropValue $licFound 'AD_Domain') -Force
        $u | Add-Member -NotePropertyName Cloud_LizenzName -NotePropertyValue (Get-PropValue $licFound 'LizenzName_Graph') -Force
        $u | Add-Member -NotePropertyName Cloud_OU -NotePropertyValue (Get-PropValue $licFound 'OU_AD') -Force
        $u | Add-Member -NotePropertyName Cloud_Company -NotePropertyValue (Get-PropValue $licFound 'Company_AD') -Force
        $u | Add-Member -NotePropertyName Cloud_ObjectID -NotePropertyValue (Get-PropValue $licFound 'Az_ObjectID') -Force
        $u | Add-Member -NotePropertyName Cloud_Displayname -NotePropertyValue (Get-PropValue $licFound 'Name') -Force
        $u | Add-Member -NotePropertyName Cloud_AP_Typ -NotePropertyValue (Get-PropValue $licFound 'AP_Typ') -Force
        $u | Add-Member -NotePropertyName Cloud_Mailbox_Typ -NotePropertyValue (Get-PropValue $licFound 'Mailbox_Typ') -Force
        $u | Add-Member -NotePropertyName Cloud_Routing -NotePropertyValue (Get-PropValue $licFound 'Routing') -Force
    }
    
    Write-DebugLog -Message "Cloud state data added: $foundCount users matched with licenses" `
        -Level 'Success' -Component 'Enrichment'
    
    Stop-DebugTimer -Timer $timer -ItemCount $foundCount
    return $Users
}

function Get-ProvisioningStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $User,
        
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    
    $migStatus = Get-PropValue $User 'postfachmigration'
    $cloudAPTyp = Get-PropValue $User 'Cloud_AP_Typ'
    $password = Get-PropValue $User 'Password'
    
    if ($migStatus -notmatch '^(YES|OK|JA)$') {
        return "IGNORE"
    }
    
    if (-not [string]::IsNullOrWhiteSpace($cloudAPTyp) -and $cloudAPTyp -match 'DGM') {
        if ([string]::IsNullOrWhiteSpace($password)) {
            return "MISSING_PW"
        }
        else {
            return "OK"
        }
    }
    else {
        return "UNPROVISIONED"
    }
}

function Test-RoutingConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $User,
        
        [Parameter(Mandatory)]
        [int]$CurrentBatchThreshold
    )
    
    $wave = Get-PropValue $User 'Wave'
    $cloudRouting = Get-PropValue $User 'Cloud_Routing'
    
    $wRaw = $wave -replace 'W|Pilot', ''
    
    $ExpectedRouting = ""
    $RoutingStatus = ""
    
    if ($wRaw -match '^\d+$') {
        $waveNum = [int]$wRaw
        
        if ($waveNum -lt $CurrentBatchThreshold) {
            $ExpectedRouting = "NOT RouteToKopano"
            if ($cloudRouting -eq "RouteToKopano") {
                $RoutingStatus = "ROUTING_FAILURE"
            }
            else {
                $RoutingStatus = "ROUTING_OK"
            }
        }
        elseif ($waveNum -ge $CurrentBatchThreshold) {
            $ExpectedRouting = "RouteToKopano"
            if ($cloudRouting -ne "RouteToKopano") {
                $RoutingStatus = "ROUTING_FAILURE"
            }
            else {
                $RoutingStatus = "ROUTING_OK"
            }
        }
    }
    elseif ($wave -eq "Pilot") {
        $ExpectedRouting = "NOT RouteToKopano"
        if ($cloudRouting -eq "RouteToKopano") {
            $RoutingStatus = "ROUTING_FAILURE"
        }
        else {
            $RoutingStatus = "ROUTING_OK"
        }
    }
    
    return [PSCustomObject]@{
        ExpectedRouting = $ExpectedRouting
        RoutingStatus = $RoutingStatus
    }
}

function Add-ProvisioningAndRoutingStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Users,
        
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    
    $timer = Start-DebugTimer -Operation 'Add-ProvisioningAndRoutingStatus'
    
    Write-Host "`n========================================="
    Write-Host "🔄 Analyzing Provisioning & Routing"
    Write-Host "========================================="
    
    Test-EmptyVariable -Value $Users -Name 'Users' -Component 'Analysis' -Critical
    
    $statusCounts = @{
        OK = 0
        UNPROVISIONED = 0
        MISSING_PW = 0
        IGNORE = 0
        ROUTING_OK = 0
        ROUTING_FAILURE = 0
    }
    
    foreach ($u in $Users) {
        $provStatus = Get-ProvisioningStatus -User $u -Config $Config
        $routingInfo = Test-RoutingConfiguration -User $u -CurrentBatchThreshold $Config.CurrentBatchThreshold
        
        $u | Add-Member -NotePropertyName ProvisioningStatus -NotePropertyValue $provStatus -Force
        $u | Add-Member -NotePropertyName RoutingStatus -NotePropertyValue $routingInfo.RoutingStatus -Force
        $u | Add-Member -NotePropertyName Routing_Expected -NotePropertyValue $routingInfo.ExpectedRouting -Force
        
        # Track counts
        if ($statusCounts.ContainsKey($provStatus)) {
            $statusCounts[$provStatus]++
        }
        if ($statusCounts.ContainsKey($routingInfo.RoutingStatus)) {
            $statusCounts[$routingInfo.RoutingStatus]++
        }
    }
    
    Write-Host "`n📊 Status Distribution:"
    Write-Host "  Provisioning:"
    Write-Host "    - OK: $($statusCounts.OK)"
    Write-Host "    - UNPROVISIONED: $($statusCounts.UNPROVISIONED)"
    Write-Host "    - MISSING_PW: $($statusCounts.MISSING_PW)"
    Write-Host "    - IGNORE: $($statusCounts.IGNORE)"
    Write-Host "  Routing:"
    Write-Host "    - OK: $($statusCounts.ROUTING_OK)"
    Write-Host "    - FAILURES: $($statusCounts.ROUTING_FAILURE)"
    
    if ($statusCounts.ROUTING_FAILURE -gt 0) {
        Write-DebugLog -Message "$($statusCounts.ROUTING_FAILURE) routing failures detected" `
            -Level 'Warning' -Component 'Analysis'
    }
    
    Write-DebugLog -Message "Status analysis complete" -Level 'Success' -Component 'Analysis'
    Stop-DebugTimer -Timer $timer -ItemCount $Users.Count
    
    return $Users
}
#endregion

#region Main Processing
function Process-Users {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Data,
        
        [Parameter(Mandatory)]
        [hashtable]$Lookups,
        
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    
    $timer = Start-DebugTimer -Operation 'Process-Users'
    
    Write-Host "`n========================================="
    Write-Host "👥 Processing Users"
    Write-Host "========================================="
    
    $users = $Data.Cancom
    Test-EmptyVariable -Value $users -Name 'Cancom Users' -Component 'Processing' -Critical
    
    $users = Add-WaveAssignments -Users $users -Lookups $Lookups
    $users = Add-PasswordData -Users $users -PasswordLookup $Lookups.Passwords
    $users = Add-CloudState -Users $users -LicenseLookup $Lookups.License
    $users = Add-ProvisioningAndRoutingStatus -Users $users -Config $Config
    
    Write-DebugLog -Message "User processing complete: $($users.Count) users" `
        -Level 'Success' -Component 'Processing'
    
    Stop-DebugTimer -Timer $timer -ItemCount $users.Count
    return $users
}

function Process-SharedMailboxes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Data,
        
        [Parameter(Mandatory)]
        [hashtable]$Lookups,
        
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    
    $timer = Start-DebugTimer -Operation 'Process-SharedMailboxes'
    
    Write-Host "`n========================================="
    Write-Host "📫 Processing Shared Mailboxes"
    Write-Host "========================================="
    
    if ($Data.SharedMailbox.Count -eq 0) {
        Write-DebugLog -Message "No shared mailboxes to process" -Level 'Info' -Component 'Processing'
        return @()
    }
    
    $mailboxes = New-Object System.Collections.Generic.List[psobject]
    $noWaveCount = 0
    
    foreach ($smb in $Data.SharedMailbox) {
        $email = (Get-PropValue $smb 'funktionspostfach').ToLower().Trim()
        if (-not $email) {
            Write-DebugLog -Message "Shared mailbox missing email address" `
                -Level 'Warning' -Component 'Processing'
            continue
        }
        
        $gkz_raw = Get-PropValue $smb 'kirchliche_stelle_gkz'
        $gkz = CanonKRO $gkz_raw 3
        
        $DekanatName = ""
        if ($gkz -and $Lookups.Gemeinden.ContainsKey($gkz)) {
            $g = $Lookups.Gemeinden[$gkz]
            $DekanatName = (Get-PropValue $g 'ORENAMEEBENE4')
        }
        
        $waveInfo = Get-WaveInfo -DekanatName $DekanatName -FallbackKey $gkz_raw `
            -Welle_Dekanat_Hash $Lookups.Welle_Dekanat -Welle_Vorher_Hash $Lookups.Welle_Vorher
        
        # Pilot override
        if (Test-PilotOverride -KRO $gkz -PilotHash $Lookups.Pilot) {
            $waveInfo.Wave = "Pilot"
            $waveInfo.Source = "Pilot_Override"
        }
        
        if (-not $waveInfo.Wave) {
            $noWaveCount++
        }
        
        $row = [PSCustomObject]@{
            funktionspostfach = $email
            Password = $Config.SharedMailboxPassword
            gkz = $gkz
            Wave = $waveInfo.Wave
            Wave_Source = $waveInfo.Source
            Wave_Warning = $waveInfo.Warning
            ticket_status = Get-PropValue $smb 'ticket_status'
            besitzermailplus = Get-PropValue $smb 'besitzermailplus'
        }
        
        $mailboxes.Add($row) | Out-Null
    }
    
    if ($noWaveCount -gt 0) {
        Write-DebugLog -Message "$noWaveCount shared mailboxes could not be assigned to a wave" `
            -Level 'Warning' -Component 'Processing'
    }
    
    Write-DebugLog -Message "Shared mailbox processing complete: $($mailboxes.Count) mailboxes" `
        -Level 'Success' -Component 'Processing'
    
    Stop-DebugTimer -Timer $timer -ItemCount $mailboxes.Count
    return $mailboxes
}
#endregion

#region Export Functions
function Export-MasterFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Users,
        
        $SharedMailboxes,
        
        [Parameter(Mandatory)]
        [string]$OutDir_Master,
        
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    
    $timer = Start-DebugTimer -Operation 'Export-MasterFiles'
    
    Write-Host "`n========================================="
    Write-Host "💾 Exporting Master Files"
    Write-Host "========================================="
    
    Test-EmptyVariable -Value $Users -Name 'Users' -Component 'Export' -Critical
    
    # Export master user file
    $masterUserPath = Join-Path $OutDir_Master "Master_User_Migration.csv"
    Export-CsvClean -Data $Users -Path $masterUserPath -Delimiter $Config.Delimiters.Default
    Write-DebugLog -Message "Master users exported" -Level 'Success' -Component 'Export'
    
    # Export shared mailboxes if any
    if ($SharedMailboxes -and $SharedMailboxes.Count -gt 0) {
        $masterSharedPath = Join-Path $OutDir_Master "Master_SharedMailboxes_Migration.csv"
        Export-CsvClean -Data $SharedMailboxes -Path $masterSharedPath -Delimiter $Config.Delimiters.Default
        Write-DebugLog -Message "Master shared mailboxes exported" -Level 'Success' -Component 'Export'
    }
    else {
        Write-DebugLog -Message "No shared mailboxes to export" -Level 'Info' -Component 'Export'
    }
    
    # Export provisioning status reports
    Export-ProvisioningReports -Users $Users -OutDir $OutDir_Master -Config $Config
    
    # Export routing failures
    Export-RoutingFailures -Users $Users -OutDir $OutDir_Master -Config $Config
    
    Stop-DebugTimer -Timer $timer
}

function Export-ProvisioningReports {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Users,
        
        [Parameter(Mandatory)]
        [string]$OutDir,
        
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    
    $timer = Start-DebugTimer -Operation 'Export-ProvisioningReports'
    
    Write-Host "`n📋 Exporting provisioning reports..."
    
    $provisionCols = @(
        'Wave', 'email', 'Password', 'Vorname', 'Nachname', 'Displayname', 'mailnickname',
        'samaccountname', 'Gemeinde', 'Dekanat', 'Kirchenbezirk', 'ObjectID', 'AP_Typ',
        'AD_Domain', 'MailboxMigration', 'gkz', 'gkz_old'
    )
    
    # Unprovisioned users
    $unprovisioned = @($Users | Where-Object { $_.ProvisioningStatus -eq 'UNPROVISIONED' })
    
    if ($unprovisioned.Count -eq 0) {
        Write-DebugLog -Message "No unprovisioned users to export" -Level 'Info' -Component 'Export'
    }
    else {
        # Nachzügler (past waves)
        $nachzuegler = @($unprovisioned | Where-Object {
            $wRaw = $_.Wave -replace 'W|Pilot', ''
            if ($wRaw -match '^\d+$') { [int]$wRaw -lt $Config.CurrentBatchThreshold }
            else { $false }
        })
        
        if ($nachzuegler.Count -gt 0) {
            $nFile = Join-Path $OutDir "Unprovisioned_Users_AD_Create_Nachzuegler.csv"
            Export-CsvClean -Data $nachzuegler -Path $nFile -Delimiter $Config.Delimiters.Default -Select $provisionCols
            Write-Host "  - Nachzügler: $($nachzuegler.Count)"
        }
        
        # Current and future waves
        foreach ($i in $Config.CurrentBatchThreshold..$Config.MaxWave) {
            $w = "W$i"
            $wUsers = @($unprovisioned | Where-Object { $_.Wave -eq $w })
            if ($wUsers.Count -gt 0) {
                $wFile = Join-Path $OutDir "Unprovisioned_Users_AD_Create_$w.csv"
                Export-CsvClean -Data $wUsers -Path $wFile -Delimiter $Config.Delimiters.Default -Select $provisionCols
                Write-Host "  - $w`: $($wUsers.Count)"
            }
        }
        
        # Pilot
        $pilot = @($unprovisioned | Where-Object { $_.Wave -eq 'Pilot' })
        if ($pilot.Count -gt 0) {
            $pFile = Join-Path $OutDir "Unprovisioned_Users_AD_Create_Pilot.csv"
            Export-CsvClean -Data $pilot -Path $pFile -Delimiter $Config.Delimiters.Default -Select $provisionCols
            Write-Host "  - Pilot: $($pilot.Count)"
        }
    }
    
    # Migration pending (ready for mailbox migration)
    $migrationPending = @($Users | Where-Object {
        $wRaw = $_.Wave -replace 'W|Pilot', ''
        $isPast = if ($wRaw -match '^\d+$') { [int]$wRaw -lt $Config.CurrentBatchThreshold } else { $false }
        
        $isPast -and
        ($_.MailboxMigration -match '^(YES|OK|JA)$') -and
        (-not [string]::IsNullOrWhiteSpace($_.Cloud_AP_Typ) -and $_.Cloud_AP_Typ -match 'DGM') -and
        (-not [string]::IsNullOrWhiteSpace($_.Password))
    })
    
    if ($migrationPending.Count -gt 0) {
        $mpFile = Join-Path $OutDir "Migration_Pending_W1-9.csv"
        Export-CsvClean -Data $migrationPending -Path $mpFile -Delimiter $Config.Delimiters.Default -Select $provisionCols
        
        # Quick2 format
        $qFile = Join-Path $OutDir "MigrationUsers_Quick2.csv"
        $quick2 = $migrationPending | Select-Object @{N = 'EmailAddress'; E = { $_.email } }
        Export-CsvClean -Data $quick2 -Path $qFile -Delimiter ','
        
        Write-Host "  - Migration Pending: $($migrationPending.Count)"
    }
    else {
        Write-DebugLog -Message "No migration-pending users to export" -Level 'Info' -Component 'Export'
    }
    
    Stop-DebugTimer -Timer $timer
}

function Export-RoutingFailures {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Users,
        
        [Parameter(Mandatory)]
        [string]$OutDir,
        
        [Parameter(Mandatory)]
        [hashtable]$Config
    )
    
    $timer = Start-DebugTimer -Operation 'Export-RoutingFailures'
    
    $routingFailures = @($Users | Where-Object { $_.RoutingStatus -eq 'ROUTING_FAILURE' })
    
    if ($routingFailures.Count -gt 0) {
        $rfFile = Join-Path $OutDir "Routing_Failures.csv"
        $cols = @('Wave', 'email', 'Cloud_Routing', 'Routing_Expected', 'ProvisioningStatus', 'MailboxMigration', 'Cloud_AP_Typ')
        Export-CsvClean -Data $routingFailures -Path $rfFile -Delimiter $Config.Delimiters.Default -Select $cols
        
        Write-Host "  ⚠️  Routing failures: $($routingFailures.Count)" -ForegroundColor Yellow
        Write-DebugLog -Message "$($routingFailures.Count) routing failures exported" `
            -Level 'Warning' -Component 'Export'
    }
    else {
        Write-Host "  ✅ All routing configurations correct" -ForegroundColor Green
        Write-DebugLog -Message "No routing failures detected" -Level 'Success' -Component 'Export'
    }
    
    Stop-DebugTimer -Timer $timer
}
#endregion

#region Summary
function Show-MigrationSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Users,
        
        $SharedMailboxes,
        
        [Parameter(Mandatory)]
        [hashtable]$Config,
        
        [Parameter(Mandatory)]
        [string]$RunRoot
    )
    
    Write-Host "`n========================================="
    Write-Host "✅ DGM Master Builder v5.1 - COMPLETE"
    Write-Host "========================================="
    Write-Host "`nOutput directory: $RunRoot"
    
    # Provisioning status
    Write-Host "`n📊 Provisioning Status:"
    $okCount = @($Users | Where-Object { $_.ProvisioningStatus -eq 'OK' }).Count
    $unprovCount = @($Users | Where-Object { $_.ProvisioningStatus -eq 'UNPROVISIONED' }).Count
    $missingPwCount = @($Users | Where-Object { $_.ProvisioningStatus -eq 'MISSING_PW' }).Count
    $ignoreCount = @($Users | Where-Object { $_.ProvisioningStatus -eq 'IGNORE' }).Count
    
    Write-Host "  ✓ OK (ready): $okCount"
    Write-Host "  ⚠ UNPROVISIONED (need AD): $unprovCount"
    Write-Host "  ⚠ MISSING_PW (need password): $missingPwCount"
    Write-Host "  - IGNORE (not migrating): $ignoreCount"
    
    # Routing status
    Write-Host "`n📊 Routing Status:"
    $routingOkCount = @($Users | Where-Object { $_.RoutingStatus -eq 'ROUTING_OK' }).Count
    $routingFailCount = @($Users | Where-Object { $_.RoutingStatus -eq 'ROUTING_FAILURE' }).Count
    
    Write-Host "  ✓ Correct routing: $routingOkCount"
    if ($routingFailCount -gt 0) {
        Write-Host "  ⚠️  ROUTING FAILURES: $routingFailCount" -ForegroundColor Red
    }
    else {
        Write-Host "  ✓ No routing failures" -ForegroundColor Green
    }
    
    # Wave distribution
    Write-Host "`n📊 Wave Distribution:"
    $pilotCount = @($Users | Where-Object { $_.Wave -eq 'Pilot' }).Count
    $noWaveCount = @($Users | Where-Object { -not $_.Wave }).Count
    
    Write-Host "  - Pilot: $pilotCount"
    
    for ($i = 1; $i -le $Config.MaxWave; $i++) {
        $cnt = @($Users | Where-Object { $_.Wave -eq "W$i" }).Count
        if ($cnt -gt 0) {
            $marker = if ($i -lt $Config.CurrentBatchThreshold) { "✓" }
            elseif ($i -eq $Config.CurrentBatchThreshold) { "→" }
            else { " " }
            Write-Host "  $marker W$($i.ToString('D2')): $cnt users"
        }
    }
    
    if ($noWaveCount -gt 0) {
        Write-Host "  ⚠️  No Wave: $noWaveCount" -ForegroundColor Yellow
    }
    
    if ($SharedMailboxes) {
        Write-Host "`n📊 Shared Mailboxes: $($SharedMailboxes.Count)"
    }
    
    # Debug summary
    if ($Global:DebugContext.Enabled) {
        Write-Host "`n========================================="
        Write-Host "🔍 Debug Summary"
        Write-Host "========================================="
        
        $totalDuration = (Get-Date) - $Global:DebugContext.StartTime
        Write-Host "Total Runtime: $($totalDuration.TotalSeconds.ToString('F2'))s"
        Write-Host "Warnings: $($Global:DebugContext.Warnings.Count)"
        
        if ($Global:DebugContext.Warnings.Count -gt 0) {
            Write-Host "`nTop Warnings:"
            $Global:DebugContext.Warnings | 
                Group-Object -Property Component | 
                Sort-Object Count -Descending | 
                Select-Object -First 5 | 
                ForEach-Object {
                    Write-Host "  - $($_.Name): $($_.Count)"
                }
        }
        
        Write-Host "`nPerformance Metrics:"
        $Global:DebugContext.Metrics.GetEnumerator() | 
            ForEach-Object {
                $op = $_.Key
                $metrics = $_.Value
                $avgDuration = ($metrics | Measure-Object -Property Duration -Average).Average.TotalSeconds
                Write-Host "  - $op`: $($avgDuration.ToString('F3'))s avg"
            }
    }
    
    Write-Host "`n========================================="
    Write-Host "🚀 Opening output folder..."
    Invoke-Item $RunRoot
}
#endregion

#region Main Execution
try {
    Write-Host "========================================="
    Write-Host "DGM Master Builder v5.1"
    if ($Global:DebugContext.Enabled) {
        Write-Host "Debug Mode: ENABLED (Level: $($Global:DebugContext.Level))" -ForegroundColor Cyan
    }
    Write-Host "========================================="
    
    $Config = Initialize-Configuration
    
    # Setup output directories
    $RunRoot = Join-Path $Config.Paths.SessionRoot ("out_" + $Config.OutputDirs.RunStamp)
    $OutDir_Master = Join-Path $RunRoot "Master"
    $OutDir_Waves = Join-Path $RunRoot "Waves"
    New-Item -ItemType Directory -Path $RunRoot, $OutDir_Master, $OutDir_Waves -Force | Out-Null
    
    # Load all data
    $Data = Import-AllData -Config $Config
    
    # Build lookups
    $Lookups = Build-AllLookups -Data $Data -Config $Config
    
    # Process users and shared mailboxes
    $Users = Process-Users -Data $Data -Lookups $Lookups -Config $Config
    $SharedMailboxes = Process-SharedMailboxes -Data $Data -Lookups $Lookups -Config $Config
    
    # Export all files
    Export-MasterFiles -Users $Users -SharedMailboxes $SharedMailboxes -OutDir_Master $OutDir_Master -Config $Config
    
    # TODO: Export-WaveFiles (similar refactor pattern)
    # TODO: Export-ExcelWorkbooks (conditional if ImportExcel available)
    
    # Save debug artifacts
    if ($Global:DebugContext.Enabled) {
        Save-DebugArtifacts -OutputPath $RunRoot
    }
    
    # Show summary
    Show-MigrationSummary -Users $Users -SharedMailboxes $SharedMailboxes -Config $Config -RunRoot $RunRoot
    
    Write-Host "`n✅ Migration processing completed successfully" -ForegroundColor Green
}
catch {
    Write-Host "`n❌ Migration processing failed" -ForegroundColor Red
    Write-Error "Error: $_"
    Write-Error $_.ScriptStackTrace
    
    if ($Global:DebugContext.Enabled -and $RunRoot) {
        Save-DebugArtifacts -OutputPath $RunRoot
        Write-Host "`n🔍 Debug artifacts saved to: $(Join-Path $RunRoot '_Debug')" -ForegroundColor Cyan
    }
    
    exit 1
}
#endregion