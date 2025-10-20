<#
.SYNOPSIS
Prüft Wellen-Ordner (Welle01..Welle16 o. ä.) auf GKZ-Konsistenz zwischen
- Wave_*_gemeinden_CT_aktiv.csv (Gemeinden) und
- Wave_*_Seafile_Mapped.csv (Benutzer)

Erzeugt CSV-Reports:
- duplicates_gemeinden.csv (doppelte GKZ in aktiven Gemeinden über alle Wellen)
- mismatched_users.csv (Benutzer mit GKZ, deren definierende Gemeinde in einer anderen Welle liegt)
- orphan_users.csv (Benutzer, deren GKZ NICHT in der Gemeinde-CSV der gleichen Welle vorkommt; oder GKZ leer)
- gemeinden_without_users_in_own_wave.csv (GKZ aus Gemeinden ohne Benutzer in derselben Welle)

.ANNOTATION
- GKZ wird als String behandelt (kein numerischer Check, kein Padding, kein Normalisieren).
- Verglichen wird exakt per String-Gleichheit nach einfachem Trim (kein Case-Topic bei Ziffern).

.HEADERS
- Gemeinden: "Gkz";"Type";"Name";"ValidFrom";"ValidUntil";"DistrictType";"DistrictName";"CtInstance";"Email";"Pfarrer_in"
- Benutzer:  email;gkz;gkz_old;ObjectID

.PARAMETER RootPath
Stammordner, der die Wellenordner enthält.

.PARAMETER WaveFolderPattern
Muster für Wellenordner (Default: 'Welle*').

.PARAMETER GemeindenFilePattern
Muster der Gemeinden-CSV (Default: '*_gemeinden_CT_aktiv.csv').

.PARAMETER UsersFilePattern
Muster der Benutzer-CSV (Default: '*_Seafile_Mapped.csv').

.PARAMETER OutDir
Zielordner für Reports (Default: <RootPath>\_Validation)

.EXAMPLE
.\Check-DGM-Waves.ps1 `
  -RootPath 'C:\Users\durstp\Evangelische Landeskirche in Württemberg\OKR Projekt DGM - Dokumente\DGM - M365\Migrationen'
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$false)]
  [string]$RootPath = 'C:\TTT\Migrationen',

  [Parameter(Mandatory=$false)]
  [string]$WaveFolderPattern = 'Welle*',

  [Parameter(Mandatory=$false)]
  [string]$GemeindenFilePattern = '*Gemeinden_W*',

  [Parameter(Mandatory=$false)]
  [string]$UsersFilePattern = '*_Seafile.csv',

  [Parameter(Mandatory=$false)]
  [string]$OutDir = $(Join-Path -Path $RootPath -ChildPath '_Validation')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-WaveNumberFromName {
  param([string]$Name)
  $m = [regex]::Match($Name, 'Welle\s*0*([0-9]+)', 'IgnoreCase')
  if ($m.Success) { return $m.Groups[1].Value.PadLeft(2,'0') } else { return $null }
}

Write-Host "🔧 Starte Prüfung..." -ForegroundColor Cyan
Write-Host "Root: $RootPath" -ForegroundColor DarkCyan

# 1) Wellenordner ermitteln
$waveDirs = Get-ChildItem -LiteralPath $RootPath -Directory -Filter $WaveFolderPattern | Sort-Object Name
if ($waveDirs.Count -eq 0) {
  throw "Keine Wellenordner unter '$RootPath' gefunden (Muster: '$WaveFolderPattern')."
}

# 2) Pro Welle: Dateien finden
$gemeindenFilesByWave = @{}
$userFilesByWave = @{}
foreach ($dir in $waveDirs) {
  $waveNum = Get-WaveNumberFromName -Name $dir.Name
  if (-not $waveNum) { continue }

  $gFile = Get-ChildItem -LiteralPath $dir.FullName -File -Filter $GemeindenFilePattern -ErrorAction SilentlyContinue | Select-Object -First 1
  $uFile = Get-ChildItem -LiteralPath $dir.FullName -File -Filter $UsersFilePattern      -ErrorAction SilentlyContinue | Select-Object -First 1

  if ($gFile) { $gemeindenFilesByWave[$waveNum] = $gFile.FullName } else {
    Write-Warning "Keine Gemeinden-CSV ($GemeindenFilePattern) in '$($dir.FullName)' gefunden."
  }
  if ($uFile) { $userFilesByWave[$waveNum]      = $uFile.FullName } else {
    Write-Warning "Keine Benutzer-CSV ($UsersFilePattern) in '$($dir.FullName)' gefunden."
  }
}

if ($gemeindenFilesByWave.Count -eq 0) { throw "Keine Dateien '$GemeindenFilePattern' in den Wellenordnern gefunden." }
if ($userFilesByWave.Count -eq 0)      { Write-Warning "Keine Dateien '$UsersFilePattern' in den Wellenordnern gefunden." }

# 3) CSVs laden
# 3.1 Gemeinden: Pro Welle Menge der GKZ-Strings + globale Liste für Duplicates/Mappings
$gemeindenPerWave = @{}      # wave -> HashSet[string] (GKZ)
$allGemeindenRecs = New-Object System.Collections.Generic.List[psobject]

foreach ($kv in $gemeindenFilesByWave.GetEnumerator()) {
  $wave = $kv.Key
  $path = $kv.Value

  $rows = @()
  try {
    $rows = Import-Csv -LiteralPath $path -Delimiter ';' -Encoding UTF8
  } catch {
    throw "Fehler beim Einlesen der Gemeinden-CSV: $path`n$($_.Exception.Message)"
  }

  $set = New-Object System.Collections.Generic.HashSet[string]
  $rowIndex = 0
  foreach ($r in $rows) {
    $rowIndex++

    # GKZ-Feld robust finden
    $gkzValue = $null
    foreach ($cand in @('Gkz','GKZ','gkz')) {
      if ($r.PSObject.Properties.Name -contains $cand) { $gkzValue = [string]$r.$cand; break }
    }
    if ($null -eq $gkzValue) { continue }

    $gkzValue = $gkzValue.Trim()  # nur Leerzeichen entfernen, keine weitere Normalisierung
    if ([string]::IsNullOrWhiteSpace($gkzValue)) { continue }

    [void]$set.Add($gkzValue)

    $allGemeindenRecs.Add([pscustomobject]@{
      GKZ   = $gkzValue
      Wave  = $wave
      File  = $path
      Line  = $rowIndex + 1 # +1 für Header
    })
  }

  $gemeindenPerWave[$wave] = $set
}

# 3.2 Benutzer: Liste (GKZ als String, keine Numerik)
$userRecords = New-Object System.Collections.Generic.List[psobject]
foreach ($kv in $userFilesByWave.GetEnumerator()) {
  $wave = $kv.Key
  $path = $kv.Value

  $rows = @()
  try {
    $rows = Import-Csv -LiteralPath $path -Delimiter ';' -Encoding UTF8
  } catch {
    throw "Fehler beim Einlesen der Benutzer-CSV: $path`n$($_.Exception.Message)"
  }

  $rowIndex = 0
  foreach ($r in $rows) {
    $rowIndex++

    # Email
    $email = ''
    foreach ($cand in @('email','Email')) {
      if ($r.PSObject.Properties.Name -contains $cand) { $email = [string]$r.$cand; break }
    }
    $email = $email.Trim()
    if ($email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { continue }

    # GKZ (String)
    $gkzRaw = ''
    foreach ($cand in @('gkz','GKZ')) {
      if ($r.PSObject.Properties.Name -contains $cand) { $gkzRaw = [string]$r.$cand; break }
    }
    $gkz = $gkzRaw
    if ($null -ne $gkz) { $gkz = $gkz.Trim() } else { $gkz = '' }

    $userRecords.Add([pscustomobject]@{
      Email = $email
      GKZ   = $gkz           # exakter String (getrimmt), keine Normalisierung
      GKZRaw= $gkzRaw        # originaler String (falls später benötigt)
      Wave  = $wave
      File  = $path
      Line  = $rowIndex + 1  # +1 für Header
    })
  }
}

# 4) Regeln prüfen
# 4.1 Doppelte GKZ in Gemeinden (über alle Wellen)
$duplicates = $allGemeindenRecs | Group-Object GKZ | Where-Object { $_.Count -gt 1 } | ForEach-Object {
  [pscustomobject]@{
    GKZ         = $_.Name
    Occurrences = $_.Count
    Waves       = ($_.Group | Select-Object -ExpandProperty Wave | Sort-Object -Unique) -join ','
    Files       = ($_.Group | Select-Object -ExpandProperty File | Sort-Object -Unique) -join ' | '
  }
}
$duplicates = $duplicates | Sort-Object GKZ

# Map GKZ -> erwartete Wave (erste Auftretung in aktiven Gemeinden gewinnt)
$gkzToExpectedWave = @{}
foreach ($g in ($allGemeindenRecs | Sort-Object GKZ, Wave)) {
  if (-not $gkzToExpectedWave.ContainsKey($g.GKZ)) {
    $gkzToExpectedWave[$g.GKZ] = $g.Wave
  }
}

# 4.2 Benutzer mit falscher Wave (Mismatch) – reine String-Gleichheit
$mismatchedUsers = foreach ($u in $userRecords) {
  if (-not [string]::IsNullOrWhiteSpace($u.GKZ) -and $gkzToExpectedWave.ContainsKey($u.GKZ)) {
    $expected = $gkzToExpectedWave[$u.GKZ]
    if ($u.Wave -ne $expected) {
      [pscustomobject]@{
        Email        = $u.Email
        GKZ          = $u.GKZ
        UserWave     = $u.Wave
        ExpectedWave = $expected
        SourceFile   = $u.File
        ExpectedFile = $(if ($userFilesByWave.ContainsKey($expected)) { $userFilesByWave[$expected] } else { "(keine *_Seafile_Mapped.csv in Welle$expected gefunden)" })
        Raw          = "$($u.Email);$($u.GKZRaw)"
      }
    }
  }
}
$mismatchedUsers = $mismatchedUsers | Sort-Object GKZ, Email, UserWave

# 4.3 Orphans (Wellen-spezifisch, nur Same-Wave-Matching maßgeblich)
#  - EmptyGKZ: GKZ leer
#  - GKZNotInSameWaveGemeinden: GKZ nicht in Gemeinden-CSV der gleichen Welle
$orphans = New-Object System.Collections.Generic.List[object]
foreach ($u in $userRecords) {
  $wave = $u.Wave
  $sameWaveSet = $null
  if ($gemeindenPerWave.ContainsKey($wave)) {
    $sameWaveSet = $gemeindenPerWave[$wave]
  }

  if ([string]::IsNullOrWhiteSpace($u.GKZ)) {
    $orphans.Add([pscustomobject]@{
      Email        = $u.Email
      GKZ          = $u.GKZ
      OrphanReason = 'EmptyGKZ'
      Wave         = $u.Wave
      File         = $u.File
    })
    continue
  }

  if ($null -eq $sameWaveSet -or -not $sameWaveSet.Contains($u.GKZ)) {
    $orphans.Add([pscustomobject]@{
      Email        = $u.Email
      GKZ          = $u.GKZ
      OrphanReason = 'GKZNotInSameWaveGemeinden'
      Wave         = $u.Wave
      File         = $u.File
    })
  }
}
$orphans = $orphans | Sort-Object OrphanReason, Wave, GKZ, Email

# 4.4 Gemeinden ohne Benutzer in der eigenen Welle
$gemeindenWithoutUsers = New-Object System.Collections.Generic.List[object]
# Index: Users pro Wave|GKZ (nur nicht-leere GKZ)
$usersIndex = @{}
foreach ($u in ($userRecords | Where-Object { -not [string]::IsNullOrWhiteSpace($_.GKZ) })) {
  $k = "$($u.Wave)|$($u.GKZ)"
  if (-not $usersIndex.ContainsKey($k)) { $usersIndex[$k] = 0 }
  $usersIndex[$k]++
}
foreach ($wave in $gemeindenPerWave.Keys | Sort-Object) {
  $set = $gemeindenPerWave[$wave]
  foreach ($gkz in $set) {
    $k = "$wave|$gkz"
    if (-not $usersIndex.ContainsKey($k)) {
      $file = $gemeindenFilesByWave[$wave]
      $gemeindenWithoutUsers.Add([pscustomobject]@{
        GKZ          = $gkz
        Wave         = $wave
        GemeindeFile = $file
      })
    }
  }
}

# 5) Reports schreiben
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$duplicates               | Export-Csv -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Path (Join-Path $OutDir 'duplicates_gemeinden.csv')
$mismatchedUsers          | Export-Csv -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Path (Join-Path $OutDir 'mismatched_users.csv')
$orphans                  | Export-Csv -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Path (Join-Path $OutDir 'orphan_users.csv')
$gemeindenWithoutUsers    | Export-Csv -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Path (Join-Path $OutDir 'gemeinden_without_users_in_own_wave.csv')

# 6) Zusammenfassung
$orphansEmptyCount    = ($orphans | Where-Object { $_.OrphanReason -eq 'EmptyGKZ' } | Measure-Object).Count
$orphansSameWaveCount = ($orphans | Where-Object { $_.OrphanReason -eq 'GKZNotInSameWaveGemeinden' } | Measure-Object).Count

Write-Host ""
Write-Host "✅ Prüfung abgeschlossen." -ForegroundColor Green
Write-Host ("Aktive Gemeinden-Dateien gefunden: {0}" -f $gemeindenFilesByWave.Count)
Write-Host ("Benutzer-Dateien gefunden: {0}" -f $userFilesByWave.Count)
Write-Host ("Gemeinden (eindeutige GKZ, aktiv): {0}" -f ($allGemeindenRecs | Select-Object -ExpandProperty GKZ -Unique | Measure-Object).Count)
Write-Host ("Benutzerzeilen gesamt: {0}" -f $userRecords.Count)
Write-Host ("Doppelte GKZs in Gemeinden (aktiv): {0}" -f ($duplicates | Measure-Object).Count)
Write-Host ("Benutzer mit Wave-Mismatch: {0}" -f ($mismatchedUsers | Measure-Object).Count)
Write-Host ("Orphans gesamt (same-wave-Check): {0}" -f ($orphans | Measure-Object).Count)
Write-Host (" ...davon EmptyGKZ: {0}" -f $orphansEmptyCount)
Write-Host (" ...davon GKZNotInSameWaveGemeinden: {0}" -f $orphansSameWaveCount)
Write-Host ("Gemeinden ohne Benutzer in eigener Welle: {0}" -f ($gemeindenWithoutUsers | Measure-Object).Count)
Write-Host ("Reports: $OutDir") -ForegroundColor DarkCyan