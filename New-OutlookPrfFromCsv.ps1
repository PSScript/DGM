<#
.SYNOPSIS
  Erzeugt Outlook-PRF-Dateien aus CSV (Type: EXCHANGE | EXO | IMAP) und
  kann sie optional direkt importieren bzw. (IMAP) Credentials im User-Kontext setzen.

.DESCRIPTION
  CSV-Spalten (Beispiel):
    SamAccountName,Email,Type,Server,Password,IsPrimary,IsShared,AccessingUser,ProfileName
  - Type:      EXCHANGE | EXO | IMAP
  - Server:    nur für IMAP (z.B. imap.schulverwaltung.de)
  - Password:  nur für IMAP (optional; Outlook fragt u.U. trotzdem)
  - IsPrimary: TRUE/FALSE (nur informativ; beeinflusst Default-Profilnamen)
  - IsShared:  TRUE/FALSE  (informativ; eigene PRF pro Zeile)
  - AccessingUser: Wer zugreift (informativ/Logik für spätere Erweiterung)
  - ProfileName: optional benutzerdefinierter Anzeigename des Profils

  Es wird pro CSV-Zeile genau EINE PRF erzeugt (ein Profil). Komplexe Multi-Account-PRFs sind möglich,
  aber bewusst NICHT Standard, damit das Boarding einfach/debuggbar bleibt.

.PARAMETER Csv
  Pfad zur Eingabe-CSV.

.PARAMETER OutDir
  Ausgabeverzeichnis für PRFs (wird erstellt, falls nicht vorhanden).

.PARAMETER ImportPrf
  Importiert jede erzeugte PRF sofort für den aktuellen Benutzer:
  outlook.exe /importprf "<prfFile>"

.PARAMETER SetImapCredential
  Legt für IMAP (nur wenn -Password gesetzt war) auch einen Credential-Manager-Eintrag an:
  cmdkey /add:<Server> /user:<Email> /pass:<Password>
  WARNUNG: Nur für Low-Security-Notfälle geeignet!

.PARAMETER DryRun
  Zeigt, was erzeugt/ausgeführt würde – ohne Dateien zu schreiben oder zu importieren.

.PARAMETER Overwrite
  Überschreibt existierende PRF-Dateien.

.EXAMPLE
  .\New-OutlookPrfFromCsv.ps1 -Csv .\users.csv -OutDir .\PRF

.EXAMPLE
  .\New-OutlookPrfFromCsv.ps1 -Csv .\users.csv -OutDir .\PRF -ImportPrf

.EXAMPLE
  .\New-OutlookPrfFromCsv.ps1 -Csv .\users.csv -OutDir .\PRF -SetImapCredential -ImportPrf

.NOTES
  - Passwörter in PRF sind von Outlook nicht offiziell unterstützt (EX/EXO ignoriert Passwörter vollständig).
  - Für IMAP kann ein "Password=" in der PRF stehen; Outlook fragt ggf. trotzdem.
  - Sicherer/besser ist SSO/Modern Auth.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]  [string]$Csv,
  [Parameter(Mandatory=$true)]  [string]$OutDir,
  [switch]$ImportPrf,
  [switch]$SetImapCredential,
  [switch]$DryRun,
  [switch]$Overwrite
)

# ----------------------------- Hilfsfunktionen -----------------------------

function Ensure-Folder {
  param([string]$Path)
  if (-not (Test-Path $Path)) {
    if ($PSBoundParameters.DryRun) { Write-Host "DRYRUN: mkdir $Path" -ForegroundColor Yellow }
    else { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
  }
}

function Sanitize-FileName {
  param([string]$Name)
  $invalid = [IO.Path]::GetInvalidFileNameChars() -join ''
  return ($Name -replace "[$invalid]", '_')
}

function To-Bool {
  param($v)
  if ($null -eq $v) { return $false }
  $s = $v.ToString().Trim().ToUpperInvariant()
  return ($s -in @('1','TRUE','YES','Y','JA'))
}

# ----------------------------- PRF-Templates -------------------------------

# Minimaler EX/EXO-Template (nutzt MAPI/Autodiscover – Server/GUID optional/leer)
$Template_EX = @"
[General]
Custom=1
ProfileName=%PROFILE_NAME%

[Service List]
Service1=Microsoft Exchange Server

[Service1]
OverwriteExistingService=Yes
UniqueService=Yes
MailboxName=%MAILBOX_NAME%
HomeServer=%EXCHANGE_SERVER%
RPCProxyServer=%EXCHANGE_PROXY%
RPCProxyAuthScheme=Negotiate
MailboxGUID=%MAILBOX_GUID%

[Microsoft Exchange Server]
HomeServer=%EXCHANGE_SERVER%
MailboxName=%MAILBOX_NAME%
"@

# Minimaler IMAP-Template – Password-Zeile wird nur eingefügt, wenn vorhanden
$Template_IMAP = @"
[General]
Custom=1
ProfileName=%PROFILE_NAME%

[Service List]
Service1=Microsoft Internet Mail

[Service1]
AccountName=%ACCOUNT_NAME%
Email=%EMAIL%
IMAPServer=%IMAP_SERVER%
IMAPUserName=%IMAP_USER%
%IMAP_PASSWORD%
"@

# ------------------------------ Ablauf -------------------------------------

if (-not (Test-Path $Csv)) { throw "CSV nicht gefunden: $Csv" }
Ensure-Folder -Path $OutDir

$rows = Import-Csv -Path $Csv
if (-not $rows -or $rows.Count -eq 0) { throw "CSV enthält keine Daten: $Csv" }

$created = 0; $imported = 0; $credSet = 0; $skipped = 0
foreach ($r in $rows) {
  $sam   = $r.SamAccountName
  $email = $r.Email
  $type  = ($r.Type    ?? '').ToUpperInvariant().Trim()
  $srv   = ($r.Server  ?? '').Trim()
  $pass  = $r.Password
  $isPrim= To-Bool $r.IsPrimary
  $isShared = To-Bool $r.IsShared
  $access = $r.AccessingUser
  $profileName = if ([string]::IsNullOrWhiteSpace($r.ProfileName)) {
    if ($isPrim) { "$sam" } else { "${sam}_$type" }
  } else { $r.ProfileName }

  if ([string]::IsNullOrWhiteSpace($sam) -or [string]::IsNullOrWhiteSpace($email) -or [string]::IsNullOrWhiteSpace($type)) {
    Write-Warning "Zeile ohne Pflichtfelder (SamAccountName/Email/Type) – übersprungen."
    $skipped++; continue
  }

  $fileName = Sanitize-FileName("$($sam)_$($type).prf")
  $outFile  = Join-Path $OutDir $fileName

  switch ($type) {
    'IMAP' {
      if ([string]::IsNullOrWhiteSpace($srv)) {
        Write-Warning "IMAP-Eintrag ohne Server für $sam – übersprungen."
        $skipped++; continue
      }
      $prf = $Template_IMAP
      $prf = $prf -replace '%PROFILE_NAME%',  [regex]::Escape($profileName).Replace('\\','\\\\')
      $prf = $prf -replace '%ACCOUNT_NAME%',  [regex]::Escape($email).Replace('\\','\\\\')
      $prf = $prf -replace '%EMAIL%',         [regex]::Escape($email).Replace('\\','\\\\')
      $prf = $prf -replace '%IMAP_SERVER%',   [regex]::Escape($srv).Replace('\\','\\\\')
      $prf = $prf -replace '%IMAP_USER%',     [regex]::Escape($email).Replace('\\','\\\\')

      if (-not [string]::IsNullOrWhiteSpace($pass)) {
        # Achtung: Outlook kann Passwortzeilen ignorieren; wir fügen sie nur ein, wenn explizit gewünscht
        $prf = $prf -replace '%IMAP_PASSWORD%', "Password=$pass"
      } else {
        $prf = $prf -replace '%IMAP_PASSWORD%', ""
      }
    }

    'EXCHANGE' { # On-Prem Exchange (vereinfachtes Template, Autodiscover)
      $prf = $Template_EX
      $prf = $prf -replace '%PROFILE_NAME%',   [regex]::Escape($profileName).Replace('\\','\\\\')
      $prf = $prf -replace '%MAILBOX_NAME%',   [regex]::Escape($email).Replace('\\','\\\\')
      # On-Prem: ggf. CAS/MBX-FQDN eintragen – Standard ist leer/Autodiscover
      $prf = $prf -replace '%EXCHANGE_SERVER%', ""
      $prf = $prf -replace '%EXCHANGE_PROXY%', ""
      $prf = $prf -replace '%MAILBOX_GUID%',   ""
    }

    'EXO' {      # Exchange Online (Office 365)
      $prf = $Template_EX
      $prf = $prf -replace '%PROFILE_NAME%',   [regex]::Escape($profileName).Replace('\\','\\\\')
      $prf = $prf -replace '%MAILBOX_NAME%',   [regex]::Escape($email).Replace('\\','\\\\')
      # Für EXO meist ausreichend, beides auf outlook.office365.com zu setzen; Autodiscover regelt den Rest
      $prf = $prf -replace '%EXCHANGE_SERVER%', "outlook.office365.com"
      $prf = $prf -replace '%EXCHANGE_PROXY%',  "outlook.office365.com"
      $prf = $prf -replace '%MAILBOX_GUID%',    ""
    }

    default {
      Write-Warning "Unbekannter Type '$type' bei $sam – übersprungen."
      $skipped++; continue
    }
  }

  if (Test-Path $outFile -and -not $Overwrite) {
    Write-Warning "PRF existiert bereits (nutze -Overwrite zum Überschreiben): $outFile"
    $skipped++; continue
  }

  if ($DryRun) {
    Write-Host "DRYRUN: write PRF -> $outFile (Type=$type, ProfileName='$profileName')" -ForegroundColor Yellow
  } else {
    Set-Content -Path $outFile -Value $prf -Encoding UTF8
  }
  $created++

  # Optional: IMAP Credential im User-Kontext setzen (nur Low-Security/Notfall)
  if ($SetImapCredential -and $type -eq 'IMAP' -and -not [string]::IsNullOrWhiteSpace($pass)) {
    $cmd = "cmdkey /add:$srv /user:$email /pass:$pass"
    if ($DryRun) {
      Write-Host "DRYRUN: $cmd" -ForegroundColor Yellow
    } else {
      Write-Warning "Setze IMAP Credential per cmdkey (nur im User-Kontext sinnvoll; Low-Security!)."
      cmd /c $cmd | Out-Null
    }
    $credSet++
  }

  # Optional: sofort importieren
  if ($ImportPrf) {
    $args = "/importprf `"$outFile`""
    if ($DryRun) {
      Write-Host "DRYRUN: outlook.exe $args" -ForegroundColor Yellow
    } else {
      try {
        Start-Process "outlook.exe" -ArgumentList $args -WindowStyle Hidden
        $imported++
      } catch {
        Write-Warning "Konnte PRF nicht importieren (ist Outlook installiert & im PATH?): $outFile"
      }
    }
  }
}

Write-Host ("Fertig: erstellt={0}, importiert={1}, creds={2}, übersprungen={3}" -f $created,$imported,$credSet,$skipped)
