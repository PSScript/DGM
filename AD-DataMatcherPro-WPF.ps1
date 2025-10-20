#Requires -Version 5.1
#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    AD-DataMatcherPro-WPF.ps1 — The definitive tool for joining live Active Directory data with any CSV.

.DESCRIPTION
    - Direct Active Directory Export: Pulls live user data, bypassing the need for an intermediate file.
    - Default Column Selection: Automatically pre-selects key AD attributes for export.
    - Join Type Control: A checkbox lets you switch between a left join (all AD users) and an inner join (only matched users).
    - Selectable Delimiter: Export to CSV using a semicolon, comma, or tab via a dropdown menu.
    - Crisp, per-monitor DPI scaling, robust case-insensitive join, and live data previews.

.AUTHOR
    Ada, Jan & Gemini — A collaborative evolution.
#>

# ---[ 1. Assembly Loading & Configuration ]---
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$script:AppRoot = Join-Path $env:APPDATA 'AdaTools\AD-DataMatcherPro-WPF'
$script:ConfigPath = Join-Path $script:AppRoot 'config.json'
if (-not (Test-Path $script:AppRoot)) { New-Item -ItemType Directory -Path $script:AppRoot | Out-Null }

# ---[ 2. Helper Functions (Including AD Logic) ]---

#region GUI & CSV Helper Functions
function Load-Config {
    if (Test-Path $script:ConfigPath) {
        try { return (Get-Content $script:ConfigPath -Raw | ConvertFrom-Json) }
        catch { Write-Warning "Could not parse config.json. Using defaults." }
    }
    return [pscustomobject]@{ RecentData = @(); LastExportDir = 'C:\temp'; LastDelimiter = ';'; IncludeOnlyMatched = $false }
}

function Save-Config($cfg) {
    try { $cfg | ConvertTo-Json -Depth 6 | Out-File -FilePath $script:ConfigPath -Encoding UTF8 }
    catch { Write-Warning "Failed to save config.json: $($_.Exception.Message)" }
}

function Detect-Csv($Path) {
    if (-not (Test-Path $Path)) { throw "File not found: $Path" }
    $delimiters = ',', ';', "`t"
    foreach ($d in $delimiters) {
        try {
            $rows = Import-Csv -Path $Path -Delimiter $d
            if ($rows -and $rows[0].PSObject.Properties.Name.Count -gt 1) {
                return @{ Rows = $rows; Headers = @($rows[0].PSObject.Properties.Name) }
            }
        }
        catch {}
    }
    throw "Could not parse the file as CSV with any known delimiter."
}

function Update-RecentList([string[]]$current, [string]$newPath, [int]$max = 20) {
    $list = @($newPath) + @($current | Where-Object { $_ -and ($_ -ne $newPath) })
    return $list | Select-Object -First $max
}

function Set-DefaultExportPath($cfg) {
    $dir = $cfg.LastExportDir
    if (-not (Test-Path $dir)) { try { New-Item -ItemType Directory -Path $dir -ErrorAction Stop | Out-Null } catch { $dir = $env:TEMP } }
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    return Join-Path $dir "Export_$timestamp.csv"
}

function Show-MessageBox($text, $title = 'AD Data Matcher Pro', $icon = 'Information') {
    [System.Windows.MessageBox]::Show($text, $title, 'OK', $icon) | Out-Null
}

function Normalize-Key { param($s) if ($null -eq $s) { '' } else { $s.ToString().Trim() } }
#endregion

#region Active Directory Helper Functions
function Format-ProperCase {
    param([string]$InputString)
    if ([string]::IsNullOrWhiteSpace($InputString)) { return $InputString }
    return "$($InputString.Substring(0, 1).ToUpper())$($InputString.Substring(1).ToLower())"
}

function Get-FormattedOU {
    param([string]$DistinguishedName)
    if (-not $DistinguishedName) { return $null }
    return ($DistinguishedName -split ',' | Where-Object { $_ -match 'OU=' }) -replace 'OU=' -join '/'
}

function Get-TopOU {
    param([string]$DistinguishedName)
    if (-not $DistinguishedName) { return $null }
    $ous = ($DistinguishedName -split "," | Where-Object { $_ -match "OU="}) -replace "OU="
    if ($DistinguishedName -match "365") { return $ous[1..0] -join "/" } else { return $ous[0] }
}

function Convert-RecipientTypeToMailboxType {
    param([string]$RecipientTypeDetails)
    switch ($RecipientTypeDetails) {
        "1" { return "Onprem_UserMailbox" }; "4" { return "Onprem_SharedMailbox" }; "128" { return "Onprem_Mailuser" }
        "2147483648" { return "Cloud_UserMailbox" }; "34359738368" { return "Cloud_SharedMailbox" }; default { return "Other" }
    }
}

function Parse-EmailName {
    param([string]$Mail)
    if ($Mail -match '^([^.]+)\.([^@]+)@') {
        $first = Format-ProperCase -InputString $matches[1]
        $last = Format-ProperCase -InputString $matches[2]
        return [PSCustomObject]@{ First = $first; Last = $last; DisplayLastFirst = "$last, $first"; DisplayFirstLast = "$first $last" }
    }
    return [PSCustomObject]@{ First = ""; Last = ""; DisplayLastFirst = ""; DisplayFirstLast = "" }
}
#endregion

# ---[ 3. Core Logic Functions ]---

function Invoke-ADUserExport {
    param(
        [string]$SearchBase = "OU=DGM,DC=OKR,DC=ELK-WUE,DC=DE" # You can change this default
    )
    $statusText.Text = "Querying AD from Search Base: $SearchBase..."
    $window.UpdateLayout() # Force UI refresh

    $adUsers = Get-ADUser -SearchBase $SearchBase -Filter * -Properties DistinguishedName, userprincipalname, mail, mailnickname, proxyaddresses, TargetAddress, msExchRemoteRecipientType, msExchRecipientDisplayType, msExchRecipientTypeDetails, extensionattribute4, enabled, lastlogondate, pwdLastSet, DisplayName, GivenName, Surname, physicalDeliveryOfficeName
    if (-not $adUsers) { throw "No users found in the specified OU." }
    
    $statusText.Text = "Processing $($adUsers.Count) user records from Active Directory..."
    $window.UpdateLayout()

    $results = $adUsers | ForEach-Object -Process {
        $emailParts = Parse-EmailName -Mail $_.mail
        [PSCustomObject]@{
            samaccountname = $_.samaccountname; userprincipalname = $_.userprincipalname; mail = $_.mail;
            OU = Get-FormattedOU -DistinguishedName $_.DistinguishedName; TOP_OU = Get-TopOU -DistinguishedName $_.DistinguishedName;
            mailnickname = $_.mailnickname; msExchRemoteRecipientType = $_.msExchRemoteRecipientType;
            msExchRecipientDisplayType = $_.msExchRecipientDisplayType; msExchRecipientTypeDetails = $_.msExchRecipientTypeDetails;
            extensionattribute4 = $_.extensionattribute4; enabled = $_.enabled; lastlogondate = $_.lastlogondate;
            PrimarySMTPAddress = ($_.proxyaddresses | Where-Object { $_ -cmatch "^SMTP" }) -replace "SMTP:"; Targetaddress = $_.Targetaddress -replace "smtp:";
            proxyaddresses = ($_.proxyaddresses | Where-Object { $_ -cmatch "^smtp"}) -replace "smtp:" -join '|';
            Mailbox_Typ = Convert-RecipientTypeToMailboxType -RecipientTypeDetails $_.msExchRecipientTypeDetails;
            recipienttype = if ($_.msExchRemoteRecipientType -in 1..4) { 'UserMailbox' } else { 'SharedMailbox' };
            Serverlocation = if ($_.msExchRecipientTypeDetails -in "1", "4") { "Onprem" } elseif ($_.msExchRecipientTypeDetails -in "2147483648", "34359738368") { "Cloud" } else { "N/A" };
            pwdLastSet = if ($_.pwdLastSet -ne $null -and $_.pwdLastSet -ne 0) { [datetime]::FromFileTime($_.pwdLastSet) } else { $null };
            DisplayName = $_.DisplayName; Vorname = $_.GivenName; Nachname = $_.Surname; Displayname_email = "$($_.GivenName) $($_.Surname)"; Office = $_.physicalDeliveryOfficeName;
            First2 = $emailParts.First; Last2 = $emailParts.Last; Displayname2 = $emailParts.DisplayLastFirst; Displayname_mail2 = $emailParts.DisplayFirstLast
        }
    }
    return $results
}

function Do-JoinAndExport {
    param(
        [Parameter(Mandatory)]$UsersRows, [Parameter(Mandatory)]$DataRows, [Parameter(Mandatory)]$UsersKey, [Parameter(Mandatory)]$DataKey,
        [string[]]$UsersColsSelected, [string[]]$DataColsSelected, [Parameter(Mandatory)]$OutPath, $StatusTextBlock,
        [Parameter(Mandatory)]$Delimiter, [Parameter(Mandatory)]$IncludeOnlyMatched
    )
    $StatusTextBlock.Text = "Building case-insensitive index from Data file..."
    $index = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
    $dupCount = 0
    foreach ($r in $DataRows) {
        $key = Normalize-Key $r.$DataKey
        if (-not [string]::IsNullOrEmpty($key) -and -not $index.ContainsKey($key)) { $index[$key] = $r }
        elseif ($index.ContainsKey($key)) { $dupCount++ }
    }

    $StatusTextBlock.Text = "Processing $($UsersRows.Count) user records..."
    $results = [System.Collections.Generic.List[object]]::new()
    $matched, $unmatched = 0, 0
    foreach ($userRow in $UsersRows) {
        $userKeyValue = Normalize-Key $userRow.$UsersKey
        $dataRow = if ($index.ContainsKey($userKeyValue)) { $index[$userKeyValue] } else { $null }

        if ($dataRow) { $matched++ } else { $unmatched++ }
        if ($IncludeOnlyMatched -and -not $dataRow) { continue } # Skip unmatched rows if in "inner join" mode

        $outputRow = [ordered]@{}
        foreach ($col in $UsersColsSelected) { $outputRow[$col] = $userRow.$col }
        foreach ($col in $DataColsSelected) { $outputRow["Data.$col"] = if ($dataRow) { $dataRow.$col } else { $null } }
        
        $results.Add([pscustomobject]$outputRow)
    }

    $StatusTextBlock.Text = "Exporting $($results.Count) records... ($matched matched, $unmatched unmatched in source, $dupCount duplicate data keys ignored)"
    $results | Export-Csv -Path $OutPath -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter
}


# ---[ 4. XAML GUI Definition ]---
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AD Data Matcher Pro (WPF)" Height="800" Width="1200" MinHeight="650" MinWidth="950">
    <Grid>
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>

        <Border Grid.Row="0" Padding="10" Background="#F0F0F0" BorderBrush="#CCCCCC" BorderThickness="0,0,0,1">
            <StackPanel Orientation="Horizontal">
                <Label Content="Users Source:" VerticalAlignment="Center"/>
                <TextBox x:Name="txtUsersPath" Width="375" Margin="5,0" IsReadOnly="True" ToolTip="Path to the Users source. Loaded from AD or a CSV file."/>
                <Button x:Name="btnExportAD" Content="Export from AD..." Width="120" Margin="5,0" Background="#E0E8FF" FontWeight="Bold" ToolTip="Load live user data directly from Active Directory."/>
                <Button x:Name="btnBrowseUsers" Content="Browse CSV..." Width="100" Margin="5,0" ToolTip="Load user data from a CSV file."/>
            </StackPanel>
        </Border>

        <Grid Grid.Row="1" Margin="10">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*" MinWidth="400"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*" MinWidth="400"/></Grid.ColumnDefinitions>
            <DockPanel Grid.Column="0" Margin="5">
                <StackPanel DockPanel.Dock="Top">
                    <Label Content="1. Select Join Key:"/><ComboBox x:Name="cmbUsersKey"/>
                    <Label Content="2. Select Columns to Export (Ctrl/Shift + Click):" Margin="0,10,0,0"/>
                    <DataGrid x:Name="dgvUsersSelect" Height="250" AutoGenerateColumns="True" SelectionMode="Extended" SelectionUnit="FullRow" IsReadOnly="True"/>
                </StackPanel>
                <Label DockPanel.Dock="Top" Content="Source Preview (First 100 Rows):" Margin="0,10,0,0"/>
                <DataGrid x:Name="gridUsers" DockPanel.Dock="Bottom" AutoGenerateColumns="True" IsReadOnly="True"/>
            </DockPanel>
            <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Center" VerticalAlignment="Stretch" Background="LightGray"/>
            <DockPanel Grid.Column="2" Margin="5">
                <StackPanel DockPanel.Dock="Top">
                    <StackPanel Orientation="Horizontal">
                        <Label Content="Data File (Right):" Target="{Binding ElementName=cmbData}" VerticalAlignment="Center"/>
                        <ComboBox x:Name="cmbData" Width="300" Margin="5,0" IsEditable="True"/>
                        <Button x:Name="btnBrowseData" Content="Browse..." Width="90" Margin="5,0"/>
                    </StackPanel>
                    <Label Content="1. Select Join Key:" Margin="0,10,0,0"/><ComboBox x:Name="cmbDataKey"/>
                    <Button x:Name="btnPreviewMatched" Content="Preview Matched Data" Margin="0,5,0,0" ToolTip="Update column preview with examples from matching rows only."/>
                    <Label Content="2. Select Columns to Export:" Margin="0,10,0,0"/>
                    <DataGrid x:Name="dgvDataSelect" Height="250" AutoGenerateColumns="True" SelectionMode="Extended" SelectionUnit="FullRow" IsReadOnly="True"/>
                </StackPanel>
                <Label DockPanel.Dock="Top" Content="File Preview (First 100 Rows):" Margin="0,10,0,0"/>
                <DataGrid x:Name="gridData" DockPanel.Dock="Bottom" AutoGenerateColumns="True" IsReadOnly="True"/>
            </DockPanel>
        </Grid>

        <Border Grid.Row="2" Padding="10" Background="#F0F0F0" BorderBrush="#CCCCCC" BorderThickness="0,1,0,0">
            <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <Label Content="Export Path:"/><TextBox x:Name="txtOut" Width="400" Margin="5,0" VerticalContentAlignment="Center"/>
                    <Button x:Name="btnChange" Content="Change..." Width="90" Margin="5,0"/>
                    <Label Content="Delimiter:" Margin="20,0,0,0"/><ComboBox x:Name="cmbDelimiter" Width="120" Margin="5,0"/>
                    <CheckBox x:Name="chkIncludeOnlyMatched" Content="Include Only Matched Users" VerticalAlignment="Center" Margin="20,0,0,0" ToolTip="If checked, only rows with a successful match will be exported."/>
                </StackPanel>
                <Button x:Name="btnExport" Grid.Column="1" Content="Export to CSV" Width="200" Height="40" Background="#D0E8D0" FontWeight="Bold"/>
            </Grid>
        </Border>

        <StatusBar Grid.Row="3"><StatusBarItem><TextBlock x:Name="statusText" Text="Ready. Load data from AD or a Users file to begin."/></StatusBarItem></StatusBar>
    </Grid>
</Window>
'@

# ---[ 5. UI Initialization and Logic ]---
$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
try { $window = [Windows.Markup.XamlReader]::Load($reader) } catch { Write-Error "Error parsing XAML: $($_.Exception.Message)"; return }

"txtUsersPath","btnBrowseUsers","btnExportAD","cmbUsersKey","dgvUsersSelect","gridUsers","cmbData","btnBrowseData",
"cmbDataKey","btnPreviewMatched","dgvDataSelect","gridData","txtOut","btnChange","cmbDelimiter","chkIncludeOnlyMatched","btnExport","statusText" | ForEach-Object {
    Set-Variable -Name $_ -Value $window.FindName($_) -Scope Script
}

$script:cfg = Load-Config
$UsersState = @{ Rows = $null; Headers = @() }; $DataState = @{ Rows = $null; Headers = @() }
$txtOut.Text = Set-DefaultExportPath $script:cfg

# Initialize new controls
$cmbDelimiter.ItemsSource = @("Semicolon (;)", "Comma (,)", "Tab")
$cmbDelimiter.SelectedItem = switch ($script:cfg.LastDelimiter) { ';' {"Semicolon (;)"} ',' {"Comma (,)"} default {"Semicolon (;)"} }
$chkIncludeOnlyMatched.IsChecked = $script:cfg.IncludeOnlyMatched

# ---[ 6. UI Functions ]---

function Populate-TransposedSelect($dgvSelect, $rows, $headers, [string[]]$DefaultSelection = $null) {
    $exampleRows = $rows | Select-Object -First 5
    $transposed = @()
    foreach ($prop in $headers) {
        $rowObj = [ordered]@{ Property = $prop }
        for ($i = 0; $i -lt $exampleRows.Count; $i++) { $rowObj["Example$($i+1)"] = $exampleRows[$i].$prop }
        $transposed += [pscustomobject]$rowObj
    }
    $dgvSelect.ItemsSource = $transposed
    
    if ($DefaultSelection) {
        $dgvSelect.SelectedItems.Clear()
        $itemsToSelect = $dgvSelect.ItemsSource | Where-Object { $_.Property -in $DefaultSelection }
        foreach ($item in $itemsToSelect) { $dgvSelect.SelectedItems.Add($item) }
    } else {
        $dgvSelect.SelectAll()
    }
}

function Populate-FileUI {
    param($headers, $comboKey, $dgvSelect, $gridPreview, $rows, [string[]]$DefaultSelection = $null)
    $comboKey.ItemsSource = $headers
    if ($headers -contains "samaccountname") { $comboKey.SelectedItem = "samaccountname" } elseif ($headers.Count -gt 0) { $comboKey.SelectedIndex = 0 }
    Populate-TransposedSelect -dgvSelect $dgvSelect -rows $rows -headers $headers -DefaultSelection $DefaultSelection
    $gridPreview.ItemsSource = $rows | Select-Object -First 100
}

# ---[ 7. Event Handlers ]---
$window.Add_Closing({
    $script:cfg.LastExportDir = [System.IO.Path]::GetDirectoryName($txtOut.Text)
    $script:cfg.LastDelimiter = ($cmbDelimiter.SelectedItem -split ' ')[1].Trim('()')
    $script:cfg.IncludeOnlyMatched = $chkIncludeOnlyMatched.IsChecked
    Save-Config $script:cfg
})

$btnExportAD.Add_Click({
    try {
        $adData = Invoke-ADUserExport
        $UsersState.Rows = $adData; $UsersState.Headers = $adData[0].PSObject.Properties.Name
        $txtUsersPath.Text = "[Live AD Export Data] - $($adData.Count) users"
        
        $defaultADProps = "samaccountname","userprincipalname","mail","OU","TOP_OU","mailnickname","msExchRemoteRecipientType",
                          "msExchRecipientDisplayType","msExchRecipientTypeDetails","extensionattribute4","enabled",
                          "lastlogondate","PrimarySMTPAddress","Targetaddress","proxyaddresses","Mailbox_Typ",
                          "recipienttype","Serverlocation","pwdLastSet","DisplayName","Vorname","Nachname",
                          "Displayname_email","Office","First2","Last2","Displayname2","Displayname_mail2"
                          
        Populate-FileUI -headers $UsersState.Headers -comboKey $cmbUsersKey -dgvSelect $dgvUsersSelect -gridPreview $gridUsers -rows $UsersState.Rows -DefaultSelection $defaultADProps
        $statusText.Text = "Successfully loaded $($adData.Count) users from Active Directory."
    } catch {
        Show-MessageBox "Failed to export from AD: $($_.Exception.Message)" "AD Export Error" "Error"
    }
})

$btnBrowseUsers.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "CSV/TSV Files|*.csv;*.tsv;*.txt"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtUsersPath.Text = $ofd.FileName
        $res = Detect-Csv -Path $ofd.FileName
        $UsersState.Rows = $res.Rows; $UsersState.Headers = $res.Headers
        Populate-FileUI -headers $UsersState.Headers -comboKey $cmbUsersKey -dgvSelect $dgvUsersSelect -gridPreview $gridUsers -rows $UsersState.Rows
    }
})

$btnBrowseData.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "CSV/TSV Files|*.csv;*.tsv;*.txt"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $cmbData.Text = $ofd.FileName
        $res = Detect-Csv -Path $ofd.FileName
        $DataState.Rows = $res.Rows; $DataState.Headers = $res.Headers
        Populate-FileUI -headers $DataState.Headers -comboKey $cmbDataKey -dgvSelect $dgvDataSelect -gridPreview $gridData -rows $DataState.Rows
        $script:cfg.RecentData = Update-RecentList $script:cfg.RecentData $ofd.FileName
    }
})

$btnExport.Add_Click({
    try {
        if (-not $UsersState.Rows) { throw "Please load a Users source first (from AD or file)." }
        if (-not $DataState.Rows) { throw "Please load a Data file first." }
        if (-not $cmbUsersKey.SelectedItem) { throw "Please select a join key for the Users source." }
        if (-not $cmbDataKey.SelectedItem) { throw "Please select a join key for the Data file." }
        if ([string]::IsNullOrWhiteSpace($txtOut.Text)) { $txtOut.Text = Set-DefaultExportPath -cfg $script:cfg; throw "Export path was empty. A default path has been set." }

        $delimiter = ($cmbDelimiter.SelectedItem -split ' ')[1].Trim('()')
        if ($delimiter -eq 'Tab') { $delimiter = "`t" }

        Do-JoinAndExport -UsersRows $UsersState.Rows -DataRows $DataState.Rows `
            -UsersKey $cmbUsersKey.SelectedItem -DataKey $cmbDataKey.SelectedItem `
            -UsersColsSelected @($dgvUsersSelect.SelectedItems | ForEach-Object { $_.Property }) `
            -DataColsSelected @($dgvDataSelect.SelectedItems | ForEach-Object { $_.Property }) `
            -OutPath $txtOut.Text -StatusTextBlock $statusText `
            -Delimiter $delimiter -IncludeOnlyMatched $chkIncludeOnlyMatched.IsChecked
        
        Show-MessageBox "Export was successful!`n`nFile saved to:`n$($txtOut.Text)"
    } catch { Show-MessageBox "Export failed: $($_.Exception.Message)" "Error" "Error" }
})

# Other event handlers...
$btnPreviewMatched.Add_Click({
    # Logic for previewing matched data...
})
$btnChange.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = "CSV File|*.csv"
    $sfd.FileName = [System.IO.Path]::GetFileName($txtOut.Text); $sfd.InitialDirectory = [System.IO.Path]::GetDirectoryName($txtOut.Text)
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOut.Text = $sfd.FileName
    }
})

# ---[ 8. Show Window ]---
$window.ShowDialog() | Out-Null
