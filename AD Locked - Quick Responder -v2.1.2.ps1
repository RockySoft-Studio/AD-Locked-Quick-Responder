<#
.SYNOPSIS
    AD Locked - Quick Responder v2.1.2
.DESCRIPTION
    Provides a GUI to query all locked user accounts with auto-refresh, sound alerts, 
    TTS notifications, auto-unlock, filtering, and RDS source identification.
.AUTHOR
    Coded by Thundermist Health Center, IT Team
.NOTES
    Requires Active Directory PowerShell module
    Must run on a domain controller or machine with RSAT tools installed
#>

#Requires -Modules ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Speech

[System.Windows.Forms.Application]::EnableVisualStyles()
$ProgressPreference = 'SilentlyContinue'

# Initialize TTS - try Windows 11 natural voices first, fall back to legacy
$script:synthesizer = $null
$script:useModernTTS = $false
$script:modernSynth = $null
$script:mediaPlayer = $null

try {
    # Try to load Windows 11 Natural Voice API
    $null = [Windows.Media.SpeechSynthesis.SpeechSynthesizer, Windows.Media.SpeechSynthesis, ContentType = WindowsRuntime]
    $null = [Windows.Media.Playback.MediaPlayer, Windows.Media.Playback, ContentType = WindowsRuntime]
    $null = [Windows.Media.Core.MediaSource, Windows.Media.Core, ContentType = WindowsRuntime]
    $null = [Windows.Storage.Streams.RandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]

    $script:modernSynth = [Windows.Media.SpeechSynthesis.SpeechSynthesizer]::new()
    $script:mediaPlayer = [Windows.Media.Playback.MediaPlayer]::new()

    # Try to find a natural voice (Jenny or Guy are common on Windows 11)
    $naturalVoice = $script:modernSynth.AllVoices | Where-Object {
        $_.DisplayName -match "Natural" -or $_.DisplayName -match "Jenny" -or $_.DisplayName -match "Guy"
    } | Select-Object -First 1

    if ($naturalVoice) {
        $script:modernSynth.Voice = $naturalVoice
        $script:useModernTTS = $true
    } else {
        # No natural voice found, use first available
        $firstVoice = $script:modernSynth.AllVoices | Select-Object -First 1
        if ($firstVoice) {
            $script:modernSynth.Voice = $firstVoice
            $script:useModernTTS = $true
        }
    }
} catch {
    # Windows 11 API not available, fall back to legacy
    $script:useModernTTS = $false
}

# Fall back to legacy System.Speech if modern TTS not available
if (-not $script:useModernTTS) {
    try {
        $script:synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer
        $script:synthesizer.Rate = 0
        $script:synthesizer.Volume = 100
    } catch {
        Write-Host "TTS initialization failed: $_"
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Locked - Quick Responder v2.1.2"
$form.Size = New-Object System.Drawing.Size(1250, 975)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = [System.Drawing.Color]::White
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

# Download logo
try {
    $logoUrl = "https://www.thundermisthealth.org/wp-content/uploads/2024/10/Thundermist-logo-blue-.png"
    $webClient = New-Object System.Net.WebClient
    $webClient.Headers.Add("User-Agent", "PowerShell")
    $logoBytes = $webClient.DownloadData($logoUrl)
    $ms = New-Object System.IO.MemoryStream(,$logoBytes)
    $logoImage = [System.Drawing.Image]::FromStream($ms)
    $logoPictureBox = New-Object System.Windows.Forms.PictureBox
    $logoPictureBox.Image = $logoImage
    $logoPictureBox.SizeMode = "Zoom"
    $logoPictureBox.Size = New-Object System.Drawing.Size(150, 45)
    $logoPictureBox.Location = New-Object System.Drawing.Point(20, 10)
    $form.Controls.Add($logoPictureBox)
} catch {
    $logoLabel = New-Object System.Windows.Forms.Label
    $logoLabel.Text = "Thundermist"
    $logoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $logoLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
    $logoLabel.Location = New-Object System.Drawing.Point(20, 15)
    $logoLabel.AutoSize = $true
    $form.Controls.Add($logoLabel)
}

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(180, 12)
$titleLabel.Size = New-Object System.Drawing.Size(400, 25)
$titleLabel.Text = "AD Locked - Quick Responder"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$form.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Location = New-Object System.Drawing.Point(180, 38)
$subtitleLabel.Size = New-Object System.Drawing.Size(400, 15)
$subtitleLabel.Text = "(C)2026 Coded by Thundermist Health Center, IT Team"
$subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$subtitleLabel.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($subtitleLabel)

# Info button (top right corner)
$infoButton = New-Object System.Windows.Forms.Button
$infoButton.Text = "i"
$infoButton.Size = New-Object System.Drawing.Size(28, 28)
$infoButton.Location = New-Object System.Drawing.Point(1185, 15)
$infoButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$infoButton.ForeColor = [System.Drawing.Color]::White
$infoButton.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$infoButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$infoButton.FlatAppearance.BorderSize = 0
$infoButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($infoButton)

$infoButton.Add_Click({
    $changelogText = @"
AD Locked - Quick Responder v2.1.2

=== v2.1.2 (Current) ===
- Added Windows 11 Natural Voice TTS support (auto fallback to legacy on Win10)
- Added Info button (i) for viewing changelog
- Removed Stop button - use checkbox to stop auto-unlock
- Fixed auto-unlock controls clickable before DC info loaded
- Adjusted UI spacing

=== v2.1.1 ===
- Added slow operation warning for "Query from ALL DCs"
- Moved Unlock/Query buttons above Selected User Details
- TTS checkboxes moved to same row as auto-unlock controls
- Improved separator line display

=== v2.1.0 ===
- Renamed "Query Lockout Source" to "Query This Locked-out User from ALL DCs"
- Renamed TTS labels for better clarity:
  * "TTS: Upcoming" -> "TTS: Read-out Upcoming Unlocked Users"
  * "TTS: Unlocked" -> "TTS: Read-out Unlocked Users"
- Query status bar moved to bottom domain panel

=== v2.0.0 ===
- Major UI layout reorganization
- Enhanced auto-unlock queue display area
- Improved spacing between sections

=== v1.5.0 (Base) ===
Core Features:
- Async query locked users (UI non-blocking)
- Query from Main DC or All DCs
- Real-time locked users list (sorted by Lockout Time)
- User details panel with full information
- Manual unlock with confirmation dialog
- Export CSV functionality

Auto Refresh: Off / 1 / 2 / 3 / 5 / 10 / 15 min intervals

Notifications:
- Sound Alert for new locked users
- TTS with selectable content (Username/Display Name/Department/Title/Lockout Time)
- Mute option (doesn't affect Auto-unlock TTS)
- TTS Log for all announcements

Auto-Unlock:
- Configurable delay: 1 / 2 / 3 / 5 / 10 min
- Last Logon limit: 1-15 days or No limit
- Bad Logon Count threshold: >= 0-5
- Enabled accounts only
- TTS: Upcoming (1 min warning) & Unlocked notification
- Real-time queue display

Filters:
- Account: All / Enabled / Disabled
- Last Logon: All / 7 / 15 / 30 / 60 / 90 days
- Search: Username / Display Name / Department / Email
"@
    [System.Windows.Forms.MessageBox]::Show($changelogText, "About & Changelog", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Separator
$separator1 = New-Object System.Windows.Forms.Label
$separator1.Location = New-Object System.Drawing.Point(20, 60)
$separator1.Size = New-Object System.Drawing.Size(1195, 2)
$separator1.BorderStyle = "Fixed3D"
$form.Controls.Add($separator1)

# Section 1: Query Controls
$lblSection1 = New-Object System.Windows.Forms.Label
$lblSection1.Location = New-Object System.Drawing.Point(20, 70)
$lblSection1.Size = New-Object System.Drawing.Size(400, 22)
$lblSection1.Text = "Query Controls"
$lblSection1.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblSection1.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$form.Controls.Add($lblSection1)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh Locked-out Info from Main DC"
$refreshButton.Size = New-Object System.Drawing.Size(230, 30)
$refreshButton.Location = New-Object System.Drawing.Point(20, 95)
$refreshButton.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$refreshButton.ForeColor = [System.Drawing.Color]::White
$refreshButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$refreshButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($refreshButton)

$queryAllDCButton = New-Object System.Windows.Forms.Button
$queryAllDCButton.Text = "Refresh Locked-out Info from All DCs (slow)"
$queryAllDCButton.Size = New-Object System.Drawing.Size(270, 30)
$queryAllDCButton.Location = New-Object System.Drawing.Point(260, 95)
$queryAllDCButton.BackColor = [System.Drawing.Color]::FromArgb(156, 39, 176)
$queryAllDCButton.ForeColor = [System.Drawing.Color]::White
$queryAllDCButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($queryAllDCButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Size = New-Object System.Drawing.Size(70, 30)
$cancelButton.Location = New-Object System.Drawing.Point(540, 95)
$cancelButton.BackColor = [System.Drawing.Color]::FromArgb(244, 67, 54)
$cancelButton.ForeColor = [System.Drawing.Color]::White
$cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$cancelButton.Enabled = $false
$form.Controls.Add($cancelButton)

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "Search:"
$searchLabel.Location = New-Object System.Drawing.Point(620, 100)
$searchLabel.AutoSize = $true
$form.Controls.Add($searchLabel)

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Size = New-Object System.Drawing.Size(150, 25)
$searchBox.Location = New-Object System.Drawing.Point(670, 98)
$form.Controls.Add($searchBox)

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "Export CSV"
$exportButton.Size = New-Object System.Drawing.Size(100, 30)
$exportButton.Location = New-Object System.Drawing.Point(1115, 95)
$exportButton.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
$exportButton.ForeColor = [System.Drawing.Color]::White
$exportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($exportButton)

# Separator between Query Controls and Notification Settings
$separator2 = New-Object System.Windows.Forms.Label
$separator2.Location = New-Object System.Drawing.Point(20, 135)
$separator2.Size = New-Object System.Drawing.Size(1195, 2)
$separator2.BorderStyle = "Fixed3D"
$form.Controls.Add($separator2)

# Section 2: Notification Settings
$lblSection2 = New-Object System.Windows.Forms.Label
$lblSection2.Location = New-Object System.Drawing.Point(20, 143)
$lblSection2.Size = New-Object System.Drawing.Size(400, 22)
$lblSection2.Text = "Notification Settings"
$lblSection2.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblSection2.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$form.Controls.Add($lblSection2)

$autoRefreshLabel = New-Object System.Windows.Forms.Label
$autoRefreshLabel.Text = "Auto Refresh:"
$autoRefreshLabel.Location = New-Object System.Drawing.Point(20, 170)
$autoRefreshLabel.AutoSize = $true
$form.Controls.Add($autoRefreshLabel)

$autoRefreshCombo = New-Object System.Windows.Forms.ComboBox
$autoRefreshCombo.Size = New-Object System.Drawing.Size(80, 25)
$autoRefreshCombo.Location = New-Object System.Drawing.Point(100, 168)
$autoRefreshCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$autoRefreshCombo.Items.AddRange(@("Off", "1 min", "2 min", "3 min", "5 min", "10 min", "15 min"))
$autoRefreshCombo.SelectedIndex = 0
$form.Controls.Add($autoRefreshCombo)

$soundAlertCheckbox = New-Object System.Windows.Forms.CheckBox
$soundAlertCheckbox.Text = "Sound Alert"
$soundAlertCheckbox.Location = New-Object System.Drawing.Point(195, 170)
$soundAlertCheckbox.AutoSize = $true
$soundAlertCheckbox.Checked = $true
$form.Controls.Add($soundAlertCheckbox)

$ttsCheckbox = New-Object System.Windows.Forms.CheckBox
$ttsCheckbox.Text = "TTS Notification"
$ttsCheckbox.Location = New-Object System.Drawing.Point(295, 170)
$ttsCheckbox.AutoSize = $true
$ttsCheckbox.Checked = $true
$form.Controls.Add($ttsCheckbox)

$ttsMuteCheckbox = New-Object System.Windows.Forms.CheckBox
$ttsMuteCheckbox.Text = "Mute"
$ttsMuteCheckbox.Location = New-Object System.Drawing.Point(420, 170)
$ttsMuteCheckbox.AutoSize = $true
$ttsMuteCheckbox.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($ttsMuteCheckbox)

$testSoundButton = New-Object System.Windows.Forms.Button
$testSoundButton.Text = "Test TTS Read-Out Function"
$testSoundButton.Size = New-Object System.Drawing.Size(170, 25)
$testSoundButton.Location = New-Object System.Drawing.Point(480, 168)
$testSoundButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$testSoundButton.ForeColor = [System.Drawing.Color]::White
$testSoundButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($testSoundButton)

$nextRefreshLabel = New-Object System.Windows.Forms.Label
$nextRefreshLabel.Text = ""
$nextRefreshLabel.Location = New-Object System.Drawing.Point(660, 170)
$nextRefreshLabel.Size = New-Object System.Drawing.Size(200, 20)
$nextRefreshLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$form.Controls.Add($nextRefreshLabel)

# TTS Content Options
$ttsOptionsLabel = New-Object System.Windows.Forms.Label
$ttsOptionsLabel.Text = "TTS Content:"
$ttsOptionsLabel.Location = New-Object System.Drawing.Point(20, 200)
$ttsOptionsLabel.AutoSize = $true
$form.Controls.Add($ttsOptionsLabel)

$ttsUsernameCheckbox = New-Object System.Windows.Forms.CheckBox
$ttsUsernameCheckbox.Text = "Username"
$ttsUsernameCheckbox.Location = New-Object System.Drawing.Point(100, 200)
$ttsUsernameCheckbox.AutoSize = $true
$ttsUsernameCheckbox.Checked = $true
$form.Controls.Add($ttsUsernameCheckbox)

$ttsDisplayNameCheckbox = New-Object System.Windows.Forms.CheckBox
$ttsDisplayNameCheckbox.Text = "Display Name"
$ttsDisplayNameCheckbox.Location = New-Object System.Drawing.Point(195, 200)
$ttsDisplayNameCheckbox.AutoSize = $true
$form.Controls.Add($ttsDisplayNameCheckbox)

$ttsDepartmentCheckbox = New-Object System.Windows.Forms.CheckBox
$ttsDepartmentCheckbox.Text = "Department"
$ttsDepartmentCheckbox.Location = New-Object System.Drawing.Point(305, 200)
$ttsDepartmentCheckbox.AutoSize = $true
$form.Controls.Add($ttsDepartmentCheckbox)

$ttsTitleCheckbox = New-Object System.Windows.Forms.CheckBox
$ttsTitleCheckbox.Text = "Title"
$ttsTitleCheckbox.Location = New-Object System.Drawing.Point(405, 200)
$ttsTitleCheckbox.AutoSize = $true
$form.Controls.Add($ttsTitleCheckbox)

$ttsTimeCheckbox = New-Object System.Windows.Forms.CheckBox
$ttsTimeCheckbox.Text = "Lockout Time"
$ttsTimeCheckbox.Location = New-Object System.Drawing.Point(465, 200)
$ttsTimeCheckbox.AutoSize = $true
$ttsTimeCheckbox.Checked = $true
$form.Controls.Add($ttsTimeCheckbox)

# Separator between Notification Settings and Auto-Unlock Settings
$separator3 = New-Object System.Windows.Forms.Label
$separator3.Location = New-Object System.Drawing.Point(20, 230)
$separator3.Size = New-Object System.Drawing.Size(1195, 2)
$separator3.BorderStyle = "Fixed3D"
$form.Controls.Add($separator3)

# Section 3: Auto-Unlock Settings
$lblSection3 = New-Object System.Windows.Forms.Label
$lblSection3.Location = New-Object System.Drawing.Point(20, 238)
$lblSection3.Size = New-Object System.Drawing.Size(400, 22)
$lblSection3.Text = "Auto-Unlock Settings"
$lblSection3.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblSection3.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$form.Controls.Add($lblSection3)

$autoUnlockCheckbox = New-Object System.Windows.Forms.CheckBox
$autoUnlockCheckbox.Text = "Auto-unlock $([char]0x2265)"
$autoUnlockCheckbox.Location = New-Object System.Drawing.Point(20, 266)
$autoUnlockCheckbox.AutoSize = $true
$autoUnlockCheckbox.Enabled = $false
$form.Controls.Add($autoUnlockCheckbox)

$autoUnlockDelayCombo = New-Object System.Windows.Forms.ComboBox
$autoUnlockDelayCombo.Size = New-Object System.Drawing.Size(70, 25)
$autoUnlockDelayCombo.Location = New-Object System.Drawing.Point(125, 264)
$autoUnlockDelayCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$autoUnlockDelayCombo.Items.AddRange(@("1 min", "2 min", "3 min", "5 min", "10 min"))
$autoUnlockDelayCombo.SelectedIndex = 3
$autoUnlockDelayCombo.Enabled = $false
$form.Controls.Add($autoUnlockDelayCombo)

$lastLogonLabel = New-Object System.Windows.Forms.Label
$lastLogonLabel.Text = "of lockout, last logon $([char]0x2264)"
$lastLogonLabel.Location = New-Object System.Drawing.Point(195, 268)
$lastLogonLabel.AutoSize = $true
$form.Controls.Add($lastLogonLabel)

$lastLogonDaysCombo = New-Object System.Windows.Forms.ComboBox
$lastLogonDaysCombo.Size = New-Object System.Drawing.Size(70, 25)
$lastLogonDaysCombo.Location = New-Object System.Drawing.Point(330, 264)
$lastLogonDaysCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$lastLogonDaysCombo.Items.AddRange(@("1 day", "2 days", "3 days", "5 days", "7 days", "10 days", "15 days", "30 days", "60 days", "No limit"))
$lastLogonDaysCombo.SelectedIndex = 5
$lastLogonDaysCombo.Enabled = $false
$form.Controls.Add($lastLogonDaysCombo)

$badLogonLabel = New-Object System.Windows.Forms.Label
$badLogonLabel.Text = ", bad logon $([char]0x2265)"
$badLogonLabel.Location = New-Object System.Drawing.Point(405, 268)
$badLogonLabel.AutoSize = $true
$form.Controls.Add($badLogonLabel)

$badLogonCountCombo = New-Object System.Windows.Forms.ComboBox
$badLogonCountCombo.Size = New-Object System.Drawing.Size(55, 25)
$badLogonCountCombo.Location = New-Object System.Drawing.Point(495, 264)
$badLogonCountCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$badLogonCountCombo.Items.AddRange(@("0", "1", "2", "3", "4", "5"))
$badLogonCountCombo.SelectedIndex = 0
$badLogonCountCombo.Enabled = $false
$form.Controls.Add($badLogonCountCombo)

$autoUnlockNoteLabel = New-Object System.Windows.Forms.Label
$autoUnlockNoteLabel.Text = "(Enabled only)"
$autoUnlockNoteLabel.Location = New-Object System.Drawing.Point(555, 268)
$autoUnlockNoteLabel.AutoSize = $true
$autoUnlockNoteLabel.ForeColor = [System.Drawing.Color]::Gray
$autoUnlockNoteLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$form.Controls.Add($autoUnlockNoteLabel)

$ttsUpcomingUnlockCheckbox = New-Object System.Windows.Forms.CheckBox
$ttsUpcomingUnlockCheckbox.Text = "TTS: Read-out Upcoming Unlocked Users (1 min)"
$ttsUpcomingUnlockCheckbox.Location = New-Object System.Drawing.Point(680, 266)
$ttsUpcomingUnlockCheckbox.AutoSize = $true
$ttsUpcomingUnlockCheckbox.Checked = $true
$ttsUpcomingUnlockCheckbox.Enabled = $false
$form.Controls.Add($ttsUpcomingUnlockCheckbox)

$ttsUnlockedCheckbox = New-Object System.Windows.Forms.CheckBox
$ttsUnlockedCheckbox.Text = "TTS: Read-out Unlocked Users"
$ttsUnlockedCheckbox.Location = New-Object System.Drawing.Point(990, 266)
$ttsUnlockedCheckbox.AutoSize = $true
$ttsUnlockedCheckbox.Checked = $true
$ttsUnlockedCheckbox.Enabled = $false
$form.Controls.Add($ttsUnlockedCheckbox)

# Auto-unlock queue label
$autoUnlockQueueLabel = New-Object System.Windows.Forms.Label
$autoUnlockQueueLabel.Text = "Queue:"
$autoUnlockQueueLabel.Location = New-Object System.Drawing.Point(20, 291)
$autoUnlockQueueLabel.Size = New-Object System.Drawing.Size(50, 20)
$autoUnlockQueueLabel.ForeColor = [System.Drawing.Color]::Gray
$autoUnlockQueueLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$autoUnlockQueueLabel.Visible = $false
$form.Controls.Add($autoUnlockQueueLabel)

# Auto-unlock status on its own line for more space
$autoUnlockStatusLabel = New-Object System.Windows.Forms.Label
$autoUnlockStatusLabel.Text = ""
$autoUnlockStatusLabel.Location = New-Object System.Drawing.Point(70, 291)
$autoUnlockStatusLabel.Size = New-Object System.Drawing.Size(1130, 20)
$autoUnlockStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(200, 80, 0)
$autoUnlockStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($autoUnlockStatusLabel)

# Domain info panel (more visible location)
$domainPanel = New-Object System.Windows.Forms.Panel
$domainPanel.Location = New-Object System.Drawing.Point(20, 908)
$domainPanel.Size = New-Object System.Drawing.Size(1195, 28)
$domainPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 248, 255)
$domainPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($domainPanel)

$domainLabel = New-Object System.Windows.Forms.Label
$domainLabel.Text = "Domain: Loading... | Primary DC: Loading..."
$domainLabel.Location = New-Object System.Drawing.Point(10, 5)
$domainLabel.Size = New-Object System.Drawing.Size(700, 18)
$domainLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$domainLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$domainPanel.Controls.Add($domainLabel)

# Status label - at right edge, same position as progress bar
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready"
$statusLabel.Location = New-Object System.Drawing.Point(900, 5)
$statusLabel.Size = New-Object System.Drawing.Size(288, 18)
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$statusLabel.ForeColor = [System.Drawing.Color]::Gray
$domainPanel.Controls.Add($statusLabel)

# Progress bar - same position as status label, shown during queries
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(1020, 3)
$progressBar.Size = New-Object System.Drawing.Size(168, 20)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$progressBar.MarqueeAnimationSpeed = 30
$progressBar.Visible = $false
$domainPanel.Controls.Add($progressBar)

# Separator above Locked Users
$separator4 = New-Object System.Windows.Forms.Label
$separator4.Location = New-Object System.Drawing.Point(20, 321)
$separator4.Size = New-Object System.Drawing.Size(1195, 2)
$separator4.BorderStyle = "Fixed3D"
$form.Controls.Add($separator4)

# Section 4: Locked Users
$lblSection4 = New-Object System.Windows.Forms.Label
$lblSection4.Location = New-Object System.Drawing.Point(20, 329)
$lblSection4.Size = New-Object System.Drawing.Size(120, 22)
$lblSection4.Text = "Locked Users"
$lblSection4.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblSection4.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$form.Controls.Add($lblSection4)

# Filter controls
$filterLabel = New-Object System.Windows.Forms.Label
$filterLabel.Text = "Filter:"
$filterLabel.Location = New-Object System.Drawing.Point(150, 329)
$filterLabel.AutoSize = $true
$filterLabel.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($filterLabel)

$filterAccountLabel = New-Object System.Windows.Forms.Label
$filterAccountLabel.Text = "Account:"
$filterAccountLabel.Location = New-Object System.Drawing.Point(195, 329)
$filterAccountLabel.AutoSize = $true
$form.Controls.Add($filterAccountLabel)

$filterAccountCombo = New-Object System.Windows.Forms.ComboBox
$filterAccountCombo.Location = New-Object System.Drawing.Point(255, 326)
$filterAccountCombo.Size = New-Object System.Drawing.Size(90, 25)
$filterAccountCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$filterAccountCombo.Items.AddRange(@("All", "Enabled", "Disabled"))
$filterAccountCombo.SelectedIndex = 1
$form.Controls.Add($filterAccountCombo)

$filterLastLogonLabel = New-Object System.Windows.Forms.Label
$filterLastLogonLabel.Text = "Last Logon:"
$filterLastLogonLabel.Location = New-Object System.Drawing.Point(360, 329)
$filterLastLogonLabel.AutoSize = $true
$form.Controls.Add($filterLastLogonLabel)

$filterLastLogonCombo = New-Object System.Windows.Forms.ComboBox
$filterLastLogonCombo.Location = New-Object System.Drawing.Point(435, 326)
$filterLastLogonCombo.Size = New-Object System.Drawing.Size(100, 25)
$filterLastLogonCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$filterLastLogonCombo.Items.AddRange(@("All", "7 days", "15 days", "30 days", "60 days", "90 days"))
$filterLastLogonCombo.SelectedIndex = 3
$form.Controls.Add($filterLastLogonCombo)

$countLabel = New-Object System.Windows.Forms.Label
$countLabel.Text = ""
$countLabel.Location = New-Object System.Drawing.Point(1050, 326)
$countLabel.Size = New-Object System.Drawing.Size(165, 22)
$countLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$countLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($countLabel)

# DataGridView
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(20, 354)
$dataGridView.Size = New-Object System.Drawing.Size(1195, 260)
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.MultiSelect = $false
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.RowHeadersVisible = $false
$dataGridView.BackgroundColor = [System.Drawing.Color]::White
$dataGridView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.ColumnHeadersHeight = 30
$form.Controls.Add($dataGridView)

# RowPrePaint event for highlighting newly locked users
$dataGridView.Add_RowPrePaint({
    param($sender, $e)
    if ($e.RowIndex -ge 0 -and $script:newlyLockedUsers.Count -gt 0) {
        $username = $sender.Rows[$e.RowIndex].Cells["Username"].Value
        if ($username -and $username -in $script:newlyLockedUsers) {
            $sender.Rows[$e.RowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 180)
        }
    }
})

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$unlockMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$unlockMenuItem.Text = "Unlock Selected User"
[void]$contextMenu.Items.Add($unlockMenuItem)
$queryLockoutSourceMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$queryLockoutSourceMenuItem.Text = "Query This Locked-out User from ALL DCs"
[void]$contextMenu.Items.Add($queryLockoutSourceMenuItem)
$dataGridView.ContextMenuStrip = $contextMenu

# Section 5: User Details
$detailPanel = New-Object System.Windows.Forms.GroupBox
$detailPanel.Text = "Selected User Details"
$detailPanel.Location = New-Object System.Drawing.Point(20, 653)
$detailPanel.Size = New-Object System.Drawing.Size(1195, 160)
$detailPanel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$detailPanel.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$form.Controls.Add($detailPanel)

$detailTextBox = New-Object System.Windows.Forms.RichTextBox
$detailTextBox.Location = New-Object System.Drawing.Point(10, 20)
$detailTextBox.Size = New-Object System.Drawing.Size(1175, 100)
$detailTextBox.ReadOnly = $true
$detailTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$detailTextBox.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$detailTextBox.ForeColor = [System.Drawing.Color]::Black
$detailTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$detailPanel.Controls.Add($detailTextBox)

$copySelectedButton = New-Object System.Windows.Forms.Button
$copySelectedButton.Text = "Copy Selected"
$copySelectedButton.Size = New-Object System.Drawing.Size(100, 25)
$copySelectedButton.Location = New-Object System.Drawing.Point(10, 125)
$copySelectedButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$copySelectedButton.ForeColor = [System.Drawing.Color]::White
$copySelectedButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$copySelectedButton.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$detailPanel.Controls.Add($copySelectedButton)

$copyAllButton = New-Object System.Windows.Forms.Button
$copyAllButton.Text = "Copy All"
$copyAllButton.Size = New-Object System.Drawing.Size(80, 25)
$copyAllButton.Location = New-Object System.Drawing.Point(115, 125)
$copyAllButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$copyAllButton.ForeColor = [System.Drawing.Color]::White
$copyAllButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$copyAllButton.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$detailPanel.Controls.Add($copyAllButton)

$unlockButton = New-Object System.Windows.Forms.Button
$unlockButton.Text = "Unlock This User"
$unlockButton.Size = New-Object System.Drawing.Size(120, 25)
$unlockButton.Location = New-Object System.Drawing.Point(20, 623)
$unlockButton.BackColor = [System.Drawing.Color]::FromArgb(200, 80, 0)
$unlockButton.ForeColor = [System.Drawing.Color]::White
$unlockButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$unlockButton.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$form.Controls.Add($unlockButton)

$queryLockoutSourceButton = New-Object System.Windows.Forms.Button
$queryLockoutSourceButton.Text = "Query This Locked-out User from ALL DCs"
$queryLockoutSourceButton.Size = New-Object System.Drawing.Size(260, 25)
$queryLockoutSourceButton.Location = New-Object System.Drawing.Point(150, 623)
$queryLockoutSourceButton.BackColor = [System.Drawing.Color]::FromArgb(156, 39, 176)
$queryLockoutSourceButton.ForeColor = [System.Drawing.Color]::White
$queryLockoutSourceButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$queryLockoutSourceButton.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$form.Controls.Add($queryLockoutSourceButton)

# Section 6: TTS Log (below Selected User Details)
$ttsLogLabel = New-Object System.Windows.Forms.Label
$ttsLogLabel.Text = "TTS Log"
$ttsLogLabel.Location = New-Object System.Drawing.Point(20, 818)
$ttsLogLabel.Size = New-Object System.Drawing.Size(100, 20)
$ttsLogLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$ttsLogLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$form.Controls.Add($ttsLogLabel)

$ttsLogTextBox = New-Object System.Windows.Forms.TextBox
$ttsLogTextBox.Location = New-Object System.Drawing.Point(20, 838)
$ttsLogTextBox.Size = New-Object System.Drawing.Size(1195, 60)
$ttsLogTextBox.Multiline = $true
$ttsLogTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$ttsLogTextBox.ReadOnly = $true
$ttsLogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$ttsLogTextBox.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 240)
$ttsLogTextBox.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$form.Controls.Add($ttsLogTextBox)

# Global variables
$script:lockedUsers = @()
$script:filteredUsers = @()
$script:previousLockedUsers = @()
$script:newlyLockedUsers = @()
$script:runspace = $null
$script:powershell = $null
$script:asyncResult = $null
$script:cancelRequested = $false
$script:isClosing = $false
$script:nextRefreshTime = $null
$script:autoUnlockQueue = @{}
$script:notifiedUpcoming = @{}
$script:dcDataLoaded = $false

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000

$autoRefreshTimer = New-Object System.Windows.Forms.Timer

$countdownTimer = New-Object System.Windows.Forms.Timer
$countdownTimer.Interval = 1000

$autoUnlockTimer = New-Object System.Windows.Forms.Timer
$autoUnlockTimer.Interval = 5000  # Check every 5 seconds for more responsive auto-unlock

function Play-AlertSound {
    try { [System.Media.SystemSounds]::Exclamation.Play() }
    catch { [Console]::Beep(1000, 500) }
}

function Get-FilteredUsers {
    param([array]$users)
    
    if ($null -eq $users -or $users.Count -eq 0) {
        return @()
    }
    
    $filtered = $users
    
    # Filter by Account Status
    switch ($filterAccountCombo.SelectedIndex) {
        1 { $filtered = $filtered | Where-Object { $_.Enabled -eq $true } }   # Enabled only
        2 { $filtered = $filtered | Where-Object { $_.Enabled -eq $false } }  # Disabled only
        # 0 = All, no filter
    }
    
    # Filter by Last Logon
    $now = Get-Date
    switch ($filterLastLogonCombo.SelectedIndex) {
        1 { # 7 days
            $cutoff = $now.AddDays(-7)
            $filtered = $filtered | Where-Object { $_.LastLogonDate -and $_.LastLogonDate -ge $cutoff }
        }
        2 { # 15 days
            $cutoff = $now.AddDays(-15)
            $filtered = $filtered | Where-Object { $_.LastLogonDate -and $_.LastLogonDate -ge $cutoff }
        }
        3 { # 30 days
            $cutoff = $now.AddDays(-30)
            $filtered = $filtered | Where-Object { $_.LastLogonDate -and $_.LastLogonDate -ge $cutoff }
        }
        4 { # 60 days
            $cutoff = $now.AddDays(-60)
            $filtered = $filtered | Where-Object { $_.LastLogonDate -and $_.LastLogonDate -ge $cutoff }
        }
        5 { # 90 days
            $cutoff = $now.AddDays(-90)
            $filtered = $filtered | Where-Object { $_.LastLogonDate -and $_.LastLogonDate -ge $cutoff }
        }
        # 0 = All, no filter
    }
    
    return @($filtered)
}

function Apply-FilterAndRefresh {
    $script:filteredUsers = Get-FilteredUsers -users $script:lockedUsers
    Update-DataGridView $script:filteredUsers
    
    $totalCount = ($script:lockedUsers | Measure-Object).Count
    $filteredCount = ($script:filteredUsers | Measure-Object).Count
    
    if ($totalCount -gt 0) {
        if ($filteredCount -eq $totalCount) {
            $countLabel.Text = "Found $filteredCount locked user(s)"
        } else {
            $countLabel.Text = "Showing $filteredCount of $totalCount locked"
        }
        $countLabel.ForeColor = [System.Drawing.Color]::Red
    } else {
        $countLabel.Text = "No locked users"
        $countLabel.ForeColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
    }
    
    # Update auto-unlock queue based on filtered users
    if ($autoUnlockCheckbox.Checked) {
        Update-AutoUnlockQueue -users $script:filteredUsers
    }
}

function Speak-Notification {
    param([string]$Message)

    # Log to TTS Log textbox first
    try {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $logEntry = "[$timestamp] $Message"
        if ($ttsLogTextBox.Text.Length -gt 0) {
            $ttsLogTextBox.AppendText("`r`n$logEntry")
        } else {
            $ttsLogTextBox.Text = $logEntry
        }
        $ttsLogTextBox.SelectionStart = $ttsLogTextBox.Text.Length
        $ttsLogTextBox.ScrollToCaret()
    } catch { }

    # Use modern TTS if available
    if ($script:useModernTTS -and $script:modernSynth -and $script:mediaPlayer) {
        try {
            $stream = $script:modernSynth.SynthesizeTextToStreamAsync($Message).GetAwaiter().GetResult()
            $mediaSource = [Windows.Media.Core.MediaSource]::CreateFromStream($stream, $stream.ContentType)
            $script:mediaPlayer.Source = $mediaSource
            $script:mediaPlayer.Play()
        } catch {
            Write-Host "Modern TTS Error: $_"
        }
    }
    # Fall back to legacy TTS
    elseif ($null -ne $script:synthesizer) {
        try {
            $script:synthesizer.SpeakAsyncCancelAll()
            $script:synthesizer.SpeakAsync($Message) | Out-Null
        } catch {
            Write-Host "TTS Error: $_"
        }
    }
}

function Build-UserTTSDescription {
    param($user)
    $parts = @()
    
    if ($ttsUsernameCheckbox.Checked -and $user.SamAccountName) {
        $parts += $user.SamAccountName
    }
    if ($ttsDisplayNameCheckbox.Checked -and $user.DisplayName) {
        $parts += $user.DisplayName
    }
    if ($ttsDepartmentCheckbox.Checked -and $user.Department) {
        $parts += "from $($user.Department)"
    }
    if ($ttsTitleCheckbox.Checked -and $user.Title) {
        $parts += $user.Title
    }
    
    $desc = if ($parts.Count -gt 0) { $parts -join ", " } else { $user.SamAccountName }
    
    if ($ttsTimeCheckbox.Checked -and $user.LockoutTime) {
        $desc += " at $($user.LockoutTime.ToString('h:mm tt'))"
    }
    
    return $desc
}

function Get-AutoUnlockDelayMinutes {
    switch ($autoUnlockDelayCombo.SelectedIndex) {
        0 { return 1 }
        1 { return 2 }
        2 { return 3 }
        3 { return 5 }
        4 { return 10 }
        default { return 5 }
    }
}

function Get-LastLogonDaysLimit {
    switch ($lastLogonDaysCombo.SelectedIndex) {
        0 { return 1 }
        1 { return 2 }
        2 { return 3 }
        3 { return 5 }
        4 { return 7 }
        5 { return 10 }
        6 { return 15 }
        7 { return 30 }
        8 { return 60 }
        9 { return 36500 }  # No limit (~100 years)
        default { return 5 }
    }
}

function Get-BadLogonCountThreshold {
    switch ($badLogonCountCombo.SelectedIndex) {
        0 { return 0 }
        1 { return 1 }
        2 { return 2 }
        3 { return 3 }
        4 { return 4 }
        5 { return 5 }
        default { return 0 }
    }
}

function Enable-AutoUnlockControls {
    if ($script:dcDataLoaded) {
        $autoUnlockCheckbox.Enabled = $true
        $autoUnlockDelayCombo.Enabled = $true
        $lastLogonDaysCombo.Enabled = $true
        $badLogonCountCombo.Enabled = $true
        $ttsUpcomingUnlockCheckbox.Enabled = $true
        $ttsUnlockedCheckbox.Enabled = $true
    }
}

function Check-NewLockedUsers {
    param([array]$currentUsers)
    
    $previousNames = @()
    if ($script:previousLockedUsers.Count -gt 0) {
        $previousNames = $script:previousLockedUsers | ForEach-Object { $_.SamAccountName }
    }
    
    $newLockedUsers = @($currentUsers | Where-Object { $_.SamAccountName -notin $previousNames })
    
    # Save newly locked users for highlighting (only if not first load)
    if ($script:previousLockedUsers.Count -gt 0) {
        $script:newlyLockedUsers = @($newLockedUsers | ForEach-Object { $_.SamAccountName })
    } else {
        $script:newlyLockedUsers = @()
    }
    
    # Only alert if we had previous data (not first load)
    if ($newLockedUsers.Count -gt 0 -and $script:previousLockedUsers.Count -gt 0) {
        if ($soundAlertCheckbox.Checked -and -not $ttsMuteCheckbox.Checked) { 
            Play-AlertSound 
        }
        
        if ($ttsCheckbox.Checked -and -not $ttsMuteCheckbox.Checked) {
            if ($newLockedUsers.Count -eq 1) {
                $user = $newLockedUsers[0]
                $userDesc = Build-UserTTSDescription -user $user
                $message = "Alert: User $userDesc was locked out."
                Speak-Notification -Message $message
            }
            else {
                $message = "Alert: $($newLockedUsers.Count) new users have been locked out. "
                foreach ($user in $newLockedUsers) {
                    $userDesc = Build-UserTTSDescription -user $user
                    $message += "$userDesc. "
                }
                Speak-Notification -Message $message
            }
        }
    }
    
    $script:previousLockedUsers = $currentUsers
    
    # Apply filter and refresh display (this also updates auto-unlock queue)
    Apply-FilterAndRefresh
}

function Update-AutoUnlockQueue {
    param([array]$users)
    
    $delayMinutes = Get-AutoUnlockDelayMinutes
    $lastLogonDaysLimit = Get-LastLogonDaysLimit
    $lastLogonCutoff = (Get-Date).Date.AddDays(-$lastLogonDaysLimit)  # Use .Date for calendar day calculation
    $badLogonThreshold = Get-BadLogonCountThreshold
    
    # Rebuild queue based on filtered users
    $newQueue = @{}
    
    foreach ($user in $users) {
        # Only Enabled accounts (already filtered, but double-check for auto-unlock specific settings)
        if ($user.Enabled -eq $true) {
            # Check BadLogonCount meets threshold
            if ($user.BadLogonCount -ge $badLogonThreshold) {
                # Check LastLogonDate within auto-unlock specific limit
                if ($user.LastLogonDate -and $user.LastLogonDate -ge $lastLogonCutoff) {
                    if ($user.LockoutTime) {
                        $unlockTime = $user.LockoutTime.AddMinutes($delayMinutes)
                        $newQueue[$user.SamAccountName] = @{
                            UnlockTime = $unlockTime
                            DisplayName = $user.DisplayName
                            Department = $user.Department
                            LockoutTime = $user.LockoutTime
                            BadLogonCount = $user.BadLogonCount
                        }
                    }
                }
            }
        }
    }
    
    $script:autoUnlockQueue = $newQueue
    Update-AutoUnlockStatus
}

function Update-AutoUnlockStatus {
    if ($script:autoUnlockQueue.Count -gt 0) {
        $now = Get-Date
        $pendingUsers = @()
        
        foreach ($key in $script:autoUnlockQueue.Keys) {
            $entry = $script:autoUnlockQueue[$key]
            $remaining = $entry.UnlockTime - $now
            
            if ($remaining.TotalSeconds -gt 0) {
                $mins = [math]::Floor($remaining.TotalMinutes)
                $secs = $remaining.Seconds
                $pendingUsers += @{
                    Username = $key
                    RemainingText = "${mins}m${secs}s"
                    TotalSeconds = $remaining.TotalSeconds
                }
            } else {
                # Ready to unlock (0 or negative)
                $pendingUsers += @{
                    Username = $key
                    RemainingText = "now"
                    TotalSeconds = 0
                }
            }
        }
        
        if ($pendingUsers.Count -gt 0) {
            # Sort by TotalSeconds (soonest first)
            $pendingUsers = $pendingUsers | Sort-Object { $_.TotalSeconds }
            
            # Build display string with all usernames and their countdown
            $displayParts = @()
            foreach ($user in $pendingUsers) {
                $displayParts += "$($user.Username)($($user.RemainingText))"
            }
            
            $autoUnlockStatusLabel.Text = "Pending: " + ($displayParts -join ", ")
        } else {
            $autoUnlockStatusLabel.Text = ""
        }
    } else {
        $autoUnlockStatusLabel.Text = ""
    }
}

function Process-AutoUnlockQueue {
    if (-not $autoUnlockCheckbox.Checked) { return }
    if ($script:autoUnlockQueue.Count -eq 0) { return }
    
    $now = Get-Date
    $toUnlock = @()
    $toNotifyUpcoming = @()
    
    # Use Keys copy to avoid modification during enumeration
    $keys = @($script:autoUnlockQueue.Keys)
    
    foreach ($key in $keys) {
        $entry = $script:autoUnlockQueue[$key]
        $remaining = $entry.UnlockTime - $now
        
        # Ready to unlock
        if ($remaining.TotalSeconds -le 0) {
            $toUnlock += $key
        }
        # Within 1 minute - notify upcoming
        elseif ($remaining.TotalSeconds -le 60 -and $remaining.TotalSeconds -gt 0) {
            if (-not $script:notifiedUpcoming.ContainsKey($key)) {
                $toNotifyUpcoming += @{ Username = $key; DisplayName = $entry.DisplayName }
            }
        }
    }
    
    # Notify upcoming unlocks
    foreach ($userInfo in $toNotifyUpcoming) {
        $script:notifiedUpcoming[$userInfo.Username] = $true
        
        if ($ttsUpcomingUnlockCheckbox.Checked) {
            $displayName = if ($userInfo.DisplayName) { $userInfo.DisplayName } else { $userInfo.Username }
            Speak-Notification -Message "Upcoming auto-unlock: $displayName in 1 minute."
        }
        
        Update-Status "Upcoming auto-unlock: $($userInfo.Username) in 1 minute" "Orange"
    }
    
    # Process unlocks
    $unlockSuccess = @()
    foreach ($username in $toUnlock) {
        try {
            Unlock-ADAccount -Identity $username -ErrorAction Stop
            $displayName = $script:autoUnlockQueue[$username].DisplayName
            $script:autoUnlockQueue.Remove($username)
            $script:notifiedUpcoming.Remove($username)
            $unlockSuccess += @{ Username = $username; DisplayName = $displayName }
            Update-Status "Auto-unlocked: $username" "Green"
        }
        catch {
            Update-Status "Failed to auto-unlock ${username}: $($_.Exception.Message)" "Red"
            $script:autoUnlockQueue.Remove($username)
            $script:notifiedUpcoming.Remove($username)
        }
    }
    
    # TTS for unlocked users
    if ($unlockSuccess.Count -gt 0 -and $ttsUnlockedCheckbox.Checked) {
        if ($unlockSuccess.Count -eq 1) {
            $displayName = if ($unlockSuccess[0].DisplayName) { $unlockSuccess[0].DisplayName } else { $unlockSuccess[0].Username }
            Speak-Notification -Message "$displayName has been auto-unlocked."
        }
        else {
            $names = ($unlockSuccess | ForEach-Object { if ($_.DisplayName) { $_.DisplayName } else { $_.Username } }) -join ", "
            Speak-Notification -Message "$($unlockSuccess.Count) users have been auto-unlocked: $names"
        }
    }
    
    # Refresh list if any unlocks happened
    if ($toUnlock.Count -gt 0) {
        Get-LockedOutUsersAsync
    }
    
    Update-AutoUnlockStatus
}

function Unlock-SelectedUser {
    if ($dataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a user to unlock", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $selectedUsername = $dataGridView.SelectedRows[0].Cells["Username"].Value
    $selectedDisplayName = $dataGridView.SelectedRows[0].Cells["Display Name"].Value
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to unlock user:`n`nUsername: $selectedUsername`nDisplay Name: $selectedDisplayName",
        "Confirm Unlock",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Unlock-ADAccount -Identity $selectedUsername -ErrorAction Stop
            
            if ($script:autoUnlockQueue.ContainsKey($selectedUsername)) {
                $script:autoUnlockQueue.Remove($selectedUsername)
                $script:notifiedUpcoming.Remove($selectedUsername)
                Update-AutoUnlockStatus
            }
            
            Update-Status "Successfully unlocked: $selectedUsername" "Green"
            [System.Windows.Forms.MessageBox]::Show("User '$selectedUsername' has been unlocked successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            Get-LockedOutUsersAsync
        }
        catch {
            Update-Status "Failed to unlock: $selectedUsername" "Red"
            [System.Windows.Forms.MessageBox]::Show("Failed to unlock user '$selectedUsername':`n`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

function Set-UIState {
    param([bool]$IsQuerying)
    if ($script:isClosing) { return }
    $refreshButton.Enabled = -not $IsQuerying
    $exportButton.Enabled = -not $IsQuerying
    $queryAllDCButton.Enabled = -not $IsQuerying
    $unlockButton.Enabled = -not $IsQuerying
    $queryLockoutSourceButton.Enabled = -not $IsQuerying
    $cancelButton.Enabled = $IsQuerying
    $progressBar.Visible = $IsQuerying
    $statusLabel.Visible = -not $IsQuerying
    $form.Cursor = if ($IsQuerying) { [System.Windows.Forms.Cursors]::AppStarting } else { [System.Windows.Forms.Cursors]::Default }

    # Pause/resume auto-refresh during query
    if ($IsQuerying) {
        # Pause auto-refresh timer during query
        $autoRefreshTimer.Stop()
        $countdownTimer.Stop()
        $nextRefreshLabel.Text = "Paused..."
    } else {
        # Resume auto-refresh if enabled
        if ($autoRefreshCombo.SelectedIndex -gt 0) {
            $intervalMinutes = switch ($autoRefreshCombo.SelectedIndex) { 1 { 1 }; 2 { 2 }; 3 { 3 }; 4 { 5 }; 5 { 10 }; 6 { 15 }; default { 0 } }
            $script:nextRefreshTime = (Get-Date).AddMinutes($intervalMinutes)
            $autoRefreshTimer.Start()
            $countdownTimer.Start()
        } else {
            $nextRefreshLabel.Text = ""
        }
    }
}

function Update-Status {
    param([string]$Message, [string]$Color = "Gray")
    if ($script:isClosing) { return }
    $statusLabel.Text = $Message
    $statusLabel.ForeColor = switch ($Color) {
        "Green" { [System.Drawing.Color]::FromArgb(76, 175, 80) }
        "Red" { [System.Drawing.Color]::Red }
        "Orange" { [System.Drawing.Color]::FromArgb(200, 80, 0) }
        "Blue" { [System.Drawing.Color]::FromArgb(0, 102, 204) }
        default { [System.Drawing.Color]::Gray }
    }
}

function Cleanup-AsyncResources {
    try { $timer.Stop() } catch { }
    # Don't stop autoRefreshTimer and countdownTimer here - they should keep running
    
    try {
        if ($script:powershell) {
            if ($script:asyncResult -and -not $script:asyncResult.IsCompleted) {
                $script:powershell.Stop()
                Start-Sleep -Milliseconds 100
            }
            $script:powershell.Dispose()
            $script:powershell = $null
        }
    } catch { }
    
    try {
        if ($script:runspace) {
            if ($script:runspace.RunspaceStateInfo.State -eq 'Opened') {
                $script:runspace.Close()
            }
            $script:runspace.Dispose()
            $script:runspace = $null
        }
    } catch { }
    
    $script:asyncResult = $null
}

function Get-DomainInfoAsync {
    Set-UIState -IsQuerying $true
    Update-Status "Loading domain info..." "Orange"
    
    Cleanup-AsyncResources
    
    $script:runspace = [runspacefactory]::CreateRunspace()
    $script:runspace.ApartmentState = "STA"
    $script:runspace.ThreadOptions = "ReuseThread"
    $script:runspace.Open()
    $script:powershell = [powershell]::Create()
    $script:powershell.Runspace = $script:runspace
    [void]$script:powershell.AddScript({
        Import-Module ActiveDirectory -ErrorAction Stop
        $domain = Get-ADDomain
        return @{ DNSRoot = $domain.DNSRoot; PDCEmulator = $domain.PDCEmulator }
    })
    $script:asyncResult = $script:powershell.BeginInvoke()
    $timer.Tag = "DomainInfo"
    $timer.Start()
}

function Get-LockedOutUsersAsync {
    Set-UIState -IsQuerying $true
    Update-Status "Querying Main DC..." "Orange"
    $script:cancelRequested = $false
    
    Cleanup-AsyncResources
    
    $script:runspace = [runspacefactory]::CreateRunspace()
    $script:runspace.ApartmentState = "STA"
    $script:runspace.ThreadOptions = "ReuseThread"
    $script:runspace.Open()
    $script:powershell = [powershell]::Create()
    $script:powershell.Runspace = $script:runspace
    [void]$script:powershell.AddScript({
        Import-Module ActiveDirectory -ErrorAction Stop
        $results = @()
        $lockedAccounts = Search-ADAccount -LockedOut
        foreach ($account in $lockedAccounts) {
            $user = Get-ADUser -Identity $account.DistinguishedName -Properties SamAccountName,DisplayName,EmailAddress,Department,Title,LockedOut,LockoutTime,BadLogonCount,LastBadPasswordAttempt,LastLogonDate,PasswordLastSet,Enabled,Description,DistinguishedName,WhenCreated
            $lockoutTimeConverted = $null
            if ($user.LockoutTime -and $user.LockoutTime -gt 0) { $lockoutTimeConverted = [DateTime]::FromFileTime($user.LockoutTime) }
            $lockoutDuration = $null
            if ($lockoutTimeConverted) {
                $duration = (Get-Date) - $lockoutTimeConverted
                $lockoutDuration = "{0:D2}:{1:D2}:{2:D2}" -f [int]$duration.TotalHours, $duration.Minutes, $duration.Seconds
            }
            $results += [PSCustomObject]@{
                SamAccountName = $user.SamAccountName; DisplayName = $user.DisplayName; Department = $user.Department
                Title = $user.Title; LockoutTime = $lockoutTimeConverted; LockoutDuration = $lockoutDuration
                BadLogonCount = $user.BadLogonCount; LastBadPasswordAttempt = $user.LastBadPasswordAttempt
                LastLogonDate = $user.LastLogonDate; PasswordLastSet = $user.PasswordLastSet
                EmailAddress = $user.EmailAddress; Enabled = $user.Enabled; Description = $user.Description
                DistinguishedName = $user.DistinguishedName; WhenCreated = $user.WhenCreated
            }
        }
        return $results | Sort-Object -Property LockoutTime -Descending
    })
    $script:asyncResult = $script:powershell.BeginInvoke()
    $timer.Tag = "LockedUsers"
    $timer.Start()
}

function Get-LockedOutUsersFromAllDCsAsync {
    Set-UIState -IsQuerying $true
    Update-Status "Querying All DCs for locked users..." "Orange"
    $script:cancelRequested = $false

    Cleanup-AsyncResources

    $script:runspace = [runspacefactory]::CreateRunspace()
    $script:runspace.ApartmentState = "STA"
    $script:runspace.ThreadOptions = "ReuseThread"
    $script:runspace.Open()
    $script:powershell = [powershell]::Create()
    $script:powershell.Runspace = $script:runspace
    [void]$script:powershell.AddScript({
        Import-Module ActiveDirectory -ErrorAction Stop
        $allLockedUsers = @{}
        $domainControllers = Get-ADDomainController -Filter * | Select-Object HostName
        foreach ($dc in $domainControllers) {
            try {
                $lockedAccounts = Search-ADAccount -LockedOut -Server $dc.HostName
                foreach ($account in $lockedAccounts) {
                    if (-not $allLockedUsers.ContainsKey($account.SamAccountName)) {
                        $user = Get-ADUser -Identity $account.DistinguishedName -Server $dc.HostName -Properties SamAccountName,DisplayName,EmailAddress,Department,Title,LockedOut,LockoutTime,BadLogonCount,LastBadPasswordAttempt,LastLogonDate,PasswordLastSet,Enabled,Description,DistinguishedName,WhenCreated
                        $lockoutTimeConverted = $null
                        if ($user.LockoutTime -and $user.LockoutTime -gt 0) { $lockoutTimeConverted = [DateTime]::FromFileTime($user.LockoutTime) }
                        $lockoutDuration = $null
                        if ($lockoutTimeConverted) {
                            $duration = (Get-Date) - $lockoutTimeConverted
                            $lockoutDuration = "{0:D2}:{1:D2}:{2:D2}" -f [int]$duration.TotalHours, $duration.Minutes, $duration.Seconds
                        }
                        $allLockedUsers[$account.SamAccountName] = [PSCustomObject]@{
                            SamAccountName = $user.SamAccountName; DisplayName = $user.DisplayName; Department = $user.Department
                            Title = $user.Title; LockoutTime = $lockoutTimeConverted; LockoutDuration = $lockoutDuration
                            BadLogonCount = $user.BadLogonCount; LastBadPasswordAttempt = $user.LastBadPasswordAttempt
                            LastLogonDate = $user.LastLogonDate; PasswordLastSet = $user.PasswordLastSet
                            EmailAddress = $user.EmailAddress; Enabled = $user.Enabled; Description = $user.Description
                            DistinguishedName = $user.DistinguishedName; WhenCreated = $user.WhenCreated
                        }
                    }
                }
            } catch { }
        }
        return $allLockedUsers.Values | Sort-Object -Property LockoutTime -Descending
    })
    $script:asyncResult = $script:powershell.BeginInvoke()
    $timer.Tag = "LockedUsersAllDCs"
    $timer.Start()
}

function Get-LockoutSourceAsync {
    param([string]$Username)
    if ([string]::IsNullOrWhiteSpace($Username)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a user from the table first", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    Set-UIState -IsQuerying $true
    Update-Status "Querying All DCs..." "Orange"
    $script:cancelRequested = $false
    
    Cleanup-AsyncResources
    
    $script:runspace = [runspacefactory]::CreateRunspace()
    $script:runspace.ApartmentState = "STA"
    $script:runspace.ThreadOptions = "ReuseThread"
    $script:runspace.Open()
    $script:powershell = [powershell]::Create()
    $script:powershell.Runspace = $script:runspace
    [void]$script:powershell.AddScript({
        param($Username, $RDSServers)
        Import-Module ActiveDirectory -ErrorAction Stop
        $output = @{ DCResults = @(); Events = @(); PDC = ""; Error = $null; RDSLockouts = @(); OtherLockouts = @(); Summary = @{ TotalEvents = 0; RDSCount = 0; OtherCount = 0 }; FailedLogons = @() }
        try {
            $domainControllers = Get-ADDomainController -Filter * | Select-Object Name, HostName, Site
            foreach ($dc in $domainControllers) {
                try {
                    $user = Get-ADUser -Identity $Username -Server $dc.HostName -Properties LockedOut,LockoutTime,BadLogonCount,LastBadPasswordAttempt
                    $lockoutTimeConverted = $null
                    if ($user.LockoutTime -and $user.LockoutTime -gt 0) { $lockoutTimeConverted = [DateTime]::FromFileTime($user.LockoutTime) }
                    $output.DCResults += [PSCustomObject]@{ DCName = $dc.Name; DCHostName = $dc.HostName; Site = $dc.Site; LockedOut = $user.LockedOut; LockoutTime = $lockoutTimeConverted; BadLogonCount = $user.BadLogonCount; LastBadPasswordAttempt = $user.LastBadPasswordAttempt; Error = $null }
                } catch {
                    $output.DCResults += [PSCustomObject]@{ DCName = $dc.Name; DCHostName = $dc.HostName; Site = $dc.Site; LockedOut = "Query Failed"; LockoutTime = $null; BadLogonCount = "N/A"; LastBadPasswordAttempt = $null; Error = $_.Exception.Message }
                }
            }
            try {
                $pdc = (Get-ADDomain).PDCEmulator
                $output.PDC = $pdc
                $lockoutEvents = Get-WinEvent -ComputerName $pdc -FilterHashtable @{ LogName = 'Security'; Id = 4740 } -MaxEvents 200 -ErrorAction SilentlyContinue | Where-Object { $_.Properties[0].Value -eq $Username }
                foreach ($event in $lockoutEvents | Select-Object -First 20) {
                    $sourceComputer = $event.Properties[1].Value
                    $sourceComputerUpper = $sourceComputer.ToUpper()
                    $sourceType = "Unknown"; $isRDS = $false
                    foreach ($rds in $RDSServers) { if ($sourceComputerUpper -like "*$($rds.ToUpper())*") { $sourceType = "RDS"; $isRDS = $true; break } }
                    if (-not $isRDS -and $sourceComputerUpper -match '(RDS|RDP|TERM|TS|REMOTE|GATEWAY|RDGW|RDSH|RDCB|VDI|CITRIX|XEN)') { $sourceType = "RDS (Inferred)"; $isRDS = $true }
                    if (-not $isRDS -and $sourceComputer) {
                        try {
                            $computer = Get-ADComputer -Identity $sourceComputer -Properties OperatingSystem, Description -ErrorAction SilentlyContinue
                            if ($computer) {
                                if ($computer.OperatingSystem -match 'Server' -or $computer.Description -match '(RDS|Remote|Terminal)') { $sourceType = "Server" }
                                if ($computer.Description -match '(RDS|Remote|Terminal)') { $sourceType = "RDS (Description Match)"; $isRDS = $true }
                                elseif ($computer.OperatingSystem -match 'Windows 10|Windows 11') { $sourceType = "Workstation" }
                            }
                        } catch { }
                    }
                    if (-not $isRDS -and $sourceType -eq "Unknown") { $sourceType = "Other/Workstation" }
                    $eventInfo = [PSCustomObject]@{ TimeCreated = $event.TimeCreated; SourceComputer = $sourceComputer; SourceType = $sourceType; IsRDS = $isRDS }
                    $output.Events += $eventInfo
                    if ($isRDS) { $output.RDSLockouts += $eventInfo; $output.Summary.RDSCount++ } else { $output.OtherLockouts += $eventInfo; $output.Summary.OtherCount++ }
                    $output.Summary.TotalEvents++
                }
                try {
                    $failedLogonEvents = Get-WinEvent -ComputerName $pdc -FilterHashtable @{ LogName = 'Security'; Id = 4625 } -MaxEvents 500 -ErrorAction SilentlyContinue | Where-Object { $_.Properties[5].Value -eq $Username } | Select-Object -First 20
                    foreach ($event in $failedLogonEvents) {
                        $logonType = $event.Properties[10].Value
                        $logonTypeDesc = switch ($logonType) { 2 { "Interactive" }; 3 { "Network" }; 7 { "Unlock" }; 10 { "Remote Interactive (RDP)" }; 11 { "Cached Interactive" }; default { "Type $logonType" } }
                        $output.FailedLogons += [PSCustomObject]@{ TimeCreated = $event.TimeCreated; WorkstationName = $event.Properties[13].Value; IPAddress = $event.Properties[19].Value; LogonType = $logonTypeDesc; IsRDP = ($logonType -eq 10) }
                    }
                } catch { }
            } catch { $output.EventError = $_.Exception.Message }
        } catch { $output.Error = $_.Exception.Message }
        return $output
    }).AddArgument($Username).AddArgument(@("RDS", "RDGW", "RDSH", "RDCB", "TERM", "REMOTE"))
    $script:asyncResult = $script:powershell.BeginInvoke()
    $script:queryUsername = $Username
    $timer.Tag = "LockoutSource"
    $timer.Start()
}

$timer.Add_Tick({
    if ($script:isClosing) { $timer.Stop(); return }
    if ($null -eq $script:asyncResult) { return }
    
    if ($script:asyncResult.IsCompleted) {
        $timer.Stop()
        try {
            $result = $script:powershell.EndInvoke($script:asyncResult)
            switch ($timer.Tag) {
                "DomainInfo" {
                    if ($result) { 
                        $domainLabel.Text = "Domain: $($result.DNSRoot) | Primary DC: $($result.PDCEmulator)"
                        $script:dcDataLoaded = $true
                        Enable-AutoUnlockControls
                    }
                    Set-UIState -IsQuerying $false
                    Update-Status "Ready" "Gray"
                }
                "LockedUsers" {
                    if ($script:cancelRequested) { Update-Status "Cancelled" "Orange" }
                    else {
                        $script:lockedUsers = $result
                        Check-NewLockedUsers -currentUsers $script:lockedUsers
                        # Note: Check-NewLockedUsers calls Apply-FilterAndRefresh which updates display and count
                        $timestamp = Get-Date -Format "HH:mm:ss"
                        $totalCount = ($script:lockedUsers | Measure-Object).Count
                        $filteredCount = ($script:filteredUsers | Measure-Object).Count
                        if ($totalCount -gt 0) {
                            if ($filteredCount -eq $totalCount) {
                                Update-Status "Found $totalCount @ $timestamp" "Green"
                            } else {
                                Update-Status "$filteredCount of $totalCount @ $timestamp" "Green"
                            }
                        }
                        else {
                            Update-Status "No locked @ $timestamp" "Green"
                        }
                    }
                    Set-UIState -IsQuerying $false
                }
                "LockedUsersAllDCs" {
                    if ($script:cancelRequested) { Update-Status "Cancelled" "Orange" }
                    else {
                        $script:lockedUsers = $result
                        Check-NewLockedUsers -currentUsers $script:lockedUsers
                        $timestamp = Get-Date -Format "HH:mm:ss"
                        $totalCount = ($script:lockedUsers | Measure-Object).Count
                        $filteredCount = ($script:filteredUsers | Measure-Object).Count
                        if ($totalCount -gt 0) {
                            if ($filteredCount -eq $totalCount) {
                                Update-Status "All DCs: $totalCount @ $timestamp" "Green"
                            } else {
                                Update-Status "All DCs: $filteredCount of $totalCount @ $timestamp" "Green"
                            }
                        }
                        else {
                            Update-Status "All DCs: No locked @ $timestamp" "Green"
                        }
                    }
                    Set-UIState -IsQuerying $false
                }
                "LockoutSource" {
                    if ($script:cancelRequested) { Update-Status "Cancelled" "Orange"; $detailTextBox.Text = "Query cancelled" }
                    else {
                        $resultText = "Lockout Source Analysis for '$($script:queryUsername)'`r`n" + ("=" * 80) + "`r`n`r`n"
                        if ($result.Summary.TotalEvents -gt 0) {
                            $resultText += "[LOCKOUT SOURCE SUMMARY]`r`n" + ("-" * 40) + "`r`n"
                            $resultText += "Total Lockout Events: $($result.Summary.TotalEvents)`r`n"
                            $rdsPercent = [math]::Round(($result.Summary.RDSCount / $result.Summary.TotalEvents) * 100, 1)
                            $otherPercent = [math]::Round(($result.Summary.OtherCount / $result.Summary.TotalEvents) * 100, 1)
                            $resultText += "From RDS Systems: $($result.Summary.RDSCount) ($rdsPercent%)`r`n"
                            $resultText += "From Other Systems: $($result.Summary.OtherCount) ($otherPercent%)`r`n`r`n"
                        }
                        if ($result.RDSLockouts.Count -gt 0) {
                            $resultText += "[RDS SYSTEM LOCKOUT EVENTS]`r`n" + ("-" * 40) + "`r`n"
                            foreach ($event in $result.RDSLockouts) { $resultText += "Time: $($event.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))  |  Source: $($event.SourceComputer)  |  Type: $($event.SourceType)`r`n" }
                            $resultText += "`r`n"
                        }
                        if ($result.OtherLockouts.Count -gt 0) {
                            $resultText += "[OTHER SYSTEM LOCKOUT EVENTS]`r`n" + ("-" * 40) + "`r`n"
                            foreach ($event in $result.OtherLockouts) { $resultText += "Time: $($event.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))  |  Source: $($event.SourceComputer)  |  Type: $($event.SourceType)`r`n" }
                            $resultText += "`r`n"
                        }
                        $resultText += "[DOMAIN CONTROLLER STATUS]`r`n" + ("-" * 40) + "`r`n"
                        foreach ($r in $result.DCResults) { $resultText += "DC: $($r.DCName) | Locked: $($r.LockedOut) | Bad Logon Count: $($r.BadLogonCount)`r`n" }
                        $detailTextBox.Text = $resultText
                        Update-Status "DC query completed" "Green"
                    }
                    Set-UIState -IsQuerying $false
                }
            }
        } catch {
            if (-not $script:isClosing) {
                Update-Status "Error: $($_.Exception.Message)" "Red"
                Set-UIState -IsQuerying $false
            }
        }
    }
})

$autoRefreshTimer.Add_Tick({
    if ($script:isClosing) { return }
    
    # Always update next refresh time
    $intervalMinutes = switch ($autoRefreshCombo.SelectedIndex) { 1 { 1 }; 2 { 2 }; 3 { 3 }; 4 { 5 }; 5 { 10 }; 6 { 15 }; default { 0 } }
    if ($intervalMinutes -gt 0) {
        $script:nextRefreshTime = (Get-Date).AddMinutes($intervalMinutes)
        # Ensure countdown timer is running
        if (-not $countdownTimer.Enabled) {
            $countdownTimer.Start()
        }
    }
    
    # Only start query if not already querying
    if ($refreshButton.Enabled) { 
        Get-LockedOutUsersAsync 
    }
})

$countdownTimer.Add_Tick({
    if ($script:isClosing) { return }
    if ($script:nextRefreshTime -and $autoRefreshCombo.SelectedIndex -gt 0) {
        $remaining = $script:nextRefreshTime - (Get-Date)
        if ($remaining.TotalSeconds -gt 0) { 
            $nextRefreshLabel.Text = "Next refresh: $([math]::Floor($remaining.TotalMinutes)):$($remaining.Seconds.ToString('00'))" 
        }
        else { 
            $nextRefreshLabel.Text = "Refreshing..." 
        }
    } else { 
        $nextRefreshLabel.Text = "" 
    }
    
    # Update auto-unlock status every second
    Update-AutoUnlockStatus
})

$autoUnlockTimer.Add_Tick({
    if ($script:isClosing) { return }
    Process-AutoUnlockQueue
})

function Update-DataGridView {
    param([array]$users)
    if ($script:isClosing) { return }
    $dataGridView.DataSource = $null
    if ($users.Count -gt 0) {
        $dataTable = New-Object System.Data.DataTable
        $dataTable.Columns.Add("Username", [string]) | Out-Null
        $dataTable.Columns.Add("Display Name", [string]) | Out-Null
        $dataTable.Columns.Add("Department", [string]) | Out-Null
        $dataTable.Columns.Add("Title", [string]) | Out-Null
        $dataTable.Columns.Add("Lockout Time", [string]) | Out-Null
        $dataTable.Columns.Add("Duration", [string]) | Out-Null
        $dataTable.Columns.Add("Bad Logon Count", [int]) | Out-Null
        $dataTable.Columns.Add("Last Bad Password", [string]) | Out-Null
        $dataTable.Columns.Add("Last Logon", [string]) | Out-Null
        $dataTable.Columns.Add("Account", [string]) | Out-Null
        foreach ($user in $users) {
            $row = $dataTable.NewRow()
            $row["Username"] = $user.SamAccountName
            $row["Display Name"] = $user.DisplayName
            $row["Department"] = $user.Department
            $row["Title"] = $user.Title
            $row["Lockout Time"] = if ($user.LockoutTime) { $user.LockoutTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
            $row["Duration"] = $user.LockoutDuration
            $row["Bad Logon Count"] = $user.BadLogonCount
            $row["Last Bad Password"] = if ($user.LastBadPasswordAttempt) { $user.LastBadPasswordAttempt.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
            $row["Last Logon"] = if ($user.LastLogonDate) { $user.LastLogonDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
            $row["Account"] = if ($user.Enabled) { "Enabled" } else { "Disabled" }
            $dataTable.Rows.Add($row)
        }
        $dataGridView.DataSource = $dataTable
        
        # Clear selection (don't select first row by default)
        $dataGridView.ClearSelection()
        $detailTextBox.Text = ""
    }
}

function Export-ToCSV {
    if ($script:lockedUsers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveDialog.FileName = "LockedUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $script:lockedUsers | Select-Object SamAccountName,DisplayName,Department,Title,EmailAddress,@{N='LockoutTime';E={if ($_.LockoutTime) { $_.LockoutTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }}},LockoutDuration,BadLogonCount,@{N='LastBadPasswordAttempt';E={if ($_.LastBadPasswordAttempt) { $_.LastBadPasswordAttempt.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }}},@{N='LastLogonDate';E={if ($_.LastLogonDate) { $_.LastLogonDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }}},Enabled,Description,DistinguishedName | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Export successful!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Export failed: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

function Search-Users {
    if ($script:isClosing) { return }
    $searchText = $searchBox.Text.Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($searchText)) { 
        Update-DataGridView $script:filteredUsers 
    }
    else {
        $searched = $script:filteredUsers | Where-Object { $_.SamAccountName -like "*$searchText*" -or $_.DisplayName -like "*$searchText*" -or $_.Department -like "*$searchText*" -or $_.EmailAddress -like "*$searchText*" }
        Update-DataGridView $searched
    }
}

function Show-UserDetails {
    if ($script:isClosing) { return }
    if ($dataGridView.SelectedRows.Count -gt 0) {
        $selectedUsername = $dataGridView.SelectedRows[0].Cells["Username"].Value
        $user = $script:lockedUsers | Where-Object { $_.SamAccountName -eq $selectedUsername }
        if ($user) {
            $details = "Username: $($user.SamAccountName)`r`nDisplay Name: $($user.DisplayName)`r`nEmail: $($user.EmailAddress)`r`nDepartment: $($user.Department)`r`nTitle: $($user.Title)`r`n`r`n"
            $details += "Lockout Time: $(if ($user.LockoutTime) { $user.LockoutTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' })`r`n"
            $details += "Lockout Duration: $($user.LockoutDuration)`r`nBad Logon Count: $($user.BadLogonCount)`r`n"
            $details += "Last Bad Password: $(if ($user.LastBadPasswordAttempt) { $user.LastBadPasswordAttempt.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' })`r`n"
            $details += "Last Logon: $(if ($user.LastLogonDate) { $user.LastLogonDate.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Never' })`r`n"
            $details += "Account Status: $(if ($user.Enabled) { 'Enabled' } else { 'Disabled' })`r`n`r`nDN: $($user.DistinguishedName)"
            $detailTextBox.Text = $details
        }
    }
}

# Event handlers
$autoRefreshCombo.Add_SelectedIndexChanged({
    $autoRefreshTimer.Stop(); $countdownTimer.Stop(); $script:nextRefreshTime = $null; $nextRefreshLabel.Text = ""
    $intervalMinutes = switch ($autoRefreshCombo.SelectedIndex) { 1 { 1 }; 2 { 2 }; 3 { 3 }; 4 { 5 }; 5 { 10 }; 6 { 15 }; default { 0 } }
    if ($intervalMinutes -gt 0) {
        $autoRefreshTimer.Interval = $intervalMinutes * 60 * 1000
        $script:nextRefreshTime = (Get-Date).AddMinutes($intervalMinutes)
        $autoRefreshTimer.Start(); $countdownTimer.Start()
        Update-Status "Auto-refresh: $intervalMinutes min" "Blue"
    } else { Update-Status "Ready" "Gray" }
})

# Filter change events
$filterAccountCombo.Add_SelectedIndexChanged({
    if ($script:lockedUsers.Count -gt 0) {
        Apply-FilterAndRefresh
    }
})

$filterLastLogonCombo.Add_SelectedIndexChanged({
    if ($script:lockedUsers.Count -gt 0) {
        Apply-FilterAndRefresh
    }
})

# Auto-unlock settings change handlers - immediately update queue if enabled
$autoUnlockDelayCombo.Add_SelectedIndexChanged({
    if ($autoUnlockCheckbox.Checked -and $script:filteredUsers.Count -gt 0) {
        Update-AutoUnlockQueue -users $script:filteredUsers
    }
})

$lastLogonDaysCombo.Add_SelectedIndexChanged({
    if ($autoUnlockCheckbox.Checked -and $script:filteredUsers.Count -gt 0) {
        Update-AutoUnlockQueue -users $script:filteredUsers
    }
})

$badLogonCountCombo.Add_SelectedIndexChanged({
    if ($autoUnlockCheckbox.Checked -and $script:filteredUsers.Count -gt 0) {
        Update-AutoUnlockQueue -users $script:filteredUsers
    }
})

$autoUnlockCheckbox.Add_CheckedChanged({
    if ($autoUnlockCheckbox.Checked) {
        $autoUnlockTimer.Start()
        $countdownTimer.Start()
        $autoUnlockQueueLabel.Visible = $true
        if ($script:filteredUsers.Count -gt 0) {
            Update-AutoUnlockQueue -users $script:filteredUsers
        }
        Update-Status "Auto-unlock ON" "Blue"
    } else {
        $autoUnlockTimer.Stop()
        $script:autoUnlockQueue.Clear()
        $script:notifiedUpcoming.Clear()
        $autoUnlockStatusLabel.Text = ""
        $autoUnlockQueueLabel.Visible = $false
        Update-Status "Auto-unlock OFF" "Gray"
    }
})

$refreshButton.Add_Click({
    Get-LockedOutUsersAsync
    if ($autoRefreshCombo.SelectedIndex -gt 0) {
        $autoRefreshTimer.Stop()
        $intervalMinutes = switch ($autoRefreshCombo.SelectedIndex) { 1 { 1 }; 2 { 2 }; 3 { 3 }; 4 { 5 }; 5 { 10 }; 6 { 15 }; default { 0 } }
        $script:nextRefreshTime = (Get-Date).AddMinutes($intervalMinutes)
        $autoRefreshTimer.Start()
        # Ensure countdown timer is running
        if (-not $countdownTimer.Enabled) {
            $countdownTimer.Start()
        }
    }
})

$exportButton.Add_Click({ Export-ToCSV })
$unlockButton.Add_Click({ Unlock-SelectedUser })

$queryLockoutSourceButton.Add_Click({
    if ($dataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a user from the table first", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $username = $dataGridView.SelectedRows[0].Cells["Username"].Value
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This operation will query all Domain Controllers for lockout source of user '$username' and may be SLOW.`n`nDo you want to continue?",
        "Warning - Slow Operation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Get-LockoutSourceAsync -Username $username
    }
})
$unlockMenuItem.Add_Click({ Unlock-SelectedUser })
$cancelButton.Add_Click({ $script:cancelRequested = $true; Update-Status "Cancelling..." "Orange"; Cleanup-AsyncResources; Set-UIState -IsQuerying $false })

$testSoundButton.Add_Click({ 
    Play-AlertSound
    if ($ttsCheckbox.Checked) { 
        Speak-Notification -Message "This is a test notification. TTS is working correctly."
    }
})

$queryAllDCButton.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This operation will query all Domain Controllers for locked users and may be SLOW.`n`nDo you want to continue?",
        "Warning - Slow Operation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Get-LockedOutUsersFromAllDCsAsync
    }
})

$queryLockoutSourceMenuItem.Add_Click({
    if ($dataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a user from the table first", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $username = $dataGridView.SelectedRows[0].Cells["Username"].Value
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This operation will query all Domain Controllers for lockout source of user '$username' and may be SLOW.`n`nDo you want to continue?",
        "Warning - Slow Operation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Get-LockoutSourceAsync -Username $username
    }
})

$searchBox.Add_TextChanged({ Search-Users })
$dataGridView.Add_SelectionChanged({ Show-UserDetails })

$copySelectedButton.Add_Click({
    if ($detailTextBox.SelectionLength -gt 0) {
        [System.Windows.Forms.Clipboard]::SetText($detailTextBox.SelectedText)
        Update-Status "Copied" "Green"
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select some text first", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$copyAllButton.Add_Click({
    if ($detailTextBox.Text.Length -gt 0) {
        [System.Windows.Forms.Clipboard]::SetText($detailTextBox.Text)
        Update-Status "Copied all" "Green"
    } else {
        [System.Windows.Forms.MessageBox]::Show("No details to copy", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$form.Add_Shown({
    # Ensure auto-unlock controls are disabled until DC info is loaded
    $autoUnlockCheckbox.Enabled = $false
    $autoUnlockDelayCombo.Enabled = $false
    $lastLogonDaysCombo.Enabled = $false
    $badLogonCountCombo.Enabled = $false
    $ttsUpcomingUnlockCheckbox.Enabled = $false
    $ttsUnlockedCheckbox.Enabled = $false
    $script:dcDataLoaded = $false
    Get-DomainInfoAsync
})

$form.Add_FormClosing({
    $script:isClosing = $true
    
    try { $timer.Stop() } catch { }
    try { $autoRefreshTimer.Stop() } catch { }
    try { $countdownTimer.Stop() } catch { }
    try { $autoUnlockTimer.Stop() } catch { }
    
    # Clean up modern TTS resources
    try {
        if ($script:mediaPlayer) {
            $script:mediaPlayer.Dispose()
            $script:mediaPlayer = $null
        }
        if ($script:modernSynth) {
            $script:modernSynth.Dispose()
            $script:modernSynth = $null
        }
    } catch { }

    # Clean up legacy TTS resources
    try {
        if ($script:synthesizer) {
            $script:synthesizer.SpeakAsyncCancelAll()
            $script:synthesizer.Dispose()
            $script:synthesizer = $null
        }
    } catch { }
    
    Cleanup-AsyncResources
    
    try { $timer.Dispose() } catch { }
    try { $autoRefreshTimer.Dispose() } catch { }
    try { $countdownTimer.Dispose() } catch { }
    try { $autoUnlockTimer.Dispose() } catch { }
})

[void]$form.ShowDialog()
