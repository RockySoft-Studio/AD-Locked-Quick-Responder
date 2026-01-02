# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

AD Locked - Quick Responder v2.3.0 is a PowerShell-based Windows Forms GUI application for IT administrators to monitor and manage Active Directory locked user accounts. Built for Thundermist Health Center IT Team.

**GitHub**: https://github.com/RockySoft-Studio/AD-Locked-Quick-Responder

## Requirements

- PowerShell 5.1+
- Active Directory PowerShell module (`#Requires -Modules ActiveDirectory`)
- Must run on a domain controller or machine with RSAT tools installed
- Windows 10/11 (uses System.Windows.Forms, System.Drawing, System.Speech)
- Windows 11 Natural Voices (optional, auto-fallback to legacy TTS)

## Running the Application

```powershell
# Run directly
.\AD Locked - Quick Responder.ps1

# Or with PowerShell explicitly
powershell -ExecutionPolicy Bypass -File ".\AD Locked - Quick Responder.ps1"
```

## Architecture

### Single-File Application
The entire application is contained in one PowerShell script using Windows Forms for the GUI.

### Key Components

**Async Query Pattern**: Uses PowerShell runspaces for non-blocking AD queries:
- `$script:runspace` - Background execution context
- `$script:powershell` - PowerShell instance for async operations
- `$script:asyncResult` - Tracks async completion
- Timer-based polling (`$timer.Add_Tick`) checks async completion and updates UI

**Main Async Functions**:
- `Get-DomainInfoAsync` - Loads domain/PDC info on startup
- `Get-LockedOutUsersAsync` - Queries locked accounts from main DC (fast)
- `Get-LockedOutUsersFromAllDCsAsync` - Queries locked accounts from all DCs (slow)
- `Get-LockoutSourceAsync` - Queries all DCs for lockout source analysis of a specific user

**Auto-Unlock System**:
- `$script:autoUnlockQueue` - Hashtable tracking users pending auto-unlock
- `$script:notifiedUpcoming` - Tracks users already notified of upcoming unlock
- `Process-AutoUnlockQueue` - Executes unlocks when delay timer expires
- Settings changes take effect immediately when auto-unlock is enabled
- Controls disabled until DC info is loaded (`$script:dcDataLoaded`)

**TTS Notification System** (v2.1.2):
- Windows 11: Uses `Windows.Media.SpeechSynthesis` for natural voices (Jenny, Guy)
- Windows 10: Falls back to `System.Speech.Synthesis.SpeechSynthesizer`
- `$script:useModernTTS` - Flag indicating which TTS API is active
- `$script:modernSynth` / `$script:mediaPlayer` - Modern TTS components
- `$script:synthesizer` - Legacy TTS fallback
- `Speak-Notification` - Unified TTS function with auto-detection

**Timer Behavior**:
- Auto-refresh timer pauses during queries (shows "Paused...")
- Resumes with fresh interval after query completes

### Global State Variables

```powershell
$script:lockedUsers        # All locked users from AD
$script:filteredUsers      # Users after applying filters
$script:previousLockedUsers # For detecting new lockouts
$script:newlyLockedUsers   # Usernames to highlight in yellow
$script:autoUnlockQueue    # Pending auto-unlock entries
$script:dcDataLoaded       # True after DC info loaded (enables auto-unlock controls)
$script:useModernTTS       # True if Windows 11 Natural Voice API available
```

### Filter Chain
`Check-NewLockedUsers` → `Apply-FilterAndRefresh` → `Get-FilteredUsers` → `Update-DataGridView`

## UI Layout (Top to Bottom)

1. Header with logo, title, and Info button (i) for changelog
2. Query Controls (Refresh from Main DC, Refresh from All DCs, Cancel, Search, Export)
3. Notification Settings (Auto-refresh interval, Sound/TTS toggles, TTS content options)
4. Auto-Unlock Settings (Delay, Last logon limit, Bad logon threshold, TTS checkboxes)
5. Queue display area (shows pending auto-unlock users when active)
6. Locked Users section with Filter Controls (Account status, Last logon)
7. DataGridView (Locked users table) - Right-click menu: Unlock, Query lockout source
8. Unlock/Query buttons (above details panel)
9. Selected User Details panel
10. TTS Log panel
11. Domain/Status panel at bottom

## Auto-Unlock Filtering Logic

Auto-unlock only processes users that pass through **two layers of filtering**:

### Layer 1: Display Filter
Applied via `Get-FilteredUsers` using:
- `filterAccountCombo` - Account status (All/Enabled/Disabled)
- `filterLastLogonCombo` - Last logon within (All/7/15/30/60/90 days)

### Layer 2: Auto-Unlock Specific Criteria
Applied in `Update-AutoUnlockQueue`:
- `$user.Enabled -eq $true` - Must be enabled account
- `$user.BadLogonCount -ge $badLogonThreshold` - Bad logon count meets threshold (`badLogonCountCombo`)
- `$user.LastLogonDate -ge $lastLogonCutoff` - Last logon within limit (`lastLogonDaysCombo`)

**Last Logon Options**: 1, 2, 3, 5, 7, 10, 15, 30, 60 days, or No limit

### Flow
```
$script:lockedUsers (all locked)
    ↓ Get-FilteredUsers (Layer 1)
$script:filteredUsers (displayed in grid)
    ↓ Update-AutoUnlockQueue (Layer 2)
$script:autoUnlockQueue (pending auto-unlock)
```

### Time Calculation
Last logon cutoff uses calendar days (midnight), not exact time:
```powershell
$lastLogonCutoff = (Get-Date).Date.AddDays(-$lastLogonDaysLimit)
```
This means "10 days" = any time within the last 10 calendar days.

### Note on BadLogonCount
`BadLogonCount = 0` does NOT mean the user is unlocked. The `Search-ADAccount -LockedOut` cmdlet guarantees all returned users have `LockedOut = True`. BadLogonCount resets independently based on domain policy timer, while lock status persists until manual unlock or lockout duration expires.

## Query Operations

| Operation | Trigger | Description |
|-----------|---------|-------------|
| Refresh from Main DC | Button | Fast query from PDC only |
| Refresh from All DCs | Button (with warning) | Slow query from all DCs |
| Query lockout source | Button/Right-click (with warning) | Query lockout events for selected user from ALL DCs |
| Unlock user | Button/Right-click menu | Unlock selected user |

## AD Cmdlets Used

- `Search-ADAccount -LockedOut` - Find locked accounts
- `Get-ADUser` - Get user properties
- `Get-ADDomain` - Get domain info
- `Get-ADDomainController` - List all DCs
- `Unlock-ADAccount` - Unlock user

## Version History

### v2.3.0 (Current)
- Detailed query progress display with batch info
  - Shows "Searching locked users..." during initial search
  - Shows "2/5 T-45s (00:12)" during batch processing
- Progress display moved to left of progress bar in status panel
- Scrollable changelog window

### v2.2.8
- Added elapsed time display during queries
- Elapsed time updates every 2.5 seconds
- Batch parallel query: Main DC query processes 5 users at a time (safer for server)
- Added 60-second timeout per batch to prevent hanging
- Batch must complete or terminate before next batch starts (prevents server overload)
- Fixed: Single locked user now correctly detected (array conversion fix)
- Added permission detection: Unlock controls disabled when user lacks privileges
- Fixed: Windows 11 TTS now runs async in background (no more UI freeze)
- Parallel DC queries - significantly faster "Query from All DCs"
- Removed event log queries (Get-WinEvent) - simplified DC status display
- Fixed: All queries now explicitly target PDC for consistent BadLogonCount
- Unlock operations also target PDC for consistency

### v2.1.2
- Windows 11 Natural Voice TTS support (auto fallback to legacy on Win10)
- Info button (i) for viewing changelog
- Removed Stop button - use checkbox to stop auto-unlock
- Fixed auto-unlock controls clickable before DC info loaded
- UI spacing adjustments

### v2.1.1
- Added slow operation warning for "Query from ALL DCs"
- Moved Unlock/Query buttons above Selected User Details
- TTS checkboxes moved to same row as auto-unlock controls

### v2.1.0
- Renamed query/TTS labels for clarity
- Query status bar moved to bottom domain panel

### v2.0.0
- Major UI layout reorganization
- Enhanced auto-unlock queue display area

### v1.5.0 (Base)
- All core features: async queries, auto-refresh, notifications, auto-unlock, filters, export
