# unai - Windows & Office AI Feature Disabler

A collection of PowerShell scripts to disable AI features, telemetry, and privacy-invasive settings in Windows 11 and Microsoft Office/365.

## Scripts

### `Windows/Disable-AllAI.ps1`
Master script that disables AI features in both Windows 11 and Microsoft Office.

**Features disabled:**
- Windows 11: Copilot, Recall/Snapshots, AI suggestions, advertising ID, telemetry, online tips, typing insights
- Office: Copilot, Editor, Designer (PowerPoint), Ideas (Excel), connected experiences, insights, online content, telemetry

**Usage:**
```powershell
# Disable everything (Windows + Office)
.\Windows\Disable-AllAI.ps1

# Skip Windows features, only disable Office
.\Windows\Disable-AllAI.ps1 -SkipWindows

# Skip Office features, only disable Windows
.\Windows\Disable-AllAI.ps1 -SkipOffice
```

### `Windows/Disable-Windows11AI.ps1`
Standalone script focused on Windows 11 AI features.

**Features disabled:**
- Windows Copilot (taskbar button and system-wide)
- Windows Recall / Snapshots
- AI-powered suggestions and recommendations
- Personalized ads and advertising ID
- Telemetry and diagnostic data
- Online tips and help
- Typing insights and text prediction
- Taskbar search box (hidden)
- Start menu position (moved to left)

### `Windows/Disable-OfficeAI.ps1`
Standalone script focused on Microsoft Office AI features.

**Features disabled:**
- Microsoft 365 Copilot
- Connected experiences (cloud-powered AI)
- Editor and AI writing assistance
- Insights and research features
- PowerPoint Designer
- Excel Ideas
- Online content and templates
- Telemetry and feedback
- Automatic updates

Supports multiple Office versions: 2010, 2013, 2016/2019/2021/365

## Requirements

- **Windows 11** (some features are Windows 11-specific)
- **PowerShell 5.1+**
- **Administrator privileges** (required)
- Microsoft Office (for Office-related scripts)

## Installation & Usage

1. **Clone or download** this repository
2. **Open PowerShell as Administrator**
   - Right-click on PowerShell and select "Run as Administrator"
3. **Navigate to the repository folder**
   ```powershell
   cd path\to\unai
   ```
4. **Run the desired script**
   ```powershell
   .\Windows\Disable-AllAI.ps1
   ```

## Important Notes

### Before Running
- Scripts modify Windows Registry settings
- Create a system restore point before running (recommended)
- Review the script contents to understand what changes will be made

### After Running
Changes will take effect after:
- Restarting Windows Explorer: `Stop-Process -Name explorer -Force`
- Signing out and back in (recommended)
- Restarting your computer
- Closing and reopening Office applications

### Limitations
- Some settings may require specific Windows/Office editions to take effect
- Enterprise/organization-managed devices may have Group Policy overrides
- Windows/Office updates may re-enable some features
- Some Copilot features require specific licensing and may not be affected

### Reverting Changes
These scripts disable features but do not provide automatic reverting. To re-enable features:
1. Use Windows Settings to manually re-enable desired features
2. Or restore from a system restore point created before running the scripts

## Compatibility

- **Windows:** Windows 11 (some settings may work on Windows 10)
- **Office:** Office 2010, 2013, 2016, 2019, 2021, Microsoft 365
- **Registry changes:** User-level (HKCU) and system-level (HKLM)

## Disclaimer

**USE AT YOUR OWN RISK**

- These scripts modify system and application settings
- Always create backups before making system changes
- Test in a non-production environment first
- The authors are not responsible for any issues or damage
- Scripts are provided as-is without warranty

## License

This project is provided as-is for educational and personal use.

## Contributing

Issues and pull requests are welcome. Please ensure:
- Scripts remain read-only (check status, don't modify without confirmation)
- Clear documentation of any new features
- Testing on multiple Windows/Office versions where applicable
