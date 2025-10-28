#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Disables AI features in Windows 11
.DESCRIPTION
    This script checks and disables various AI features in Windows 11 including Copilot,
    Recall, and related privacy settings. Must be run as Administrator.
.NOTES
    Author: Generated Script
    Date: 2025-10-28
    Requires: Windows 11, Administrator privileges
#>

Write-Host "=== Windows 11 AI Feature Disabler ===" -ForegroundColor Cyan
Write-Host "Checking and disabling AI features...`n" -ForegroundColor Cyan

# Function to set registry value and report status
function Set-RegistryValueWithStatus {
    param(
        [string]$Path,
        [string]$Name,
        [object]$Value,
        [string]$Type,
        [string]$Description
    )
    
    try {
        # Create the key if it doesn't exist
        if (!(Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        
        # Check current value
        $currentValue = $null
        try {
            $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $Name
        } catch {}
        
        if ($currentValue -eq $Value) {
            Write-Host "[OK] $Description - Already disabled" -ForegroundColor Green
        } else {
            Write-Host "[CHANGING] $Description" -ForegroundColor Yellow
            Write-Host "  Current: $currentValue -> New: $Value" -ForegroundColor Gray
            Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -Force
            Write-Host "[DONE] $Description - Disabled" -ForegroundColor Green
        }
    } catch {
        Write-Host "[ERROR] $Description - Failed: $_" -ForegroundColor Red
    }
}

# 1. Disable Windows Copilot
Write-Host "`n--- Windows Copilot ---" -ForegroundColor Yellow
Set-RegistryValueWithStatus `
    -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
    -Name "ShowCopilotButton" `
    -Value 0 `
    -Type "DWord" `
    -Description "Windows Copilot taskbar button"

Set-RegistryValueWithStatus `
    -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" `
    -Name "TurnOffWindowsCopilot" `
    -Value 1 `
    -Type "DWord" `
    -Description "Windows Copilot system-wide"

# 2. Disable Recall (Snapshots)
Write-Host "`n--- Windows Recall / Snapshots ---" -ForegroundColor Yellow
Set-RegistryValueWithStatus `
    -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" `
    -Name "SubscribedContent-353698Enabled" `
    -Value 0 `
    -Type "DWord" `
    -Description "Windows Recall suggestions"

Set-RegistryValueWithStatus `
    -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" `
    -Name "DisableAIDataAnalysis" `
    -Value 1 `
    -Type "DWord" `
    -Description "AI Data Analysis"

# 3. Disable AI-powered suggestions and recommendations
Write-Host "`n--- AI Suggestions & Recommendations ---" -ForegroundColor Yellow
Set-RegistryValueWithStatus `
    -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
    -Name "ShowSyncProviderNotifications" `
    -Value 0 `
    -Type "DWord" `
    -Description "Sync provider notifications"

Set-RegistryValueWithStatus `
    -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" `
    -Name "SubscribedContent-338393Enabled" `
    -Value 0 `
    -Type "DWord" `
    -Description "Suggested content in Settings"

Set-RegistryValueWithStatus `
    -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" `
    -Name "SubscribedContent-353694Enabled" `
    -Value 0 `
    -Type "DWord" `
    -Description "Suggested content in Start"

# 4. Disable personalized ads and tracking
Write-Host "`n--- Privacy & Tracking ---" -ForegroundColor Yellow
Set-RegistryValueWithStatus `
    -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" `
    -Name "Enabled" `
    -Value 0 `
    -Type "DWord" `
    -Description "Advertising ID"

Set-RegistryValueWithStatus `
    -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo" `
    -Name "DisabledByGroupPolicy" `
    -Value 1 `
    -Type "DWord" `
    -Description "Advertising ID (Group Policy)"

# 5. Set diagnostic data to minimum
Write-Host "`n--- Diagnostic Data ---" -ForegroundColor Yellow
Set-RegistryValueWithStatus `
    -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" `
    -Name "AllowTelemetry" `
    -Value 0 `
    -Type "DWord" `
    -Description "Telemetry collection (0=Security only)"

Set-RegistryValueWithStatus `
    -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection" `
    -Name "AllowTelemetry" `
    -Value 0 `
    -Type "DWord" `
    -Description "Telemetry collection (secondary)"

# 6. Disable online tips and suggestions
Write-Host "`n--- Online Tips & Help ---" -ForegroundColor Yellow
Set-RegistryValueWithStatus `
    -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" `
    -Name "AllowOnlineTips" `
    -Value 0 `
    -Type "DWord" `
    -Description "Online tips"

# 7. Disable typing insights
Write-Host "`n--- Typing & Input ---" -ForegroundColor Yellow
Set-RegistryValueWithStatus `
    -Path "HKCU:\Software\Microsoft\Input\Settings" `
    -Name "InsightsEnabled" `
    -Value 0 `
    -Type "DWord" `
    -Description "Typing insights"

Set-RegistryValueWithStatus `
    -Path "HKCU:\Software\Microsoft\TabletTip\1.7" `
    -Name "EnableAutocorrection" `
    -Value 0 `
    -Type "DWord" `
    -Description "Text prediction"

# 8. Disable Taskbar Search Box
Write-Host "`n--- Taskbar Search ---" -ForegroundColor Yellow
Set-RegistryValueWithStatus `
    -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" `
    -Name "SearchboxTaskbarMode" `
    -Value 0 `
    -Type "DWord" `
    -Description "Taskbar search box (0=Hidden, 1=Icon only, 2=Box)"

# 9. Move Start Menu to Left
Write-Host "`n--- Start Menu Position ---" -ForegroundColor Yellow
Set-RegistryValueWithStatus `
    -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
    -Name "TaskbarAl" `
    -Value 0 `
    -Type "DWord" `
    -Description "Start menu alignment (0=Left, 1=Center)"

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "AI features and taskbar settings have been configured:" -ForegroundColor White
Write-Host "  - AI features disabled" -ForegroundColor Gray
Write-Host "  - Taskbar search hidden" -ForegroundColor Gray
Write-Host "  - Start menu moved to left" -ForegroundColor Gray
Write-Host "`nChanges require:" -ForegroundColor White
Write-Host "  - Restarting Windows Explorer (or reboot)" -ForegroundColor Gray
Write-Host "`nTo restart Explorer now, run: Stop-Process -Name explorer -Force" -ForegroundColor Yellow
