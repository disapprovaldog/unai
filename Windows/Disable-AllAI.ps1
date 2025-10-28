#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Disables AI features in Windows 11 and Microsoft Office
.DESCRIPTION
    This master script runs both the Windows 11 and Office AI disablers.
    It checks current status and disables various AI features including Copilot,
    Recall, Editor, and connected experiences.
.NOTES
    Author: Generated Script
    Date: 2025-10-28
    Requires: Windows 11, Administrator privileges
.PARAMETER SkipWindows
    Skip Windows 11 AI feature disabling
.PARAMETER SkipOffice
    Skip Office AI feature disabling
#>

param(
    [switch]$SkipWindows,
    [switch]$SkipOffice
)

Write-Host @"
╔═══════════════════════════════════════════════════════════════╗
║                                                               ║
║          Windows 11 & Office AI Feature Disabler             ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan

Write-Host "`nThis script will disable AI features and check their current status." -ForegroundColor White
Write-Host "Administrator privileges are required.`n" -ForegroundColor Yellow

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[ERROR] This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    exit 1
}

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
        if (!(Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        
        $currentValue = $null
        try {
            $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $Name
        } catch {}
        
        if ($currentValue -eq $Value) {
            Write-Host "  [✓] $Description - Already disabled" -ForegroundColor Green
        } else {
            Write-Host "  [~] $Description - Changing..." -ForegroundColor Yellow
            Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -Force
            Write-Host "  [✓] $Description - Now disabled" -ForegroundColor Green
        }
    } catch {
        Write-Host "  [✗] $Description - Failed: $_" -ForegroundColor Red
    }
}

# ============================================================================
# WINDOWS 11 AI FEATURES
# ============================================================================
if (-not $SkipWindows) {
    Write-Host "`n" + ("="*70) -ForegroundColor Cyan
    Write-Host " WINDOWS 11 AI FEATURES" -ForegroundColor Cyan
    Write-Host ("="*70) -ForegroundColor Cyan

    Write-Host "`n[1/8] Windows Copilot" -ForegroundColor Yellow
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
        -Name "ShowCopilotButton" -Value 0 -Type "DWord" `
        -Description "Copilot taskbar button"
    
    Set-RegistryValueWithStatus `
        -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" `
        -Name "TurnOffWindowsCopilot" -Value 1 -Type "DWord" `
        -Description "Copilot system policy"

    Write-Host "`n[2/8] Windows Recall / Snapshots" -ForegroundColor Yellow
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" `
        -Name "SubscribedContent-353698Enabled" -Value 0 -Type "DWord" `
        -Description "Recall suggestions"
    
    Set-RegistryValueWithStatus `
        -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" `
        -Name "DisableAIDataAnalysis" -Value 1 -Type "DWord" `
        -Description "AI data analysis"

    Write-Host "`n[3/8] AI Suggestions & Recommendations" -ForegroundColor Yellow
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
        -Name "ShowSyncProviderNotifications" -Value 0 -Type "DWord" `
        -Description "Sync provider notifications"
    
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" `
        -Name "SubscribedContent-338393Enabled" -Value 0 -Type "DWord" `
        -Description "Settings suggestions"
    
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" `
        -Name "SubscribedContent-353694Enabled" -Value 0 -Type "DWord" `
        -Description "Start menu suggestions"

    Write-Host "`n[4/8] Privacy & Advertising" -ForegroundColor Yellow
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" `
        -Name "Enabled" -Value 0 -Type "DWord" `
        -Description "Advertising ID"
    
    Set-RegistryValueWithStatus `
        -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo" `
        -Name "DisabledByGroupPolicy" -Value 1 -Type "DWord" `
        -Description "Advertising ID policy"

    Write-Host "`n[5/8] Diagnostic Data & Telemetry" -ForegroundColor Yellow
    Set-RegistryValueWithStatus `
        -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" `
        -Name "AllowTelemetry" -Value 0 -Type "DWord" `
        -Description "Telemetry (security only)"
    
    Set-RegistryValueWithStatus `
        -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection" `
        -Name "AllowTelemetry" -Value 0 -Type "DWord" `
        -Description "Telemetry secondary"

    Write-Host "`n[6/8] Online Tips & Help" -ForegroundColor Yellow
    Set-RegistryValueWithStatus `
        -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" `
        -Name "AllowOnlineTips" -Value 0 -Type "DWord" `
        -Description "Online tips"

    Write-Host "`n[7/8] Typing Insights" -ForegroundColor Yellow
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Input\Settings" `
        -Name "InsightsEnabled" -Value 0 -Type "DWord" `
        -Description "Typing insights"
    
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\TabletTip\1.7" `
        -Name "EnableAutocorrection" -Value 0 -Type "DWord" `
        -Description "Text prediction"

    Write-Host "`n[8/10] Windows Search Highlights" -ForegroundColor Yellow
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\SearchSettings" `
        -Name "IsDeviceSearchHistoryEnabled" -Value 0 -Type "DWord" `
        -Description "Search history"

    Write-Host "`n[9/10] Taskbar Search" -ForegroundColor Yellow
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" `
        -Name "SearchboxTaskbarMode" -Value 0 -Type "DWord" `
        -Description "Taskbar search box (0=Hidden)"

    Write-Host "`n[10/10] Start Menu Position" -ForegroundColor Yellow
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" `
        -Name "TaskbarAl" -Value 0 -Type "DWord" `
        -Description "Start menu alignment (0=Left)"
}

# ============================================================================
# MICROSOFT OFFICE AI FEATURES
# ============================================================================
if (-not $SkipOffice) {
    Write-Host "`n" + ("="*70) -ForegroundColor Cyan
    Write-Host " MICROSOFT OFFICE AI FEATURES" -ForegroundColor Cyan
    Write-Host ("="*70) -ForegroundColor Cyan

    $officeVersions = @("16.0")  # Office 2016/2019/2021/365

    Write-Host "`n[1/8] Microsoft 365 Copilot" -ForegroundColor Yellow
    foreach ($version in $officeVersions) {
        Set-RegistryValueWithStatus `
            -Path "HKCU:\Software\Policies\Microsoft\office\$version\common\copilot" `
            -Name "PolicyEnabled" -Value 0 -Type "DWord" `
            -Description "Copilot user policy (v$version)"
        
        Set-RegistryValueWithStatus `
            -Path "HKLM:\SOFTWARE\Policies\Microsoft\office\$version\common\copilot" `
            -Name "PolicyEnabled" -Value 0 -Type "DWord" `
            -Description "Copilot system policy (v$version)"
    }

    Write-Host "`n[2/8] Connected Experiences" -ForegroundColor Yellow
    foreach ($version in $officeVersions) {
        Set-RegistryValueWithStatus `
            -Path "HKCU:\Software\Policies\Microsoft\office\$version\common\privacy" `
            -Name "disconnectedstate" -Value 2 -Type "DWord" `
            -Description "Connected experiences (v$version)"
        
        Set-RegistryValueWithStatus `
            -Path "HKCU:\Software\Policies\Microsoft\office\$version\common\privacy" `
            -Name "usercontentdisabled" -Value 2 -Type "DWord" `
            -Description "Content analysis (v$version)"
    }

    Write-Host "`n[3/8] Editor & Writing Assistance" -ForegroundColor Yellow
    foreach ($version in $officeVersions) {
        Set-RegistryValueWithStatus `
            -Path "HKCU:\Software\Microsoft\Office\$version\Common\General" `
            -Name "ServiceLevelOptions" -Value 0 -Type "DWord" `
            -Description "Intelligent services (v$version)"
        
        Set-RegistryValueWithStatus `
            -Path "HKCU:\Software\Microsoft\Office\$version\Common\MailSettings" `
            -Name "EnableTextPrediction" -Value 0 -Type "DWord" `
            -Description "Text predictions (v$version)"
    }

    Write-Host "`n[4/8] Insights & Research" -ForegroundColor Yellow
    foreach ($version in $officeVersions) {
        Set-RegistryValueWithStatus `
            -Path "HKCU:\Software\Microsoft\Office\$version\Common\Research\Options" `
            -Name "EnableInsights" -Value 0 -Type "DWord" `
            -Description "Insights (v$version)"
    }

    Write-Host "`n[5/8] PowerPoint Designer" -ForegroundColor Yellow
    foreach ($version in $officeVersions) {
        Set-RegistryValueWithStatus `
            -Path "HKCU:\Software\Microsoft\Office\$version\PowerPoint\Options" `
            -Name "DisableDesigner" -Value 1 -Type "DWord" `
            -Description "Designer (v$version)"
    }

    Write-Host "`n[6/8] Excel Ideas" -ForegroundColor Yellow
    foreach ($version in $officeVersions) {
        Set-RegistryValueWithStatus `
            -Path "HKCU:\Software\Microsoft\Office\$version\Excel\Options" `
            -Name "DisableIdeas" -Value 1 -Type "DWord" `
            -Description "Ideas (v$version)"
    }

    Write-Host "`n[7/8] Online Content" -ForegroundColor Yellow
    foreach ($version in $officeVersions) {
        Set-RegistryValueWithStatus `
            -Path "HKCU:\Software\Policies\Microsoft\office\$version\common\internet" `
            -Name "useonlinecontent" -Value 0 -Type "DWord" `
            -Description "Online content (v$version)"
    }

    Write-Host "`n[8/8] Telemetry & Feedback" -ForegroundColor Yellow
    foreach ($version in $officeVersions) {
        Set-RegistryValueWithStatus `
            -Path "HKCU:\Software\Policies\Microsoft\office\$version\common" `
            -Name "sendtelemetry" -Value 3 -Type "DWord" `
            -Description "Telemetry (v$version)"
        
        Set-RegistryValueWithStatus `
            -Path "HKCU:\Software\Policies\Microsoft\office\$version\common\feedback" `
            -Name "enabled" -Value 0 -Type "DWord" `
            -Description "Feedback (v$version)"
    }
}

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host "`n" + ("="*70) -ForegroundColor Cyan
Write-Host " COMPLETED" -ForegroundColor Cyan
Write-Host ("="*70) -ForegroundColor Cyan

Write-Host "`nAI features have been disabled. To apply all changes:" -ForegroundColor White
Write-Host "  1. Sign out and sign back in (recommended)" -ForegroundColor Gray
Write-Host "  2. Or restart your computer" -ForegroundColor Gray
Write-Host "  3. Close and reopen Office applications" -ForegroundColor Gray

Write-Host "`nTo restart Explorer now (applies some Windows changes):" -ForegroundColor Yellow
Write-Host "  Stop-Process -Name explorer -Force" -ForegroundColor Cyan

Write-Host "`nNotes:" -ForegroundColor Yellow
Write-Host "  • Some features require specific Windows/Office editions" -ForegroundColor Gray
Write-Host "  • Enterprise environments may override these settings" -ForegroundColor Gray
Write-Host "  • Future updates might re-enable some features" -ForegroundColor Gray
Write-Host ""
