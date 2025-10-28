#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Disables AI features in Microsoft Office/Microsoft 365
.DESCRIPTION
    This script checks and disables various AI features in Office applications including
    Copilot, Editor suggestions, and connected experiences. Must be run as Administrator.
.NOTES
    Author: Generated Script
    Date: 2025-10-28
    Requires: Administrator privileges
    Note: Some settings apply per-user, others system-wide
#>

Write-Host "=== Microsoft Office AI Feature Disabler ===" -ForegroundColor Cyan
Write-Host "Checking and disabling AI features in Office...`n" -ForegroundColor Cyan

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

# Detect Office version
$officeVersions = @("16.0", "15.0", "14.0")  # Office 2016/2019/2021/365, 2013, 2010
$officeApps = @("Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher", "Visio")

# 1. Disable Microsoft 365 Copilot
Write-Host "`n--- Microsoft 365 Copilot ---" -ForegroundColor Yellow
foreach ($version in $officeVersions) {
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Policies\Microsoft\office\$version\common\copilot" `
        -Name "PolicyEnabled" `
        -Value 0 `
        -Type "DWord" `
        -Description "Copilot (Office $version)"
    
    Set-RegistryValueWithStatus `
        -Path "HKLM:\SOFTWARE\Policies\Microsoft\office\$version\common\copilot" `
        -Name "PolicyEnabled" `
        -Value 0 `
        -Type "DWord" `
        -Description "Copilot system-wide (Office $version)"
}

# 2. Disable Connected Experiences (cloud-powered AI)
Write-Host "`n--- Connected Experiences ---" -ForegroundColor Yellow
foreach ($version in $officeVersions) {
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Policies\Microsoft\office\$version\common\privacy" `
        -Name "disconnectedstate" `
        -Value 2 `
        -Type "DWord" `
        -Description "Disconnect all connected experiences (Office $version)"
    
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Policies\Microsoft\office\$version\common\privacy" `
        -Name "usercontentdisabled" `
        -Value 2 `
        -Type "DWord" `
        -Description "Disable content analysis (Office $version)"
    
    Set-RegistryValueWithStatus `
        -Path "HKLM:\SOFTWARE\Policies\Microsoft\office\$version\common\privacy" `
        -Name "disconnectedstate" `
        -Value 2 `
        -Type "DWord" `
        -Description "Connected experiences system-wide (Office $version)"
}

# 3. Disable Editor and AI writing assistance
Write-Host "`n--- Editor & Writing Assistance ---" -ForegroundColor Yellow
foreach ($version in $officeVersions) {
    # Disable Editor in Word
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Office\$version\Word\Options" `
        -Name "EnableLinguisticServices" `
        -Value 0 `
        -Type "DWord" `
        -Description "Word linguistic services (Office $version)"
    
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Office\$version\Word\Options" `
        -Name "EnableProofingServices" `
        -Value 0 `
        -Type "DWord" `
        -Description "Word proofing services (Office $version)"
    
    # Disable intelligent services
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Office\$version\Common\General" `
        -Name "ServiceLevelOptions" `
        -Value 0 `
        -Type "DWord" `
        -Description "Intelligent services (Office $version)"
    
    # Disable text predictions
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Office\$version\Common\MailSettings" `
        -Name "EnableTextPrediction" `
        -Value 0 `
        -Type "DWord" `
        -Description "Text predictions (Office $version)"
}

# 4. Disable insights and research
Write-Host "`n--- Insights & Research ---" -ForegroundColor Yellow
foreach ($version in $officeVersions) {
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Policies\Microsoft\office\$version\common\research" `
        -Name "options" `
        -Value 0 `
        -Type "DWord" `
        -Description "Research options (Office $version)"
    
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Office\$version\Common\Research\Options" `
        -Name "EnableInsights" `
        -Value 0 `
        -Type "DWord" `
        -Description "Insights (Office $version)"
}

# 5. Disable Designer in PowerPoint
Write-Host "`n--- PowerPoint Designer ---" -ForegroundColor Yellow
foreach ($version in $officeVersions) {
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Office\$version\PowerPoint\Options" `
        -Name "DisableDesigner" `
        -Value 1 `
        -Type "DWord" `
        -Description "PowerPoint Designer (Office $version)"
    
    Set-RegistryValueWithStatus `
        -Path "HKLM:\SOFTWARE\Policies\Microsoft\office\$version\powerpoint\options" `
        -Name "DisableDesigner" `
        -Value 1 `
        -Type "DWord" `
        -Description "PowerPoint Designer system-wide (Office $version)"
}

# 6. Disable Ideas in Excel
Write-Host "`n--- Excel Ideas ---" -ForegroundColor Yellow
foreach ($version in $officeVersions) {
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Microsoft\Office\$version\Excel\Options" `
        -Name "DisableIdeas" `
        -Value 1 `
        -Type "DWord" `
        -Description "Excel Ideas (Office $version)"
}

# 7. Disable online content and templates
Write-Host "`n--- Online Content ---" -ForegroundColor Yellow
foreach ($version in $officeVersions) {
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Policies\Microsoft\office\$version\common\internet" `
        -Name "useonlinecontent" `
        -Value 0 `
        -Type "DWord" `
        -Description "Online content (Office $version)"
    
    Set-RegistryValueWithStatus `
        -Path "HKLM:\SOFTWARE\Policies\Microsoft\office\$version\common\internet" `
        -Name "useonlinecontent" `
        -Value 0 `
        -Type "DWord" `
        -Description "Online content system-wide (Office $version)"
}

# 8. Disable telemetry and feedback
Write-Host "`n--- Telemetry & Feedback ---" -ForegroundColor Yellow
foreach ($version in $officeVersions) {
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Policies\Microsoft\office\$version\common" `
        -Name "sendtelemetry" `
        -Value 3 `
        -Type "DWord" `
        -Description "Telemetry - Neither (Office $version)"
    
    Set-RegistryValueWithStatus `
        -Path "HKCU:\Software\Policies\Microsoft\office\$version\common\feedback" `
        -Name "enabled" `
        -Value 0 `
        -Type "DWord" `
        -Description "Feedback collection (Office $version)"
    
    Set-RegistryValueWithStatus `
        -Path "HKLM:\SOFTWARE\Policies\Microsoft\office\$version\common" `
        -Name "sendtelemetry" `
        -Value 3 `
        -Type "DWord" `
        -Description "Telemetry system-wide (Office $version)"
}

# 9. Disable automatic updates that might re-enable features
Write-Host "`n--- Update Settings ---" -ForegroundColor Yellow
foreach ($version in $officeVersions) {
    Set-RegistryValueWithStatus `
        -Path "HKLM:\SOFTWARE\Policies\Microsoft\office\$version\common\officeupdate" `
        -Name "enableautomaticupdates" `
        -Value 0 `
        -Type "DWord" `
        -Description "Automatic updates (Office $version)"
}

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "Office AI features have been disabled. Changes will take effect:" -ForegroundColor White
Write-Host "  - Immediately for new Office app instances" -ForegroundColor Gray
Write-Host "  - After restarting any currently running Office apps" -ForegroundColor Gray
Write-Host "`nNote: Some features may require enterprise/admin licensing to fully control" -ForegroundColor Yellow
Write-Host "If you're in an organization, some settings may be managed by your IT admin" -ForegroundColor Yellow
