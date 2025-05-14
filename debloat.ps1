#Requires -RunAsAdministrator
#Requires -Version 5.1

# Windows Debloat Script
# This script removes unnecessary Windows components and optimizes system performance

# Set script execution policy
Set-ExecutionPolicy Bypass -Scope Process -Force

# Script Variables
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$configFile = Join-Path $scriptPath "config.json"
$logFile = Join-Path $scriptPath "debloat.log"
$registryBackupPath = Join-Path $scriptPath "registry_backup"

# Initialize log file
if (Test-Path $logFile) {
    try {
        Remove-Item $logFile -Force
    }
    catch {
        $logFile = Join-Path $scriptPath "debloat_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        Write-Host "Created new log file: $logFile"
    }
}

# Functions
function Write-Log {
    param($Message)
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    try {
        Add-Content -Path $logFile -Value $logMessage -ErrorAction Stop
        Write-Host $logMessage
    }
    catch {
        Write-Host "Warning: Could not write to log file: $logMessage"
    }
}

function Test-AdminPrivileges {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin) {
        Write-Log "Error: This script requires administrator privileges"
        exit 1
    }
}

function Create-RestorePoint {
    Write-Log "Creating system restore point..."
    Enable-ComputerRestore -Drive $env:SystemDrive
    Checkpoint-Computer -Description "Windows Debloat Tool - Before Changes" -RestorePointType "MODIFY_SETTINGS"
}

function Backup-RegistryKeys {
    Write-Log "Backing up registry keys..."
    
    if (-not (Test-Path $registryBackupPath)) {
        New-Item -ItemType Directory -Path $registryBackupPath | Out-Null
    }

    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies",
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection",
        "HKLM:\SYSTEM\CurrentControlSet\Services",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    )

    foreach ($path in $registryPaths) {
        $backupFile = Join-Path $registryBackupPath "$((Split-Path $path -Leaf)).reg"
        $regPath = $path.Replace("HKLM:\", "HKLM\").Replace("HKCU:\", "HKCU\")
        reg export $regPath $backupFile /y
    }
}

function Remove-BloatwareApps {
    param($apps)
    
    Write-Log "Removing bloatware applications..."
    
    foreach ($app in $apps.remove) {
        Write-Log "Removing app: $app"
        Get-AppxPackage $app -AllUsers | Remove-AppxPackage -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like $app | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
    }
}

function Disable-Services {
    param($services)
    
    Write-Log "Disabling unnecessary services..."
    
    foreach ($service in $services.disable) {
        Write-Log "Disabling service: $service"
        Set-Service -Name $service -StartupType Disabled -ErrorAction SilentlyContinue
        Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
    }
}

function Set-PrivacySettings {
    param($privacy)
    
    Write-Log "Applying privacy settings..."
    
    # Function to ensure registry path exists
    function Ensure-RegistryPath {
        param($Path)
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
            Write-Log "Created registry path: $Path"
        }
    }

    if ($privacy.disableTelemetry) {
        Write-Log "Disabling telemetry..."
        Ensure-RegistryPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection"
        Ensure-RegistryPath "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection" -Name "AllowTelemetry" -Value 0 -Type DWord -Force
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0 -Type DWord -Force
    }

    if ($privacy.disableCortana) {
        Write-Log "Disabling Cortana..."
        Ensure-RegistryPath "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name "AllowCortana" -Value 0 -Type DWord -Force
    }

    if ($privacy.disableActivityHistory) {
        Write-Log "Disabling Activity History..."
        Ensure-RegistryPath "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -Name "EnableActivityFeed" -Value 0 -Type DWord -Force
    }

    if ($privacy.disableWebSearch) {
        Write-Log "Disabling web search..."
        Ensure-RegistryPath "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name "DisableWebSearch" -Value 1 -Type DWord -Force
    }

    if ($privacy.disableAdvertisingID) {
        Write-Log "Disabling Advertising ID..."
        Ensure-RegistryPath "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo"
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo" -Name "DisabledByGroupPolicy" -Value 1 -Type DWord -Force
    }

    if ($privacy.disableErrorReporting) {
        Write-Log "Disabling Error Reporting..."
        Ensure-RegistryPath "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting"
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting" -Name "Disabled" -Value 1 -Type DWord -Force
    }

    if ($privacy.disableWindowsTips) {
        Write-Log "Disabling Windows Tips..."
        Ensure-RegistryPath "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableSoftLanding" -Value 1 -Type DWord -Force
    }
}

function Set-PerformanceSettings {
    param($performance)
    
    Write-Log "Optimizing performance settings..."
    
    if ($performance.disableSuperfetch) {
        Write-Log "Disabling Superfetch..."
        Set-Service -Name "SysMain" -StartupType Disabled
        Stop-Service -Name "SysMain" -Force
    }

    if ($performance.optimizeVisualEffects) {
        Write-Log "Optimizing visual effects..."
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" -Name "VisualFXSetting" -Value 2
    }
}

function Remove-Features {
    param($features)
    
    Write-Log "Removing Windows features..."
    
    if ($features.removeOneDrive) {
        Write-Log "Removing OneDrive..."
        Stop-Process -Name OneDrive -ErrorAction SilentlyContinue
        Start-Process "$env:SystemRoot\SysWOW64\OneDriveSetup.exe" "/uninstall" -NoNewWindow -Wait
    }

    if ($features.removeTeams) {
        Write-Log "Removing Microsoft Teams..."
        $teamsPath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams')
        if (Test-Path $teamsPath) {
            Remove-Item $teamsPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Set-TaskbarSettings {
    param($taskbar)
    
    Write-Host "Configuring taskbar settings..."
    
    # Create the registry paths if they don't exist
    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    $searchRegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
    
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    if (-not (Test-Path $searchRegistryPath)) {
        New-Item -Path $searchRegistryPath -Force | Out-Null
    }

    # Configure taskbar alignment
    if ($taskbar.alignLeft) {
        Write-Host "Setting taskbar alignment to left..."
        Set-ItemProperty -Path $registryPath -Name "TaskbarAl" -Value 0 -Type DWord -Force
    }

    # Hide Task View button
    if ($taskbar.hideTaskView) {
        Write-Host "Hiding Task View button..."
        Set-ItemProperty -Path $registryPath -Name "ShowTaskViewButton" -Value 0 -Type DWord -Force
    }

    # Configure Search icon (multiple registry keys for different Windows versions)
    if ($taskbar.hideSearch) {
        Write-Host "Hiding Search from taskbar..."
        # Windows 11 and newer versions
        Set-ItemProperty -Path $registryPath -Name "SearchboxTaskbarMode" -Value 0 -Type DWord -Force
        # Windows 10 and alternate registry locations
        Set-ItemProperty -Path $searchRegistryPath -Name "SearchboxTaskbarMode" -Value 0 -Type DWord -Force
        Set-ItemProperty -Path $registryPath -Name "ShowSearchBox" -Value 0 -Type DWord -Force
        Set-ItemProperty -Path $registryPath -Name "ShowSearch" -Value 0 -Type DWord -Force
        Set-ItemProperty -Path $registryPath -Name "SearchboxTaskbarModeCache" -Value 0 -Type DWord -Force
    }

    # Prevent This PC from opening on Explorer restart
    $explorerSettings = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Set-ItemProperty -Path $explorerSettings -Name "LaunchTo" -Value 1 -Type DWord -Force

    # Restart Explorer more cleanly
    Write-Host "Restarting Explorer to apply taskbar changes..."
    try {
        # Save current Explorer windows
        $prevExplorerWindows = (New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.LocationURL }
        
        # Kill all explorer processes
        Get-Process -Name explorer -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 2

        # Start Explorer without a window
        Start-Process explorer -WindowStyle Hidden
        Start-Sleep -Seconds 2

        # Restore previous Explorer windows if any existed
        if ($prevExplorerWindows) {
            foreach ($window in $prevExplorerWindows) {
                if ($window) {
                    (New-Object -ComObject Shell.Application).Open($window)
                }
            }
        }
        
        Write-Host "Explorer restarted successfully"
    }
    catch {
        Write-Host "Warning: Could not restart Explorer cleanly. Error: $_"
        # Fallback method if the above fails
        Start-Process explorer -WindowStyle Hidden
    }
}

function Disable-Widgets {
    param($widgets)
    
    if ($widgets.disable) {
        Write-Host "Disabling Windows Widgets..."
        
        # Disable Widgets through Registry for current user
        try {
            $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            $loadKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", $true)
            if ($loadKey) {
                $loadKey.SetValue("TaskbarDa", 0, [Microsoft.Win32.RegistryValueKind]::DWord)
                $loadKey.Close()
                Write-Host "Successfully disabled widgets in taskbar"
            }
        }
        catch {
            Write-Host "Warning: Could not modify registry for widgets. Error: $_"
        }

        # Try to remove the Widgets app package
        try {
            Get-AppxPackage "MicrosoftWindows.Client.WebExperience" -AllUsers | Remove-AppxPackage -AllUsers
            Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like "MicrosoftWindows.Client.WebExperience" | Remove-AppxProvisionedPackage -Online
            Write-Host "Successfully removed Widgets app package"
        }
        catch {
            Write-Host "Warning: Could not remove Widgets app package. Error: $_"
        }

        # Disable the service
        try {
            Set-Service "WpnService" -StartupType Disabled
            Stop-Service "WpnService" -Force
            Write-Host "Successfully disabled Widgets service"
        }
        catch {
            Write-Host "Warning: Could not disable Widgets service. Error: $_"
        }
    }
}

# Main Script
try {
    Write-Host "Starting Windows Debloat Tool..."
    
    # Check for admin privileges
    Test-AdminPrivileges

    # Load configuration
    if (-not (Test-Path $configFile)) {
        Write-Host "Error: Configuration file not found"
        exit 1
    }
    $config = Get-Content -Path $configFile | ConvertFrom-Json

    # Create restore point if enabled
    if ($config.createRestorePoint) {
        Write-Host "Creating system restore point..."
        try {
            Enable-ComputerRestore -Drive $env:SystemDrive
            Checkpoint-Computer -Description "Windows Debloat Tool - Before Changes" -RestorePointType "MODIFY_SETTINGS"
        }
        catch {
            Write-Host "Warning: Failed to create restore point. Error: $_"
        }
    }

    # Backup registry if enabled
    if ($config.backupRegistry) {
        Write-Host "Backing up registry keys..."
        try {
            Backup-RegistryKeys
        }
        catch {
            Write-Host "Warning: Failed to backup registry. Error: $_"
        }
    }

    # Execute debloating operations
    try {
        Write-Host "Removing bloatware apps..."
        Remove-BloatwareApps $config.apps
    }
    catch {
        Write-Host "Error during app removal: $_"
    }

    try {
        Write-Host "Disabling services..."
        Disable-Services $config.services
    }
    catch {
        Write-Host "Error during service disabling: $_"
    }

    try {
        Write-Host "Applying privacy settings..."
        Set-PrivacySettings $config.privacy
    }
    catch {
        Write-Host "Error during privacy settings: $_"
    }

    try {
        Write-Host "Optimizing performance..."
        Set-PerformanceSettings $config.performance
    }
    catch {
        Write-Host "Error during performance optimization: $_"
    }

    try {
        Write-Host "Removing features..."
        Remove-Features $config.features
    }
    catch {
        Write-Host "Error during feature removal: $_"
    }

    try {
        Write-Host "Configuring taskbar..."
        Set-TaskbarSettings $config.taskbar
    }
    catch {
        Write-Host "Error during taskbar configuration: $_"
    }

    try {
        Write-Host "Handling widgets..."
        Disable-Widgets $config.widgets
    }
    catch {
        Write-Host "Error during widget configuration: $_"
    }

    Write-Host "`nWindows Debloat Tool completed!"
    Write-Host "Note: Some changes may require a system restart to take effect."
    Write-Host "`nPress any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
catch {
    Write-Host "`nCritical Error occurred: $_"
    Write-Host $_.ScriptStackTrace
    Write-Host "`nPress any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
} 