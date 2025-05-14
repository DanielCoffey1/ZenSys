# Windows Debloat GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Drawing.Design

# Get absolute path to the script directory
$scriptDirectory = $PSScriptRoot
if (-not $scriptDirectory) {
    $scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
}

# Initialize script paths
$logFile = Join-Path $scriptDirectory "debloat.log"
$script:lastLogMessage = $null

# Create custom colors (Modern Black Theme)
$backgroundColor = [System.Drawing.Color]::Black           # Main background
$contentBgColor = [System.Drawing.Color]::FromArgb(15, 15, 15)  # Content panel
$titleBarColor = [System.Drawing.Color]::Black           # Title bar
$textColor = [System.Drawing.Color]::White               # Regular text
$accentColor = [System.Drawing.Color]::FromArgb(0, 122, 204)   # Buttons and highlights
$warningColor = [System.Drawing.Color]::FromArgb(255, 69, 58)  # Warning text

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows Debloat Tool"
$form.Size = New-Object System.Drawing.Size(600,450)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "None"
$form.BackColor = $backgroundColor
$form.ForeColor = $textColor
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Add form shadow and make it draggable
$form.Add_MouseDown({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        $form.Capture = $false
        $msg = [System.Windows.Forms.Message]::Create($form.Handle, 0xA1, 0x2, 0)
        [System.Windows.Forms.NativeWindow]::FromHandle($form.Handle).DefWndProc([ref]$msg)
    }
})

# Create title bar panel
$titlePanel = New-Object System.Windows.Forms.Panel
$titlePanel.Size = New-Object System.Drawing.Size(600, 30)
$titlePanel.BackColor = $titleBarColor
$titlePanel.Dock = "Top"

# Add title label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Windows Debloat Tool"
$titleLabel.ForeColor = $textColor
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$titleLabel.Location = New-Object System.Drawing.Point(10, 5)
$titleLabel.AutoSize = $true
$titlePanel.Controls.Add($titleLabel)

# Add close button
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "X"
$closeButton.Size = New-Object System.Drawing.Size(30, 30)
$closeButton.Location = New-Object System.Drawing.Point(570, 0)
$closeButton.FlatStyle = "Flat"
$closeButton.FlatAppearance.BorderSize = 0
$closeButton.BackColor = $titleBarColor
$closeButton.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
$closeButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$closeButton.Cursor = "Hand"
$closeButton.Add_Click({ $form.Close() })
$closeButton.Add_MouseEnter({ 
    $this.ForeColor = [System.Drawing.Color]::White
    $this.BackColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
})
$closeButton.Add_MouseLeave({ 
    $this.ForeColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
    $this.BackColor = $titleBarColor
})
$titlePanel.Controls.Add($closeButton)

$form.Controls.Add($titlePanel)

# Create main content panel
$contentPanel = New-Object System.Windows.Forms.Panel
$contentPanel.Size = New-Object System.Drawing.Size(580, 400)
$contentPanel.Location = New-Object System.Drawing.Point(10, 40)
$contentPanel.BackColor = $contentBgColor
$form.Controls.Add($contentPanel)

# Create description label with modern styling
$descLabel = New-Object System.Windows.Forms.Label
$descLabel.Location = New-Object System.Drawing.Point(10,10)
$descLabel.Size = New-Object System.Drawing.Size(560,40)
$descLabel.Text = "This tool will help optimize your Windows installation by removing bloatware, improving privacy, and enhancing performance."
$descLabel.ForeColor = $textColor
$descLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$contentPanel.Controls.Add($descLabel)

# Create warning label with modern styling
$warnLabel = New-Object System.Windows.Forms.Label
$warnLabel.Location = New-Object System.Drawing.Point(10,60)
$warnLabel.Size = New-Object System.Drawing.Size(560,40)
$warnLabel.Text = "WARNING: Please make sure to read the documentation before proceeding. Some changes may not be reversible."
$warnLabel.ForeColor = $warningColor
$warnLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$contentPanel.Controls.Add($warnLabel)

# Create start button with modern styling
$startButton = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(190,200)
$startButton.Size = New-Object System.Drawing.Size(200,45)
$startButton.Text = "Start Debloating"
$startButton.FlatStyle = "Flat"
$startButton.FlatAppearance.BorderSize = 0
$startButton.BackColor = $accentColor
$startButton.ForeColor = $textColor
$startButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
$startButton.Cursor = "Hand"
$startButton.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 184) })
$startButton.Add_MouseLeave({ $this.BackColor = $accentColor })
$contentPanel.Controls.Add($startButton)

# Create splash screen form
$splashForm = New-Object System.Windows.Forms.Form
$splashForm.WindowState = "Maximized"
$splashForm.FormBorderStyle = "None"
$splashForm.BackColor = $backgroundColor
$splashForm.Opacity = 0.95
$splashForm.ShowInTaskbar = $false
$splashForm.TopMost = $true

# Calculate center position for elements based on screen size
$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height

# Create title label at the top
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Windows Debloat Tool"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 24)
$titleLabel.ForeColor = $textColor
$titleLabel.AutoSize = $true
$titleLabel.TextAlign = "MiddleCenter"
# Calculate position first, then set location
$titleX = [Math]::Floor(($screenWidth - $titleLabel.PreferredWidth) / 2)
$titleY = [Math]::Floor($screenHeight * 0.2)  # 20% from top
$titleLabel.Location = New-Object System.Drawing.Point($titleX, $titleY)
$splashForm.Controls.Add($titleLabel)

# Create spinning circle animation
$spinningCircle = New-Object System.Windows.Forms.PictureBox
$spinningCircle.Size = New-Object System.Drawing.Size(150,150)
# Calculate position first, then set location
$circleX = [Math]::Floor(($screenWidth - $spinningCircle.Width) / 2)
$circleY = [Math]::Floor($screenHeight * 0.4)  # 40% from top
$spinningCircle.Location = New-Object System.Drawing.Point($circleX, $circleY)
$spinningCircle.BackColor = [System.Drawing.Color]::Transparent
$splashForm.Controls.Add($spinningCircle)

# Create status label below the circle
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Size = New-Object System.Drawing.Size(800,60)
# Calculate position first, then set location
$statusX = [Math]::Floor(($screenWidth - $statusLabel.Width) / 2)
$statusY = $spinningCircle.Location.Y + $spinningCircle.Height + 40
$statusLabel.Location = New-Object System.Drawing.Point($statusX, $statusY)
$statusLabel.TextAlign = "MiddleCenter"
$statusLabel.ForeColor = $textColor
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$statusLabel.BackColor = [System.Drawing.Color]::Transparent
$splashForm.Controls.Add($statusLabel)

# Create activity indicator label
$activityLabel = New-Object System.Windows.Forms.Label
$activityLabel.Size = New-Object System.Drawing.Size(800,30)
# Calculate position first, then set location
$activityX = [Math]::Floor(($screenWidth - $activityLabel.Width) / 2)
$activityY = $statusLabel.Location.Y + $statusLabel.Height + 10
$activityLabel.Location = New-Object System.Drawing.Point($activityX, $activityY)
$activityLabel.TextAlign = "MiddleCenter"
$activityLabel.ForeColor = $accentColor
$activityLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$activityLabel.BackColor = [System.Drawing.Color]::Transparent
$splashForm.Controls.Add($activityLabel)

# Add escape key handler to close splash screen
$splashForm.KeyPreview = $true
$splashForm.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") {
        $splashForm.Close()
    }
})

# Animation timer
$animationTimer = New-Object System.Windows.Forms.Timer
$animationTimer.Interval = 50

$animationTimer.Add_Tick({
    $script:angle += 10
    if ($script:angle -ge 360) { $script:angle = 0 }
    
    $bitmap = New-Object System.Drawing.Bitmap(150,150)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    
    # Draw background circle
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(40,40,40), 6)
    $graphics.DrawEllipse($pen, 15, 15, 120, 120)
    
    # Draw progress arc
    $pen.Color = $accentColor
    $graphics.DrawArc($pen, 15, 15, 120, 120, $script:angle, ($script:progress * 360 / 100))
    
    $spinningCircle.Image = $bitmap
    $graphics.Dispose()
    $pen.Dispose()
})

# Function to get user-friendly message
function Get-FriendlyMessage {
    param($Message)
    
    # If the message starts with a timestamp, remove it
    if ($Message -match '^\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}:\s(.+)$') {
        $Message = $Matches[1]
    }
    
    # Remove any "(number)" prefix
    $Message = $Message -replace '^\(\d+\)\s*', ''
    
    # Clean up common log message patterns
    $Message = $Message -replace '^Warning: ', ''
    $Message = $Message -replace '^Error: ', ''
    
    switch -Regex ($Message) {
        'Disabling Telemetry' { return "Enhancing your privacy by disabling telemetry..." }
        'Disabling Cortana' { return "Removing Cortana..." }
        'Removing Windows Bloatware' { return "Removing unnecessary apps..." }
        'Disabling Bing Search' { return "Removing Bing integration..." }
        'Disabling Location Tracking' { return "Protecting your location data..." }
        'Disabling Advertising ID' { return "Removing advertising tracking..." }
        'Disabling Activity History' { return "Clearing activity history..." }
        'Disabling Background Apps' { return "Optimizing background processes..." }
        'Disabling Error Reporting' { return "Configuring error handling..." }
        'Disabling Windows Tips' { return "Removing unwanted notifications..." }
        'Disabling OneDrive' { return "Removing OneDrive integration..." }
        'Disabling News and Interests' { return "Removing news feed..." }
        'Disabling Game DVR' { return "Optimizing gaming features..." }
        'Removing Xbox Features' { return "Removing Xbox components..." }
        'Removing Microsoft Teams' { return "Removing Teams..." }
        'Removing Office' { return "Removing Microsoft Office..." }
        'Removing Store Apps' { return "Removing unnecessary Store apps..." }
        'Removing UWP Apps' { return "Removing Windows apps..." }
        'Removing Zune' { return "Removing media features..." }
        'Removing Weather' { return "Removing weather widget..." }
        'Removing People' { return "Removing People app..." }
        'Removing Maps' { return "Removing Maps app..." }
        'Removing Phone' { return "Removing Phone app..." }
        'Removing Camera' { return "Removing Camera app..." }
        'Removing Alarms' { return "Removing Alarms app..." }
        'Removing Calculator' { return "Removing Calculator app..." }
        'Removing Sound Recorder' { return "Removing Voice Recorder..." }
        'Removing Photos' { return "Removing Photos app..." }
        'Removing 3D' { return "Removing 3D apps..." }
        'Removing Paint3D' { return "Removing Paint 3D..." }
        'Removing Mixed Reality' { return "Removing VR features..." }
        'Removing Edge' { return "Removing Edge components..." }
        'Removing Internet Explorer' { return "Removing Internet Explorer..." }
        'Disabling Services' { return "Optimizing system services..." }
        'Disabling Scheduled Tasks' { return "Optimizing scheduled tasks..." }
        'Disabling Windows Search' { return "Optimizing search features..." }
        'Disabling Windows Defender' { return "Configuring security features..." }
        'Disabling Windows Update' { return "Configuring update settings..." }
        'Disabling Superfetch' { return "Optimizing system performance..." }
        'Disabling Print Service' { return "Configuring printer services..." }
        'Disabling Media Player' { return "Removing media features..." }
        'Disabling Remote Desktop' { return "Securing remote access..." }
        'Disabling Indexing' { return "Optimizing file indexing..." }
        'Disabling Web Search' { return "Removing web search integration..." }
        'Disabling Consumer Features' { return "Removing consumer features..." }
        'Disabling AutoLogger' { return "Optimizing system logging..." }
        'Disabling Reserved Storage' { return "Optimizing storage usage..." }
        'Disabling Storage Sense' { return "Configuring storage settings..." }
        'Disabling Hibernate' { return "Configuring power settings..." }
        'Disabling Fast Startup' { return "Optimizing startup settings..." }
        'Moving Taskbar' { return "Customizing taskbar..." }
        'Restoring Classic Context Menu' { return "Restoring classic menus..." }
        'Restoring Classic Explorer' { return "Customizing File Explorer..." }
        'Restarting Explorer' { return "Applying visual changes..." }
        'Creating restore point' { return "Creating backup point..." }
        'Starting main debloating' { return "Starting system optimization..." }
        'Script completed' { return "Finishing up..." }
        'Removing app: (.+)' { return "Removing $($Matches[1])..." }
        'Disabling service: (.+)' { return "Optimizing service: $($Matches[1])..." }
        default {
            # If it's a raw log message, clean it up and return it
            if ($Message -match '^[\w\s]+\.{3}$') {
                return $Message  # Return as is if it's already in the right format
            }
            if ($Message -match 'Removing|Uninstalling|Disabling') {
                return "$Message..."  # Add ellipsis if it's an action
            }
            return $Message
        }
    }
}

# Function to update status
function Update-Status {
    param($Message)
    if ($statusLabel.InvokeRequired) {
        $statusLabel.Invoke([Action]{
            $script:lastMessage = $Message
            $statusLabel.Text = $Message
            # Update activity label with friendly message
            $friendlyMessage = Get-FriendlyMessage $Message
            $activityLabel.Text = $friendlyMessage
        })
    } else {
        $script:lastMessage = $Message
        $statusLabel.Text = $Message
        # Update activity label with friendly message
        $friendlyMessage = Get-FriendlyMessage $Message
        $activityLabel.Text = $friendlyMessage
    }
}

# Timer tick handler for background job
$jobTimer = New-Object System.Windows.Forms.Timer
$jobTimer.Interval = 100  # Faster interval for more responsive updates

$jobTimer.Add_Tick({
    if ($script:job) {
        # Check if we're already completed to prevent multiple completion messages
        if ($script:completed) {
            return
        }

        # Check for new log entries
        try {
            if ($logFile -and (Test-Path $logFile)) {
                $newLogContent = Get-Content -Path $logFile -Tail 1 -ErrorAction SilentlyContinue
                if ($newLogContent -and $newLogContent -ne $script:lastLogMessage) {
                    $script:lastLogMessage = $newLogContent
                    Update-Status $newLogContent
                }
            }
        }
        catch {
            # Silently continue if there's an issue reading the log file
        }

        $output = Receive-Job -Job $script:job -Keep
        if ($output) {
            foreach ($item in $output) {
                if ($item -is [hashtable]) {
                    switch ($item.Type) {
                        "Progress" {
                            Update-Status $item.Message
                            $script:progress += 1
                            if ($script:progress -gt 90) { $script:progress = 90 }
                        }
                        "Result" {
                            if ($item.Success) {
                                # Set completed flag to prevent multiple completion messages
                                $script:completed = $true
                                
                                $script:progress = 95
                                Update-Status "Applying final changes..."
                                
                                # Load config for taskbar settings
                                $config = Get-Content -Path (Join-Path $scriptDirectory "config.json") | ConvertFrom-Json
                                
                                # Apply final changes
                                if (Apply-FinalChanges -config $config) {
                                    $script:progress = 100
                                    Update-Status "Completed successfully!"
                                    $activityLabel.Text = "Done!"
                                    Start-Sleep -Seconds 1
                                    $splashForm.Close()
                                    [System.Windows.Forms.MessageBox]::Show("Debloating completed successfully! It's recommended to restart your computer.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                } else {
                                    Update-Status "Error during final changes"
                                    $activityLabel.Text = "Error!"
                                    Start-Sleep -Seconds 1
                                    $splashForm.Close()
                                    [System.Windows.Forms.MessageBox]::Show("An error occurred during final changes. Some visual settings may not be applied.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                }
                                
                                $jobTimer.Stop()
                                $animationTimer.Stop()
                                Remove-Job -Job $script:job
                                $script:job = $null
                                $startButton.Enabled = $true
                            } else {
                                # Set completed flag to prevent multiple error messages
                                $script:completed = $true
                                
                                Update-Status "Error: $($item.Error)"
                                $activityLabel.Text = "Error!"
                                Start-Sleep -Seconds 1
                                $splashForm.Close()
                                [System.Windows.Forms.MessageBox]::Show("An error occurred: $($item.Error)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                                $jobTimer.Stop()
                                $animationTimer.Stop()
                                Remove-Job -Job $script:job
                                $script:job = $null
                                $startButton.Enabled = $true
                            }
                        }
                    }
                }
                elseif ($item -is [string]) {
                    Update-Status $item
                    $script:progress += 1
                    if ($script:progress -gt 90) { $script:progress = 90 }
                }
            }
        }

        if ($script:job.State -eq 'Failed') {
            # Set completed flag to prevent multiple error messages
            $script:completed = $true
            
            $jobTimer.Stop()
            $animationTimer.Stop()
            Update-Status "Job failed: $($script:job.ChildJobs[0].JobStateInfo.Reason.Message)"
            $activityLabel.Text = "Failed!"
            Start-Sleep -Seconds 1
            $splashForm.Close()
            [System.Windows.Forms.MessageBox]::Show("Job failed: $($script:job.ChildJobs[0].JobStateInfo.Reason.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Remove-Job -Job $script:job
            $script:job = $null
            $startButton.Enabled = $true
        }
    }
})

# Function to apply taskbar settings and restart Explorer
function Apply-FinalChanges {
    param($config)
    
    try {
        Update-Status "Applying final visual changes..."
        
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
        if ($config.taskbar.alignLeft) {
            Set-ItemProperty -Path $registryPath -Name "TaskbarAl" -Value 0 -Type DWord -Force
        }

        # Hide Task View button
        if ($config.taskbar.hideTaskView) {
            Set-ItemProperty -Path $registryPath -Name "ShowTaskViewButton" -Value 0 -Type DWord -Force
        }

        # Configure Search icon
        if ($config.taskbar.hideSearch) {
            Set-ItemProperty -Path $registryPath -Name "SearchboxTaskbarMode" -Value 0 -Type DWord -Force
            Set-ItemProperty -Path $searchRegistryPath -Name "SearchboxTaskbarMode" -Value 0 -Type DWord -Force
            Set-ItemProperty -Path $registryPath -Name "ShowSearchBox" -Value 0 -Type DWord -Force
            Set-ItemProperty -Path $registryPath -Name "ShowSearch" -Value 0 -Type DWord -Force
            Set-ItemProperty -Path $registryPath -Name "SearchboxTaskbarModeCache" -Value 0 -Type DWord -Force
        }

        # Prevent This PC from opening on Explorer restart
        Set-ItemProperty -Path $registryPath -Name "LaunchTo" -Value 1 -Type DWord -Force

        Update-Status "Restarting Explorer..."
        Start-Sleep -Seconds 1

        # Restart Explorer more safely
        $explorerProcess = Get-Process -Name "explorer" -ErrorAction SilentlyContinue
        if ($explorerProcess) {
            $explorerProcess | Stop-Process -Force
            Start-Sleep -Seconds 2
        }
        
        Start-Process "explorer.exe"
        
        return $true
    }
    catch {
        Write-Error "Error applying final changes: $_"
        return $false
    }
}

# Handle the start button click
$startButton.Add_Click({
    $startButton.Enabled = $false
    $script:progress = 0
    $script:completed = $false  # Reset completion flag
    
    try {
        # Check if running as admin
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdmin) {
            [System.Windows.Forms.MessageBox]::Show("Please run this tool as Administrator!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $startButton.Enabled = $true
            return
        }

        # Get path to main script
        $mainScript = Join-Path $scriptDirectory "debloat.ps1"
        
        if (-not (Test-Path $mainScript)) {
            throw "Could not find debloat.ps1 at path: $mainScript"
        }

        # Start background job with script block that includes progress updates
        $script:job = Start-Job -ScriptBlock {
            param($mainScript)
            
            try {
                # Helper function to write progress
                function Write-ProgressUpdate {
                    param($Message)
                    Write-Output @{
                        Type = "Progress"
                        Message = $Message
                    }
                }

                # Check if we can create a restore point
                Write-ProgressUpdate "Checking system protection status..."
                
                # Try to create restore point silently
                try {
                    $null = Enable-ComputerRestore -Drive $env:SystemDrive -ErrorAction SilentlyContinue
                    
                    # Temporarily modify registry to allow more frequent restore points
                    $restorePointKey = "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\SystemRestore"
                    $originalFrequency = $null
                    
                    try {
                        # Backup original frequency value if it exists
                        if (Test-Path $restorePointKey) {
                            $originalFrequency = Get-ItemProperty -Path $restorePointKey -Name "SystemRestorePointCreationFrequency" -ErrorAction SilentlyContinue
                        }
                        
                        # Set frequency to 0 to allow immediate creation
                        Set-ItemProperty -Path $restorePointKey -Name "SystemRestorePointCreationFrequency" -Value 0 -Type DWord -Force
                        
                        # Check if System Protection is enabled
                        $volumeInfo = Get-ComputerRestorePoint -ErrorAction SilentlyContinue
                        if ($null -eq $volumeInfo) {
                            Write-ProgressUpdate "System Protection is not enabled. Proceeding without restore point."
                        } else {
                            # Attempt to create new restore point
                            $null = Checkpoint-Computer -Description "Windows Debloat Tool - Before Changes" -RestorePointType "MODIFY_SETTINGS" -ErrorAction SilentlyContinue
                            
                            # Verify if restore point was created
                            $newPoint = Get-ComputerRestorePoint -ErrorAction SilentlyContinue | 
                                Where-Object { $_.Description -like "*Windows Debloat Tool*" } |
                                Sort-Object CreationTime -Descending |
                                Select-Object -First 1
                                
                            if ($newPoint -and $newPoint.CreationTime -gt (Get-Date).AddMinutes(-5)) {
                                Write-ProgressUpdate "Created new system protection checkpoint."
                            } else {
                                Write-ProgressUpdate "Using existing system protection checkpoint."
                            }
                        }
                    }
                    finally {
                        # Restore original frequency value if it existed
                        if ($originalFrequency -ne $null) {
                            Set-ItemProperty -Path $restorePointKey -Name "SystemRestorePointCreationFrequency" -Value $originalFrequency.SystemRestorePointCreationFrequency -Type DWord -Force
                        } else {
                            Remove-ItemProperty -Path $restorePointKey -Name "SystemRestorePointCreationFrequency" -ErrorAction SilentlyContinue
                        }
                    }
                }
                catch {
                    Write-ProgressUpdate "Proceeding without system protection checkpoint."
                }

                Write-ProgressUpdate "Starting system optimization..."
                
                # Create a new process to run the script to isolate it from Explorer restarts
                $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                $processInfo.FileName = "powershell.exe"
                $processInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$mainScript`""
                $processInfo.RedirectStandardOutput = $true
                $processInfo.RedirectStandardError = $true
                $processInfo.UseShellExecute = $false
                $processInfo.CreateNoWindow = $true
                
                $process = New-Object System.Diagnostics.Process
                $process.StartInfo = $processInfo
                $process.Start() | Out-Null
                
                # Read output while the process is running
                while (-not $process.HasExited) {
                    $line = $process.StandardOutput.ReadLine()
                    if ($line -ne $null) {
                        if ($line -match '^\s*$') { continue }
                        $cleanLine = $line -replace '^\(\d+\)\s*', ''
                        Write-ProgressUpdate $cleanLine
                        Start-Sleep -Milliseconds 50
                    }
                    
                    $errorLine = $process.StandardError.ReadLine()
                    if ($errorLine -ne $null) {
                        Write-ProgressUpdate "Error: $errorLine"
                        Start-Sleep -Milliseconds 50
                    }
                }
                
                # Read any remaining output
                $remainingOutput = $process.StandardOutput.ReadToEnd()
                $remainingError = $process.StandardError.ReadToEnd()
                
                if ($remainingOutput) {
                    foreach ($line in $remainingOutput.Split("`n")) {
                        if ($line -match '^\s*$') { continue }
                        $cleanLine = $line -replace '^\(\d+\)\s*', ''
                        Write-ProgressUpdate $cleanLine
                    }
                }
                
                if ($remainingError) {
                    foreach ($line in $remainingError.Split("`n")) {
                        if ($line -match '^\s*$') { continue }
                        Write-ProgressUpdate "Error: $line"
                    }
                }
                
                # Check process exit code
                if ($process.ExitCode -eq 0) {
                    Write-ProgressUpdate "Optimization completed successfully."
                    @{
                        Type = "Result"
                        Success = $true
                        Error = $null
                    }
                } else {
                    throw "Script exited with code $($process.ExitCode)"
                }
            }
            catch {
                Write-ProgressUpdate "Error occurred: $($_.Exception.Message)"
                @{
                    Type = "Result"
                    Success = $false
                    Error = $_.Exception.Message
                }
            }
        } -ArgumentList $mainScript

        # Show splash screen and start animations
        $animationTimer.Start()
        $jobTimer.Start()
        Update-Status "Initializing..."
        $splashForm.Show()
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $_`n`nScript Directory: $scriptDirectory", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $startButton.Enabled = $true
    }
})

# Show the form
$form.ShowDialog()

# Cleanup
if ($script:job) {
    Stop-Job -Job $script:job
    Remove-Job -Job $script:job
}
if ($animationTimer) {
    $animationTimer.Stop()
    $animationTimer.Dispose()
}
if ($jobTimer) {
    $jobTimer.Stop()
    $jobTimer.Dispose()
} 