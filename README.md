# Windows Debloat Tool

A powerful and safe Windows 10/11 debloating tool that helps you remove unnecessary components, disable unwanted services, and improve system performance while customizing the Windows experience.

## Features

### Taskbar Customization

- Align taskbar to the left (Windows 11)
- Hide search bar/icon
- Remove task view button
- Remove widgets and weather
- Clean up taskbar appearance

### Privacy Improvements

- Disable telemetry and data collection
- Disable activity history
- Disable location tracking
- Disable advertising ID
- Disable web search integration
- Disable Cortana
- Disable error reporting
- Disable Windows tips

### Bloatware Removal

- Remove pre-installed Windows apps:
  - Bing Weather
  - Windows Widgets
  - Zune Music/Video
  - Other unnecessary Microsoft apps
- Keep essential apps:
  - Windows Calculator
  - Windows Store
  - Photos
  - Microsoft Edge
  - Notepad
  - Screen Sketch

### Service Optimization

- Disable unnecessary Windows services:
  - Connected User Experiences and Telemetry (DiagTrack)
  - Windows Push Notifications (WpnService)
  - Other non-essential services

### Performance Optimization

- Disable Superfetch
- Optimize visual effects
- Option to disable indexing
- Option to disable OneDrive
- Configurable Windows Search settings

### Feature Management

- Option to remove OneDrive
- Remove Xbox features
- Remove Microsoft Teams
- Remove Cortana

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- Administrator privileges

## Installation

1. Clone or download this repository
2. Extract the files to a directory of your choice
3. Make sure you have the following files:
   - `debloat.ps1` (main script)
   - `config.json` (configuration file)

## Usage

### Quick Start

1. Right-click on PowerShell and select "Run as Administrator"
2. Navigate to the script directory:
   ```powershell
   cd "path\to\script\directory"
   ```
3. Run the script:
   ```powershell
   Set-ExecutionPolicy Bypass -Scope Process -Force
   .\debloat.ps1
   ```

### Customization

You can customize the tool's behavior by editing `config.json`:

- Enable/disable specific features
- Choose which apps to remove or keep
- Configure privacy settings
- Adjust performance options
- Customize taskbar settings

## Safety Features

- Creates a system restore point before making changes
- Backs up registry keys before modification
- Handles errors gracefully
- Logs all operations
- Allows selective component removal
- Provides options to revert changes

## Warning

Please read through the configuration before running. Some changes may not be reversible, and certain removals might affect functionality you want to keep. It's recommended to:

1. Review the `config.json` file first
2. Create a manual system restore point
3. Note which features you're disabling

## Troubleshooting

If you encounter issues:

1. Check the `debloat.log` file for error messages
2. Restore from the automatically created restore point
3. Use the registry backups in the `registry_backup` folder

## After Running

Some changes require a system restart to take full effect. After running the script:

1. Save your work
2. Restart your computer
3. Verify the changes

## License

MIT License

## Detailed Feature List

The tool includes the following features:

Safety measures:

- Creates a system restore point before making changes
- Backs up registry keys
- Handles errors gracefully
- Logs all operations

Bloatware removal:

- Removes pre-installed Windows apps
- Keeps essential apps (configurable)
- Removes provisioned packages

Service optimization:

- Disables unnecessary Windows services
- Configurable through the config file

Privacy improvements:

- Disables telemetry
- Disables Cortana
- Disables activity history
- Other privacy-related settings

Performance optimization:

- Disables Superfetch
- Optimizes visual effects
- Optional indexing and Windows Defender settings

Feature removal:

- OneDrive removal
- Teams removal
- Xbox features removal
