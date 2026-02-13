# Installation Guide

This guide will help you install Microsoft Office using the Install-Office tool.

## System Requirements

- **Operating System**: Windows 10 or Windows 11
- **PowerShell**: Version 5.1 or higher
- **Disk Space**: At least 4 GB free space
- **Internet Connection**: Required for downloading Office files
- **Administrator Rights**: Required for installation

## Quick Start (Recommended)

The easiest way to install Office is using the Setup Wizard:

1. **Download** this repository or clone it:
   ```bash
   git clone https://github.com/adrianmpala/Install-Office.git
   cd Install-Office
   ```

2. **Right-click** on `Setup.ps1` and select **"Run with PowerShell"** (as Administrator)
   
   Or open PowerShell as Administrator and run:
   ```powershell
   .\Setup.ps1
   ```

3. **Follow the wizard** to:
   - Select Office version (2019, 2021, 2024, or 365)
   - Choose update channel
   - Select components to install (Word, Excel, PowerPoint, etc.)
   - Download and install Office

That's it! The wizard handles everything automatically.

## Manual Installation

If you prefer manual control or the wizard doesn't work:

### Step 1: Set Execution Policy

Open PowerShell as Administrator and run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser -Force
```

### Step 2: Download Office

Navigate to the Install-Office folder and choose your Office version:

**For Office 2019:**
```powershell
.\Download.ps1 -Branch ProPlus2019Retail -Channel Current -Components Word,Excel,PowerPoint,Outlook
```

**For Office 2021:**
```powershell
.\Download.ps1 -Branch ProPlus2021Volume -Channel PerpetualVL2021 -Components Word,Excel,PowerPoint,Outlook
```

**For Office 2024:**
```powershell
.\Download.ps1 -Branch ProPlus2024Volume -Channel PerpetualVL2024 -Components Word,Excel,PowerPoint,Outlook
```

**For Office 365:**
```powershell
.\Download.ps1 -Branch O365ProPlusRetail -Channel Current -Components Word,Excel,PowerPoint,Outlook,OneDrive,Teams
```

**Note:** The download process is silent - no progress will be shown. Wait for the command to complete (may take 10-30 minutes depending on your connection).

### Step 3: Install Office

After downloading, run:

```powershell
.\Install.ps1
```

This will install the Office components you downloaded.

### Step 4: Configure Office (Optional)

Apply recommended Office configurations:

```powershell
.\Configure_Office.ps1
```

This script configures various Office settings such as:
- Disabling telemetry
- Enabling Developer tab
- Configuring AutoSave
- And more...

## Available Components

You can choose which Office applications to install:

- **Word** - Word processor
- **Excel** - Spreadsheet application
- **PowerPoint** - Presentation software
- **Outlook** - Email and calendar
- **Access** - Database management
- **Publisher** - Desktop publishing
- **OneNote** - Note-taking application
- **OneDrive** - Cloud storage
- **Teams** - Collaboration platform
- **ProjectPro2019Volume** - Project management (2019)
- **ProjectPro2021Volume** - Project management (2021)
- **ProjectPro2024Volume** - Project management (2024)

## Office Versions

### Office 2019
- **Branch**: ProPlus2019Retail
- **Channel**: Current
- **Support**: Windows 10 and 11

### Office 2021
- **Branch**: ProPlus2021Volume
- **Channel**: PerpetualVL2021
- **Support**: Windows 10 and 11

### Office 2024
- **Branch**: ProPlus2024Volume
- **Channel**: PerpetualVL2024
- **Support**: Windows 10 and 11

### Office 365
- **Branch**: O365ProPlusRetail
- **Channel**: Current or SemiAnnual
- **Support**: Windows 10 and 11
- **Note**: Requires subscription

## Update Channels

- **Current**: Latest features as soon as available (recommended for most users)
- **SemiAnnual**: Updates twice per year with more stability
- **PerpetualVL2021**: For Office 2021 volume license
- **PerpetualVL2024**: For Office 2024 volume license

## Troubleshooting

### Execution Policy Error

If you get an error about execution policy:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser -Force
```

### Download Doesn't Start

1. Make sure you're connected to the internet
2. Check that `Default.xml` exists in the folder
3. Try running PowerShell as Administrator

### Installation Fails

1. Make sure you ran `Download.ps1` first
2. Check that the `Office` folder exists in the script directory
3. Run `Install.ps1` as Administrator
4. Make sure no Office applications are currently running

### Office Not Activating

The tool downloads and installs Office but doesn't activate it. You'll need to:
- Use a valid product key
- Or sign in with a Microsoft 365 account (for Office 365)
- Or use your organization's activation method

## Uninstalling Office

To uninstall Office, use the official uninstall tools:

1. Run the uninstall command:
   ```cmd
   .\Uninstall_Office_updates.cmd
   ```

2. Or use the Office_Uninstall tool:
   ```cmd
   cd Office_Uninstall
   Office_Uninstall.cmd
   ```

## Additional Resources

- [Official Office Deployment Tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
- [Office Configuration Tool](https://config.office.com/deploymentsettings)
- [Office Update Channels](https://learn.microsoft.com/en-us/deployoffice/overview-update-channels)
- [Deploy Office Guide](https://learn.microsoft.com/en-us/deployoffice/deployment-guide-microsoft-365-apps)

## Support

For issues with this tool, please visit:
- [GitHub Repository](https://github.com/adrianmpala/Install-Office)
- [GitHub Issues](https://github.com/adrianmpala/Install-Office/issues)

For Office product support, visit [Microsoft Support](https://support.microsoft.com/office).

---

**Note**: This is an unofficial tool for deploying Microsoft Office. All Office products are property of Microsoft Corporation.
