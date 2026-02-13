#Requires -Version 5.1

<#
	.SYNOPSIS
	Setup wizard for Install-Office tool
	
	.DESCRIPTION
	This script guides users through the process of downloading and installing Microsoft Office.
	It provides an interactive menu to select Office version, channel, and components.
	
	.EXAMPLE
	.\Setup.ps1
	
	.NOTES
	Run as administrator for full functionality
#>

[CmdletBinding()]
param()

# Check if running as administrator
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin)
{
	Write-Warning -Message "This script should be run as Administrator for best results."
	Write-Host "Some features may not work correctly without administrator privileges." -ForegroundColor Yellow
	Write-Host ""
	$continue = Read-Host "Do you want to continue anyway? (Y/N)"
	if ($continue -ne "Y" -and $continue -ne "y")
	{
		exit
	}
}

# Check PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 5)
{
	Write-Error "PowerShell 5.1 or higher is required. Current version: $($PSVersionTable.PSVersion)"
	exit
}

# Check Windows version
$OSVersion = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption
if ($OSVersion -notmatch "Windows 10" -and $OSVersion -notmatch "Windows 11")
{
	Write-Warning "Office 2019, 2021, 2024, & 365 support Windows 10 & Windows 11 only."
	Write-Host "Current OS: $OSVersion" -ForegroundColor Yellow
	Write-Host ""
}

# Display welcome message
Clear-Host
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Microsoft Office Installation Setup Wizard" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This wizard will help you download and install Microsoft Office." -ForegroundColor White
Write-Host ""

# Set execution policy if needed
$CurrentPolicy = Get-ExecutionPolicy -Scope CurrentUser
if ($CurrentPolicy -ne "Bypass" -and $CurrentPolicy -ne "Unrestricted")
{
	Write-Host "Current execution policy: $CurrentPolicy" -ForegroundColor Yellow
	Write-Host "The execution policy needs to be changed to run the installation scripts." -ForegroundColor Yellow
	$setPolicy = Read-Host "Change execution policy to Bypass for CurrentUser? (Y/N)"
	
	if ($setPolicy -eq "Y" -or $setPolicy -eq "y")
	{
		try
		{
			Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser -Force
			Write-Host "Execution policy changed successfully." -ForegroundColor Green
		}
		catch
		{
			Write-Error "Failed to change execution policy: $_"
			exit
		}
	}
	else
	{
		Write-Warning "Installation may fail without proper execution policy."
	}
	Write-Host ""
}

# Step 1: Select Office Branch
Write-Host "Step 1: Select Office Version" -ForegroundColor Cyan
Write-Host "1. Office 2019 (ProPlus2019Retail)" -ForegroundColor White
Write-Host "2. Office 2021 (ProPlus2021Volume)" -ForegroundColor White
Write-Host "3. Office 2024 (ProPlus2024Volume)" -ForegroundColor White
Write-Host "4. Office 365 (O365ProPlusRetail)" -ForegroundColor White
Write-Host ""

do
{
	$branchChoice = Read-Host "Enter your choice (1-4)"
}
while ($branchChoice -notmatch "^[1-4]$")

$Branch = switch ($branchChoice)
{
	"1" { "ProPlus2019Retail" }
	"2" { "ProPlus2021Volume" }
	"3" { "ProPlus2024Volume" }
	"4" { "O365ProPlusRetail" }
}

Write-Host "Selected: $Branch" -ForegroundColor Green
Write-Host ""

# Step 2: Select Channel
Write-Host "Step 2: Select Update Channel" -ForegroundColor Cyan

if ($Branch -eq "ProPlus2019Retail" -or $Branch -eq "O365ProPlusRetail")
{
	Write-Host "1. Current (Recommended)" -ForegroundColor White
	Write-Host "2. SemiAnnual" -ForegroundColor White
	Write-Host ""
	
	do
	{
		$channelChoice = Read-Host "Enter your choice (1-2)"
	}
	while ($channelChoice -notmatch "^[1-2]$")
	
	$Channel = switch ($channelChoice)
	{
		"1" { "Current" }
		"2" { "SemiAnnual" }
	}
}
elseif ($Branch -eq "ProPlus2021Volume")
{
	$Channel = "PerpetualVL2021"
	Write-Host "Channel: $Channel (Auto-selected)" -ForegroundColor Green
}
elseif ($Branch -eq "ProPlus2024Volume")
{
	$Channel = "PerpetualVL2024"
	Write-Host "Channel: $Channel (Auto-selected)" -ForegroundColor Green
}

Write-Host "Selected: $Channel" -ForegroundColor Green
Write-Host ""

# Step 3: Select Components
Write-Host "Step 3: Select Office Components to Install" -ForegroundColor Cyan
Write-Host "Available components:" -ForegroundColor White
Write-Host "  1. Word" -ForegroundColor White
Write-Host "  2. Excel" -ForegroundColor White
Write-Host "  3. PowerPoint" -ForegroundColor White
Write-Host "  4. Outlook" -ForegroundColor White
Write-Host "  5. Access" -ForegroundColor White
Write-Host "  6. Publisher" -ForegroundColor White
Write-Host "  7. OneNote" -ForegroundColor White
Write-Host "  8. OneDrive" -ForegroundColor White
Write-Host "  9. Teams" -ForegroundColor White
Write-Host ""
Write-Host "Enter component numbers separated by commas (e.g., 1,2,3 for Word, Excel, PowerPoint)" -ForegroundColor Yellow
Write-Host "Or press ENTER for default: Word, Excel, PowerPoint, Outlook" -ForegroundColor Yellow
Write-Host ""

$componentInput = Read-Host "Enter your choices"

if ([string]::IsNullOrWhiteSpace($componentInput))
{
	$Components = @("Word", "Excel", "PowerPoint", "Outlook")
	Write-Host "Using default components: Word, Excel, PowerPoint, Outlook" -ForegroundColor Green
}
else
{
	$componentMap = @{
		"1" = "Word"
		"2" = "Excel"
		"3" = "PowerPoint"
		"4" = "Outlook"
		"5" = "Access"
		"6" = "Publisher"
		"7" = "OneNote"
		"8" = "OneDrive"
		"9" = "Teams"
	}
	
	$Components = @()
	$choices = $componentInput -split "," | ForEach-Object { $_.Trim() }
	foreach ($choice in $choices)
	{
		if ($componentMap.ContainsKey($choice))
		{
			$Components += $componentMap[$choice]
		}
	}
	
	if ($Components.Count -eq 0)
	{
		Write-Warning "No valid components selected. Using default."
		$Components = @("Word", "Excel", "PowerPoint", "Outlook")
	}
}

Write-Host "Selected components: $($Components -join ', ')" -ForegroundColor Green
Write-Host ""

# Step 4: Confirm and Download
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Installation Summary" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Office Version: $Branch" -ForegroundColor White
Write-Host "Update Channel: $Channel" -ForegroundColor White
Write-Host "Components: $($Components -join ', ')" -ForegroundColor White
Write-Host ""
Write-Host "This will download approximately 2-4 GB of data." -ForegroundColor Yellow
Write-Host ""

$confirm = Read-Host "Proceed with download? (Y/N)"

if ($confirm -ne "Y" -and $confirm -ne "y")
{
	Write-Host "Installation cancelled." -ForegroundColor Yellow
	exit
}

Write-Host ""
Write-Host "Starting download..." -ForegroundColor Cyan
Write-Host "Please be patient. This may take 10-30 minutes depending on your connection, and there will be no progress indication." -ForegroundColor Yellow
Write-Host ""

# Run Download.ps1
try
{
	& "$PSScriptRoot\Download.ps1" -Branch $Branch -Channel $Channel -Components $Components
	
	Write-Host ""
	
	# Check if Office folder was created to verify successful download
	# The expected path is: Office\Data\<version>\stream.x64.x-none.dat
	if (Test-Path -Path "$PSScriptRoot\Office\Data\*\stream.x64.x-none.dat")
	{
		Write-Host "Download completed successfully!" -ForegroundColor Green
		Write-Host "Office files are ready for installation." -ForegroundColor Green
		Write-Host ""
		
		# Ask if user wants to install now
		$installNow = Read-Host "Do you want to install Office now? (Y/N)"
		
		if ($installNow -eq "Y" -or $installNow -eq "y")
		{
			Write-Host ""
			Write-Host "Starting Office installation..." -ForegroundColor Cyan
			Write-Host ""
			
			# Run Install.ps1
			try
			{
				& "$PSScriptRoot\Install.ps1"
				
				Write-Host ""
				Write-Host "================================================" -ForegroundColor Green
				Write-Host "  Installation Complete!" -ForegroundColor Green
				Write-Host "================================================" -ForegroundColor Green
				Write-Host ""
				Write-Host "Office has been installed successfully." -ForegroundColor Green
				Write-Host "You may need to restart your computer." -ForegroundColor Yellow
				Write-Host ""
				
				# Optional: Run Configure_Office.ps1
				$configure = Read-Host "Do you want to apply recommended Office configurations? (Y/N)"
				
				if ($configure -eq "Y" -or $configure -eq "y")
				{
					Write-Host ""
					Write-Host "Applying configurations..." -ForegroundColor Cyan
					& "$PSScriptRoot\Configure_Office.ps1"
					Write-Host "Configurations applied successfully." -ForegroundColor Green
				}
			}
			catch
			{
				Write-Error "Installation failed: $_"
				Write-Host ""
				Write-Host "You can try running Install.ps1 manually as administrator." -ForegroundColor Yellow
			}
		}
		else
		{
			Write-Host ""
			Write-Host "Office files have been downloaded." -ForegroundColor Green
			Write-Host "To install Office, run the following command as administrator:" -ForegroundColor White
			Write-Host "  .\Install.ps1" -ForegroundColor Cyan
		}
	}
	else
	{
		Write-Warning "Office files were not found. Download may have failed."
		Write-Host "Please check that Default.xml exists and try again." -ForegroundColor Yellow
	}
}
catch
{
	Write-Error "An error occurred during download: $_"
	Write-Host ""
	Write-Host "You can try running Download.ps1 manually:" -ForegroundColor Yellow
	Write-Host "  .\Download.ps1 -Branch $Branch -Channel $Channel -Components $($Components -join ',')" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "Setup wizard completed." -ForegroundColor Cyan
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
