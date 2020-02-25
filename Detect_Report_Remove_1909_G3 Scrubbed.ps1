<#
Copyright 2018-2020 Brian Thorp
This script finds MSI Uninstall strings in the registry by name and version, and runs thier uninstall silently
Used in my In Place Upgrade Task Sequence - Could probably be easily adapted to run within a PSAppDeploy if you want to use this before installing O365
#>

# -----------------------------------------------------------------------------
# Update this Section for your needs
# -----------------------------------------------------------------------------
$OSv = "1909" # Windows 10 1909
$GenericCompanyFolder = "ContosoCorp"

# -----------------------------------------------------------------------------
$Path = "C:\$GenericCompanyFolder\IPU\$OSv" # Keeps things organized with the rest of my scripts
$Global:CMLogFilePath   = "$Path\Logs\GeneralAppRemoval.log"
$Global:CMLogFileSize   = "40"
# -----------------------------------------------------------------------------
# Modified Internet CMTraceLog file functions

# Run this first to seed the files~
function Start-CMTraceLog
{
    # Checks for path to log file and creates if it does not exist
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
            
    )

    $indexoflastslash = $Path.lastindexof('\')
    $directory = $Path.substring(0, $indexoflastslash)

    if (!(test-path -path $directory))
    {
        New-Item -ItemType Directory -Path $directory
    }
    else
    {
        # Directory Exists, do nothing    
    }
}

function Write-CMTraceLog
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
            
        [Parameter()]
        [ValidateSet(1, 2, 3)]
        [int]$LogLevel = 1,

        [Parameter()]
        [string]$Component,

        [Parameter()]
        [ValidateSet('Info','Warning','Error')]
        [string]$Type
    )
    $LogPath = $Global:CMLogFilePath

    Switch ($Type)
    {
        Info {$LogLevel = 1}
        Warning {$LogLevel = 2}
        Error {$LogLevel = 3}
    }

    # Get Date message was triggered
    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"

    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'

    # When used as a module, this gets the line number and position and file of the calling script
    # $RunLocation = "$($MyInvocation.ScriptName | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)"

    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), $Component, $LogLevel
    $Line = $Line -f $LineFormat

    # Write new line in the log file
    Add-Content -Value $Line -Path $LogPath

    # Roll log file over at size threshold
    if ((Get-Item $Global:CMLogFilePath).Length / 1KB -gt $Global:CMLogFileSize)
    {
        $log = $Global:CMLogFilePath
        Remove-Item ($log.Replace(".log", ".lo_"))
        Rename-Item $Global:CMLogFilePath ($log.Replace(".log", ".lo_")) -Force
    }
} 

# Start up the logs
Start-CMTraceLog -Path $Global:CMLogFilePath

Write-CMTraceLog -Message "=====================================================" -Type "Info" -Component "Main"
Write-CMTraceLog -Message "Starting Script..." -Type "Info" -Component "Main"
Write-CMTraceLog -Message "=====================================================" -Type "Info" -Component "Main"


# Function to find MSI-based Uninstallers and Run their uninstall silently
function Remove-MSIApp
{
    param(
        $DisplayName,
        $FileFlag
    )

    Write-CMTraceLog -Message "Start Detection of: $DisplayName" -Type "Info" -Component "Main"

    # PS App Deploy $is64bit
    [boolean]$Is64Bit = [boolean]((Get-WmiObject -Class 'Win32_Processor' -ErrorAction 'SilentlyContinue' | Where-Object { $_.DeviceID -eq 'CPU0' } | Select-Object -ExpandProperty 'AddressWidth') -eq 64)

    $path32 = "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $path64 = "\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

    # =============================================================================
    # -----------------------------------------------------------------------------
    # Run regular code to check for install status
    # Note that this code chunk probably should be updated to the 2020 version~
    # -----------------------------------------------------------------------------
    # Pre-Flight Null
    $32bit = $false
    $64bit = $false
    $Installed32 = $null
    $Installed64 = $null
    # write-host "Software Name:    $DisplayName"
    # write-host "Software Version: $Version"

    $Installed32 = Get-ChildItem HKLM:$path32 -Recurse -ErrorAction Stop | Get-ItemProperty -name DisplayName -ErrorAction SilentlyContinue | Where-Object {$_.DisplayName -like $DisplayName}
    if ($is64bit)
    {
        $Installed64 = Get-ChildItem HKLM:$path64 -Recurse -ErrorAction Stop | Get-ItemProperty -name DisplayName -ErrorAction SilentlyContinue | Where-Object {$_.DisplayName -like $DisplayName}
    }


    # If found in registry,
    if ($null -ne $Installed32)
    {
        Write-CMTraceLog -Message "   App detected on 32-bit registry tree" -Type "Info" -Component "Main"

        foreach ($Entry in $Installed32)
        {
            write-host "Removing $Entry.displayname"
            $Guid = $entry.Pschildname
            $RegistryPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$GUID"
            {
                $UninstallString = (Get-ItemPropertyValue -Path $RegistryPath -Name "UninstallString") + " /qn"

                Write-CMTraceLog -Message "   Attempting to uninstall $GUID | $UninstallString" -Type "Info" -Component "Main"

                &cmd.exe /c $UninstallString
            }
        }
    }

    # If found in registry under 64bit path,
    if ($null -ne $installed64)
    {
        Write-CMTraceLog -Message "   App detected on 64-bit registry tree" -Type "Info" -Component "Main"

        foreach ($Entry in $installed64)
        {
            write-host "Removing $Entry.displayname"
            $Guid = $entry.Pschildname
            $RegistryPath = "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$GUID"
            if (Test-Path $RegistryPath)
            {
                $UninstallString = (Get-ItemPropertyValue -Path $RegistryPath -Name "UninstallString") + " /qn"

                Write-CMTraceLog -Message "   Attempting to uninstall $GUID | $UninstallString" -Type "Info" -Component "Main"

                &cmd.exe /c $UninstallString
            }
        }
    }
}

# Remove junk that interferes with Office 365 Installations
# DisplayName - How this appears in Add Remove Programs
# FileFlag - This seeds a text file into the IPU folder so you can use it as a True/False if exists in your TS to add things back
Remove-MSIApp -DisplayName "Microsoft InfoPath*" -FileFlag "InfoPath"
Remove-MSIApp -DisplayName "Microsoft SharePoint Designer*" -FileFlag "SPDesigner"
Remove-MSIApp -DisplayName "Microsoft Access database engine*" -FileFlag "ADE2010"
# Power Query is built-in to Excel now
Remove-MSIApp -DisplayName "Microsoft Power Query for Excel" -FileFlag ""
Remove-MSIApp -DisplayName "Microsoft Visual Studio 2010 Tools for Office Runtime" -FileFlag "VSTO"
Remove-MSIApp -DisplayName "Skype for Business Web App Plug-in" -FileFlag ""
Remove-MSIApp -DisplayName "Microsoft Skype for Business MUI*" -FileFlag ""
Remove-MSIApp -DisplayName "Skype Meetings App" -FileFlag ""

Write-CMTraceLog -Message "End Script" -Type "Info" -Component "Main"
