# Inspired by - https://github.com/SCConfigMgr/ConfigMgr
# ConfigMgr/Operating System Deployment/Invoke-RemoveBuiltinApps.ps1
# CMTrace compatible logs
# APPX List - https://docs.microsoft.com/en-us/windows/application-management/apps-in-windows-10

Begin
{
    ########################################################################################################################################################################
    $CompanyName            = "Contoso"

    # CMTrace Compatible Log Files
    $Global:CMLogFilePath   = "C:\$CompanyName\RemoveAPPXv4.log"
    $Global:CMLogFileSize   = "1024" # Rollover size in KB
    ########################################################################################################################################################################
    # White list of appx packages to keep installed
    $WhiteListedApps = New-Object -TypeName System.Collections.ArrayList
    [System.Collections.ArrayList]$WhiteListedApps.AddRange(`
    @(
        "Microsoft.DesktopAppInstaller",    # NO REMOVE - App Installer
        "Microsoft.Windows.Photos",
        "Microsoft.StorePurchaseApp",
        "Microsoft.MicrosoftStickyNotes",
        "Microsoft.WindowsAlarms",
        "Microsoft.WindowsCalculator", 
        "Microsoft.WindowsSoundRecorder", 
        "Microsoft.BingWeather",            # MSN Weather
        "Microsoft.WindowsMaps",            # Maps
        "Microsoft.WindowsFeedbackHub",     # Microsoft Windows Feedback Hub
        "Microsoft.WindowsStore",           # NO REMOVE - Windows Store
        "Microsoft.WindowsMaps",
        "Microsoft.WindowsCamera",
        "Microsoft.AAD.BrokerPlugin",
        "Microsoft.WindowsAlarms",
        "Microsoft.MSPaint"
    ))

    # Windows 10 version 1809
    $WhiteListedApps.AddRange(`
    @(
        "Microsoft.ScreenSketch",
        "Microsoft.HEIFImageExtension",
        "Microsoft.VP9VideoExtensions",
        "Microsoft.WebMediaExtensions",
        "Microsoft.WebpImageExtension"
    ))

    # Windows 10 version 1903
    # No new apps
}
Process 
{
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

        # return 0;
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

    # Start the log up
    Start-CMTraceLog -Path $Global:CMLogFilePath

    # Depreciated for CM Trace Logs (BWT)
    function Write-LogEntry
    {
        param(
            [parameter(Mandatory = $true, HelpMessage = "Value added to the RemovedApps.log file.")]
            [ValidateNotNullOrEmpty()]
            [string]$Value,

            [parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
            [ValidateNotNullOrEmpty()]
            [string]$FileName = "RemovedApps.log"
        )
        # Determine log file location
        $LogFilePath = Join-Path -Path $env:windir -ChildPath "Temp\$($FileName)"

        # Add value to log file
        try
        {
            Out-File -InputObject $Value -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
        }
        catch [System.Exception]
        {
            Write-Warning -Message "Unable to append log entry to RemovedApps.log file"
        }
    }

    function Test-RegistryValue
    {
        param(
            [parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Path,
    
            [parameter(Mandatory = $false)]
            [ValidateNotNullOrEmpty()]
            [string]$Name
        )
        # If item property value exists return True, else catch the failure and return False
        try
        {
            if ($PSBoundParameters["Name"])
            {
                $Existence = Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Name -ErrorAction Stop
            }
            else
            {
                $Existence = Get-ItemProperty -Path $Path -ErrorAction Stop
            }
            
            if ($null -ne $Existence)
            {
                return $true
            }
        }
        catch [System.Exception]
        {
            return $false
        }
    }    

    function Set-RegistryValue
    {
        param(
            [parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Path,

            [parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Name,

            [parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$Value,

            [parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [ValidateSet("DWORD", "String")]
            [string]$Type
        )
        try
        {
            $RegistryValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if ($null -ne $RegistryValue)
            {
                Set-ItemProperty -Path $Path -Name $Name -Value $Value -Force -ErrorAction Stop
            }
            else
            {
                New-ItemProperty -Path $Path -Name $Name -PropertyType $Type -Value $Value -Force -ErrorAction Stop | Out-Null
            }
        }
        catch [System.Exception]
        {
            #Write-Warning -Message "Failed to create or update registry value '$($Name)' in '$($Path)'. Error message: $($_.Exception.Message)"
            Write-CMTraceLog -Message "Failed to create or update registry value '$($Name)' in '$($Path)'. Error message: $($_.Exception.Message)" -Type "Error" -Component "Set-RegistryValue"
        }
    }

    # Initial logging
    # Write-LogEntry -Value "Starting built-in AppxPackage, AppxProvisioningPackage and Feature on Demand V2 removal process"
    Write-CMTraceLog -Message "Starting built-in AppxPackage, AppxProvisioningPackage and Feature on Demand V2 removal process" -Type "Info" -Component "Main"

    #region V2Removal
    
    # Disable automatic store updates and disable InstallService
    try
    {
        # Disable auto-download of store apps
        
        #Write-LogEntry -Value "Adding registry value to disable automatic store updates"
        Write-CMTraceLog -Message "Adding registry value to disable automatic store updates" -Type "Info" -Component "Disable Store Updates"

        $RegistryWindowsStorePath = "HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore"
        if (-not(Test-Path -Path $RegistryWindowsStorePath))
        {
            New-Item -Path $RegistryWindowsStorePath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        Set-RegistryValue -Path $RegistryWindowsStorePath -Name "AutoDownload" -Value "2" -Type "DWORD" -ErrorAction Stop

        # Disable the InstallService service
        #Write-LogEntry -Value "Attempting to stop the InstallService service for automatic store updates"
        Write-CMTraceLog -Message "Attempting to stop the InstallService service for automatic store updates" -Type "Info" -Component "Disable Store Updates"

        Stop-Service -Name "InstallService" -Force -ErrorAction Stop
        
        #Write-LogEntry -Value "Attempting to set the InstallService startup behavior to Disabled"
        Write-CMTraceLog -Message "Attempting to set the InstallService startup behavior to Disabled" -Type "Info" -Component "Disable Store Updates"

        Set-Service -Name "InstallService" -StartupType "Disabled" -ErrorAction Stop
    }
    catch [System.Exception]
    {
        #Write-LogEntry -Value "Failed to disable automatic store updates: $($_.Exception.Message)"
        Write-CMTraceLog -Message "Failed to disable automatic store updates: $($_.Exception.Message)" -Type "Error" -Component "Disable Store Updates"
    }
    #endregion V2Removal

    # Determine provisioned apps
    $AppArrayList = Get-AppxProvisionedPackage -Online | Select-Object -ExpandProperty DisplayName

    # Loop through the list of appx packages
    foreach ($App in $AppArrayList)
    {
        #Write-LogEntry -Value "Processing appx package: $($App)"
        Write-CMTraceLog -Message "Processing appx package: $($App)" -Type "Info" -Component "Process AppX Packages"

        # If application name not in appx package white list, remove AppxPackage and AppxProvisioningPackage
        if ($App -in $WhiteListedApps)
        {
            #Write-LogEntry -Value "Skipping excluded application package: $($App)"
            Write-CMTraceLog -Message "   Skipping excluded application package: $($App)" -Type "Info" -Component "Process AppX Packages"
        }
        else
        {
            # Gather package names
            $AppPackageFullName = Get-AppxPackage -Name $App | Select-Object -ExpandProperty PackageFullName -First 1
            $AppProvisioningPackageName = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like $App } | Select-Object -ExpandProperty PackageName -First 1

            # Attempt to remove AppxPackage
            if ($null -ne $AppPackageFullName)
            {
                try
                {
                    # Write-LogEntry -Value "Removing AppxPackage: $($AppPackageFullName)"
                    Write-CMTraceLog -Message "   Removing AppxPackage: $($AppPackageFullName)" -Type "Info" -Component "Process AppX Packages"
                    Remove-AppxPackage -Package $AppPackageFullName -ErrorAction Stop | Out-Null
                }
                catch [System.Exception]
                {
                    #Write-LogEntry -Value "Removing AppxPackage '$($AppPackageFullName)' failed: $($_.Exception.Message)"
                    Write-CMTraceLog -Message "   Removing AppxPackage '$($AppPackageFullName)' failed: $($_.Exception.Message)" -Type "Error" -Component "Process AppX Packages"
                }
            }
            else
            {
                #Write-LogEntry -Value "Unable to locate AppxPackage: $($AppPackageFullName)"
                Write-CMTraceLog -Message "   Unable to locate AppxPackage: $($AppPackageFullName)" -Type "Error" -Component "Process AppX Packages"
            }

            # Attempt to remove AppxProvisioningPackage
            if ($null -ne $AppProvisioningPackageName)
            {
                try
                {
                    #Write-LogEntry -Value "Removing AppxProvisioningPackage: $($AppProvisioningPackageName)"
                    Write-CMTraceLog -Message "Removing AppxProvisioningPackage: $($AppProvisioningPackageName)" -Type "Info" -Component "Process AppX Packages"
                    Remove-AppxProvisionedPackage -PackageName $AppProvisioningPackageName -Online -ErrorAction Stop | Out-Null
                }
                catch [System.Exception]
                {
                    #Write-LogEntry -Value "Removing AppxProvisioningPackage '$($AppProvisioningPackageName)' failed: $($_.Exception.Message)"
                    Write-CMTraceLog -Message "   Removing AppxProvisioningPackage '$($AppProvisioningPackageName)' failed: $($_.Exception.Message)" -Type "Error" -Component "Process AppX Packages"
                }
            }
            else
            {
                #Write-LogEntry -Value "Unable to locate AppxProvisioningPackage: $($AppProvisioningPackageName)"
                Write-CMTraceLog -Message "   Unable to locate AppxProvisioningPackage: $($AppProvisioningPackageName)" -Type "Error" -Component "Process AppX Packages"
            }
        }
    }

    # Enable store automatic updates
    Write-CMTraceLog -Message "Begin Enable Windows Store Updates" -Type "Info" -Component "Enable Windows Store Updates"
    try
    {
        $RegistryWindowsStorePath = "HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore"
        if (Test-RegistryValue -Path $RegistryWindowsStorePath -Name "AutoDownload")
        {
            #Write-LogEntry -Value "Attempting to remove automatic store update registry values"
            Write-CMTraceLog -Message "   Attempting to remove automatic store update registry values" -Type "Info" -Component "Enable Windows Store Updates"
            Remove-ItemProperty -Path $RegistryWindowsStorePath -Name "AutoDownload" -Force -ErrorAction Stop
        }
        #Write-LogEntry -Value "Attempting to set the InstallService startup behavior to Manual"
        Write-CMTraceLog -Message "   Attempting to set the InstallService startup behavior to Manual" -Type "Info" -Component "Enable Windows Store Updates"

        Set-Service -Name "InstallService" -StartupType "Manual" -ErrorAction Stop
    }
    catch [System.Exception]
    {
        #Write-LogEntry -Value "Failed to enable automatic store updates: $($_.Exception.Message)"
        Write-CMTraceLog -Message "   Failed to enable automatic store updates: $($_.Exception.Message)" -Type "Error" -Component "Enable Windows Store Updates"
    }

    #Write-LogEntry -Value "Starting Features on Demand V2 removal process"
    Write-CMTraceLog -Message "Starting Features on Demand V2 removal process" -Type "Info" -Component "Feature on Demand V2 Removal"

    # Get Features On Demand that should be removed
    try
    {
        $OSBuildNumber = Get-WmiObject -Class "Win32_OperatingSystem" | Select-Object -ExpandProperty BuildNumber
        Write-CMTraceLog -Message "   OS Build: $OSBuildNumber" -Type "Info" -Component "Feature on Demand V2 Removal"

        # Handle cmdlet limitations for older OS builds
        if ($OSBuildNumber -le "16299")
        {
            $OnDemandFeatures = Get-WindowsCapability -Online -ErrorAction Stop | Where-Object { $_.Name -notmatch $WhiteListOnDemand -and $_.State -like "Installed" } | Select-Object -ExpandProperty Name
        }
        else
        {
            $OnDemandFeatures = Get-WindowsCapability -Online -LimitAccess -ErrorAction Stop | Where-Object { $_.Name -notmatch $WhiteListOnDemand -and $_.State -like "Installed" } | Select-Object -ExpandProperty Name
        }

        foreach ($Feature in $OnDemandFeatures)
        {
            try
            {
                #Write-LogEntry -Value "Removing Feature on Demand V2 package: $($Feature)"
                Write-CMTraceLog -Message "   Removing Feature on Demand V2 package: $($Feature)" -Type "Info" -Component "Feature on Demand V2 Removal"

                # Handle cmdlet limitations for older OS builds
                if ($OSBuildNumber -le "16299")
                {
                    Get-WindowsCapability -Online -ErrorAction Stop | Where-Object { $_.Name -like $Feature } | Remove-WindowsCapability -Online -ErrorAction Stop | Out-Null
                }
                else
                {
                    Get-WindowsCapability -Online -LimitAccess -ErrorAction Stop | Where-Object { $_.Name -like $Feature } | Remove-WindowsCapability -Online -ErrorAction Stop | Out-Null
                }
            }
            catch [System.Exception]
            {
                #Write-LogEntry -Value "Removing Feature on Demand V2 package failed: $($_.Exception.Message)"
                Write-CMTraceLog -Message "      Removing Feature on Demand V2 package failed: $($_.Exception.Message)" -Type "Error" -Component "Feature on Demand V2 Removal"
            }
        }    
    }
    catch [System.Exception]
    {
        #Write-LogEntry -Value "Attempting to list Feature on Demand V2 packages failed: $($_.Exception.Message)"
        Write-CMTraceLog -Message "   Attempting to list Feature on Demand V2 packages failed: $($_.Exception.Message)" -Type "Error" -Component "Feature on Demand V2 Removal"
    }

    # Complete
    #Write-LogEntry -Value "Completed built-in AppxPackage, AppxProvisioningPackage and Feature on Demand V2 removal process"
    Write-CMTraceLog -Message "   Completed built-in AppxPackage, AppxProvisioningPackage and Feature on Demand V2 removal process" -Type "Info" -Component "Feature on Demand V2 Removal"
}
