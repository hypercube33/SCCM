$Company = "Contoso"
$OSV = "1909"

# Function used to look up software - dont modify
function Get-ARPv
{
    param(
        $DisplayName,
        $Version
    )

    # -----------------------------------------------------------------------------
    # Global Stuff
    # -----------------------------------------------------------------------------
    $InstallGlobal = $null

    # PS App Deploy $is64bit
    [boolean]$Is64Bit = [boolean]((Get-WmiObject -Class 'Win32_Processor' -ErrorAction 'SilentlyContinue' | Where-Object { $_.DeviceID -eq 'CPU0' } | Select-Object -ExpandProperty 'AddressWidth') -eq 64)

    $path32 = "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $path64 = "\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

    # =============================================================================
    # -----------------------------------------------------------------------------
    # Run regular code to check for install status
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
        $key = Get-ItemProperty -Path $Installed32.PSPath
        if ([version]($key.DisplayVersion) -ge [version]$Version)
        {
            $32bit = $True
        }
    }

    # If found in registry under 64bit path,
    if ($null -ne $installed64)
    {
        $key = Get-ItemProperty -Path $Installed64.PSPath
        if ([version]($key.DisplayVersion) -ge [version]$Version)
        {
            $64bit = $True
        }
    }

    # Installed, take existing result and 
    if ($32bit -or $64bit)      {$InstallGlobal = $True}
    else                        {$InstallGlobal = $false}

    return $InstallGlobal
}

# Function used to find VSTO - Compatible with Windows 7 (x64 only)
function Get-VSTO
{
    param(
        $TargetVersion
    )
    
    $vsto = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R"
    if (Test-Path $vsto)
    {
        $key = Get-ItemProperty -Path $VSTO
        $Version = [version]$key.Version
        if ($version -ge [version]$targetversion)
        {
            return $true
        }
    }
}

# -----------------------------------------------------------------------------------------------------------------------------------------
# Make your changes below - set variables per example per each piece of software required, or write your own function above 
# for specific software. Must return $True or $False
# Note for Get-ARPv you must pass a proper version number e.g 1.0 - wildcards are not supported
# -----------------------------------------------------------------------------------------------------------------------------------------

$O365        = Get-ARPv -DisplayName "Microsoft Office 365 ProPlus - en-us" -Version "16.0.0.0"
$Visio       = Get-ARPv -DisplayName "Microsoft Visio Professional*" -Version "16.0.0.0"
$Project     = Get-ARPv -DisplayName "Microsoft Project Professional*" -Version "16.0.0.0"



# -----------------------------------------------------------------------------------------------------------------------------------------
# -And all of your application test results below to return to SCCM if the application(s) are properly installed if there are multiple
# -----------------------------------------------------------------------------------------------------------------------------------------
if ($O365 -or $Visio -or $Project)
{
    # Uninstall Office 365
    &cmd.exe /c "setup.exe /configure SilentUninstallConfig.xml"

    # Document Office 365
    write-host "Office Uninstalled" | Out-File -FilePath "C:\$Company\IPU\$OSV\Office_Scrubbed.txt"
}