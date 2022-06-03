<#
PSAD 6.2 Detection Rule Adds:
-Strict Lookup for targeting one version only
#>

# Function used to look up software - dont modify
function Get-ARPv
{
    param(
        $DisplayName,
        $Version,
        $strict = $false
    )

    # write-host "Strict is: $strict"
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
        foreach ($key32 in $Installed32)
        {
            $key = Get-ItemProperty -Path $Key32.PSPath
            if ($Strict)
            {
                if ([version]($key.DisplayVersion) -eq [version]$Version)
                {
                    $32bit = $True
                }
            }
            if (!$Strict)
            {
                if ([version]($key.DisplayVersion) -ge [version]$Version)
                {
                    $32bit = $True
                }
            }
            
        }
    }

    # If found in registry under 64bit path,
    if ($null -ne $installed64)
    {
        foreach ($key64 in $Installed64)
        {
            $key = Get-ItemProperty -Path $Key64.PSPath
            if ($Strict)
            {
                if ([version]($key.DisplayVersion) -eq [version]$Version)
                {
                    $64bit = $True
                }
            }
            if (!$Strict)
            {
                if ([version]($key.DisplayVersion) -ge [version]$Version)
                {
                    $64bit = $True
                }
            }
            
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


$PlantronicsHub = Get-ARPv -DisplayName "Plantronics Hub Software" -Version "3.22.53245.32743" -strict $true


# -----------------------------------------------------------------------------------------------------------------------------------------
# -And all of your application test results below to return to SCCM if the application(s) are properly installed if there are multiple
# -----------------------------------------------------------------------------------------------------------------------------------------
if ($PlantronicsHub)
{
    write-host "Installed"
}
