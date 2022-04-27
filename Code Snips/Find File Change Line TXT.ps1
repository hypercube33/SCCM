if (test-path $FileLocation)
{
    $file = Get-Content $FileLocation
    $containsWord = $file | ForEach-Object { $_ -match "$INI" }
    if ($containsWord -contains $true)
    {
        write-host "Installed"
    }
}