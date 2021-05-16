function SetSettingsDirectory($settings, $settingsPath)
{
    $dir = Split-Path $settingsPath -Parent
    $settings.directory = $dir
}

function GetFullPathFromSettingsRelativePath($settings, $path)
{
    if (-not $path)
    {
        return ""
    }

    Push-Location $settings.directory
    $fullPath = Resolve-Path $path
    Pop-Location
    $fullPath.Path
}