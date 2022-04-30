function InitSettings($settingsPath, $notifierPath)
{
    . $settingsPath

    $dir = Split-Path $settingsPath -Parent
    $settings.path = $settingsPath
    $settings.directory = $dir
    $settings.notifierPath = $notifierPath

    $getFullPath = {
        param($relativePath)
        if (-not $relativePath)
        {
            return ""
        }

        Push-Location $this.directory
        $fullPath = Resolve-Path $relativePath
        Pop-Location

        if (-not $fullPath)
        {
            Write-Host "Path does not exist. [$relativePath]"
        }
        $fullPath.Path
    }
    Add-Member -InputObject $settings -MemberType ScriptMethod -Name "GetFullPath" -Value $getFullPath

    $settings
}
