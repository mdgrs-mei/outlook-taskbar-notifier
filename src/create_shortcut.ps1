Add-Type -AssemblyName System.Windows.Forms
$scriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$rootDir = Split-Path $scriptDir -Parent

. (Join-Path $scriptDir "file_dialog.ps1")

$notifierPath = Join-Path $scriptDir "outlook_taskbar_notifier.ps1"
$settingsPath = GetInputFilePathWithDialog "Select a setting file" "Settings File|*.ps1" $rootDir
if (-not $settingsPath)
{
    return
}
$shortcutPath = GetOutputFilePathWithDialog "Specify output location" ".lnk" "Shortcut File|*.lnk" "shell:startup"
if (-not $shortcutPath)
{
    return
}

. $settingsPath
function GetIconPath()
{
    $settingsDir = Split-Path $script:settingsPath -Parent
    $iconPath = $script:settings.iconPath
    if (-not $iconPath)
    {
        return ""
    }

    Push-Location $settingsDir
    $iconPath = Resolve-Path $iconPath
    Pop-Location
    $iconPath.Path
}
$iconPath = GetIconPath
if ($iconPath)
{
    # Create ico file
    $icoPath = [System.IO.Path]::ChangeExtension($iconPath, ".ico")
    & (Join-Path $scriptDir "convert_image_to_ico.ps1") $iconPath $icoPath
    $iconLocation = $icoPath + ", 0"
}
else
{
    $iconLocation = "powershell.exe" + ", 0"
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-ExecutionPolicy Unrestricted -WindowStyle Hidden -File `"$notifierPath`" `"$settingsPath`""
$shortcut.WorkingDirectory = Split-Path $settingsPath -Parent
$shortcut.WindowStyle = 7 # Minimized
$shortcut.IconLocation = $iconLocation
$shortcut.Save()