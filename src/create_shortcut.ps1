
$srcDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$rootDir = Split-Path $srcDir -Parent

Set-Location $srcDir
. .\file_dialog.ps1
. .\settings_helper.ps1

$notifierPath = Join-Path $srcDir "outlook_taskbar_notifier.ps1"
$settingsPath = GetInputFilePathWithDialog "Select a settings file" "Settings File|*.ps1" $rootDir
if (-not $settingsPath)
{
    return
}
$shortcutPath = GetOutputFilePathWithDialog "Specify output location (The default is your Startup folder)" ".lnk" "Shortcut File|*.lnk" "shell:startup"
if (-not $shortcutPath)
{
    return
}

$settings = InitSettings $settingsPath $notifierPath

$iconPath = $settings.GetFullPath($settings.iconPath)
if ($iconPath)
{
    # Create ico file
    $icoPath = [System.IO.Path]::ChangeExtension($iconPath, ".ico")
    & (Join-Path $srcDir "convert_image_to_ico.ps1") $iconPath $icoPath
    $iconLocation = $icoPath + ", 0"
}
else
{
    $iconLocation = "powershell.exe" + ", 0"
    Write-Host "Could not find the icon file. The default PowerShell icon will be used."
    Read-Host  "Press any key to continue"
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$notifierPath`" `"$settingsPath`""
$shortcut.WorkingDirectory = Split-Path $settingsPath -Parent
$shortcut.WindowStyle = 7 # Minimized
$shortcut.IconLocation = $iconLocation
$shortcut.Save()