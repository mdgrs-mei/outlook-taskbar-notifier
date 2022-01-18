Add-Type -AssemblyName PresentationFramework

$scriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
Set-Location $scriptDir

. .\settings_helper.ps1
. .\process.ps1
. .\outlook_folder.ps1
. .\delegate_command.ps1
. .\window.ps1
. .\action_generator.ps1

$settingsPath = $args[0]
. $settingsPath
SetSettingsDirectory $settings $settingsPath

$outlookFolder = [OutlookFolder]::new()
$outlookFolder.Init($settings.outlook.folderPath, $settings.outlook.exePath)

$windowTitle = $outlookFolder.GetName()
$window = [Window]::new()
$window.Init(".\window.xaml", $windowTitle, $settings)

$actionGenerator = [ActionGenerator]::new()
$actionGenerator.Init($outlookFolder, $window)

$clickActions = $actionGenerator.CreateActionSequence($settings.clickActions)
$window.SetOnClickedFunction($clickActions)

foreach ($thumbButtonSetting in $settings.thumbButtons)
{
    $button = $window.AddThumbButton($thumbButtonSetting)
    $actionSequence = $actionGenerator.CreateActionSequence($thumbButtonSetting.clickActions)
    $button.Command = New-Object DelegateCommand($actionSequence)
}

$window.StartTimerFunction({
    $outlookFolder.InitOutlookIfNotValid()
    $unreadCount = $outlookFolder.GetUnreadCount()
    $window.UpdateUnreadCount($unreadCount)

    if ($settings.unreadItemsSummary.enable)
    {
        $unreadItemsSummary = $outlookFolder.GetUnreadItemsSummary($settings.unreadItemsSummary.maxItemCount)
        $window.SetTaskbarItemInfoDescription($unreadItemsSummary)
    }
}, $settings.updateUnreadCountIntervalInSeconds)
$window.ShowDialog()
$window.Term()

$actionGenerator.Term()
$outlookFolder.Term()
