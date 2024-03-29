﻿Add-Type -AssemblyName PresentationFramework

$notifierPath = $MyInvocation.MyCommand.Path
$srcDir = Split-Path $notifierPath -Parent
Set-Location $srcDir

. .\settings_helper.ps1
. .\process.ps1
. .\outlook_folder.ps1
. .\delegate_command.ps1
. .\window.ps1
. .\action_generator.ps1

$settingsPath = $args[0]
$settings = InitSettings $settingsPath $notifierPath

$outlookFolder = [OutlookFolder]::new()
$outlookFolder.Init($settings)

$windowTitle = $outlookFolder.GetName()
$window = [Window]::new()
$window.Init(".\window.xaml", $windowTitle, $settings)
if (-not ($args -Contains "-SkipJumpList"))
{
    $window.InitJumpList()
}

$updateUnreadFunc = {
    $unreadCount = $outlookFolder.GetUnreadCount()
    $window.UpdateUnreadCount($unreadCount)

    if ($settings.unreadItemsSummary.enable)
    {
        $unreadItemsSummary = $outlookFolder.GetUnreadItemsSummary(
            $settings.unreadItemsSummary.maxItemCount,
            $settings.unreadItemsSummary.maxItemCharacterCount)
        $window.SetTaskbarItemInfoDescription($unreadItemsSummary)
    }
}.GetNewClosure()

$actionGenerator = [ActionGenerator]::new()
$actionGenerator.Init($outlookFolder, $window, $settings, $updateUnreadFunc)

$clickActions = $actionGenerator.CreateActionSequence($settings.clickActions)
$window.SetOnClickedFunction($clickActions)

foreach ($thumbButtonSetting in $settings.thumbButtons)
{
    $button = $window.AddThumbButton($thumbButtonSetting)
    $actionSequence = $actionGenerator.CreateActionSequence($thumbButtonSetting.clickActions)
    $button.Command = New-Object DelegateCommand($actionSequence)
}

function Update()
{
    $outlookFolder.InitOutlookIfNotValid()
    if ($settings.doNotDisturb.globalMode)
    {
        $window.ReferToGlobalDoNotDisturb()
    }
    Invoke-Command $updateUnreadFunc
}

Update
$window.StartTimerFunction({Update}, $settings.updateUnreadCountIntervalInSeconds)
$window.ShowDialog()
$window.Term()

$actionGenerator.Term()
$outlookFolder.Term()
