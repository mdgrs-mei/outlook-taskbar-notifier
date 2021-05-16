# Outlook Taskbar Notifier
Outlook Taskbar Notifier is a Windows taskbar application that notifies you of unread emails for an Outlook folder by showing an overlay badge and optionally flashing the taskbar icon.

![taskbar](./docs/taskbar.png)

The notifier makes it easier for you to find new emails moved into sub folders. If you use folders to receive notifications from some web services (Jira, GitHub, GitLab etc.) you can boot the notifier for each folder and they will work like desktop app notifications for the services.

## Features
- Notifies unread email count for an Outlook folder
- Customizable icons and badge colors
- Customizable actions performed when the notifier is clicked

## Requirements
- Tested on Windows 10 and Powershell 5.1
- Outlook desktop app needs to be installed

## Installation
1. Download and extract the zip anywhere you like
1. Edit `settings.ps1` (See [Settings](#Settings))
1. Run [`tools/create_shortcut.bat`](#toolscreate_shortcutbat) and save the shortcut
1. Run the shortcut

# Settings
You can customize the behavior by a settings file. A sample settings file is placed at [sample/settings.ps1](./sample/settings.ps1).

## Outlook settings

```powershell
outlook = @{
    folderPath = "\\your-email-address@sample.com\folder-name"
    exePath = "C:\Program Files\Microsoft Office\root\office16\OUTLOOK.EXE"
}
```
`folderPath` is a path of the outlook folder which the notifier monitors. You can list all your folder paths by running [`tools/list_outlook_folders.bat`](#toolslist_outlook_foldersbat).

`exePath` is a path to Outlook exe. It is used in `FocusOnFolder` action.

## Icon image

![icons](./docs/icons.png)
```powershell
iconPath = ".\icon.png"
```
An icon file used for the title bar. The image is converted to `.ico` file during the shortcut creation and also used as a shortcut icon. `.bmp`, `.png`, `.tif`, `.gif` and `.jpg` with single resolution are supported.

## Overlay icon

![overlay_icon](./docs/overlay_icon.png)
``` powershell
overlayIcon = @{
    backgroundColor = "DeepPink"
    textColor = "White"
}
```

You can change the unread badge color. Available WPF color names are listed here:
https://docs.microsoft.com/en-us/dotnet/api/system.windows.media.colors?view=net-5.0

## Taskbar icon flashing

``` powershell
flashTaskbar = @{
    enable = $true
    rateInMillisecond = 300
    count = 2
}
```
You can change the flash rate and count or you can completely disable this feature.

## Click actions

``` powershell
clickActions = @(
    ,@("FocusOnFolder")
)
```

When the notifier is clicked, actions you specify here are executed sequentially.

### Actions
|Action Name|Description|
|---|---|
|FocusOnFolder|Opens the Outlook folder which the notifier is monitoring.|
|MarkAllAsRead|Marks all emails in the folder as read.|
|ToggleDoNotDisturb|Toggles Do Not Disturb mode. In Do Not Disturb mode, unread count update and icon flashing are paused.|
|FocusOnApp|Focuses on the specified app window. The app name is a process name which you can check in the properties window from Task Manager.|
|SendKeysToAppInFocus|Sends keyboard inputs to the app in focus. The key string format follows Windows.Forms.SendKeys format: https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys?view=net-5.0|
|SleepMilliseconds|Waits for specified milliseconds to tweak the timing for SendKeys.|
|RunCommand|Runs any executable with arguments.|

### More complex examples
1. Marks all as read and opens a web page. Recommended for web services which have notifications pages but don't support desktop notifications.

    ``` powershell
    # Opens GitHub notifications page.
    clickActions = @(
        ,@("MarkAllAsRead")
        ,@("RunCommand", '"C:\Program Files\Mozilla Firefox\firefox.exe"', "-url", "https://github.com/notifications")
    )
    ```

1. Marks all as read and sends keys to an app. Recommended for services which have shortcut commands to their notifications page.

    ``` powershell
    # Opens GitHub notifications page by sending shortcut keys to a pinned tab in Chrome.
    clickActions = @(
        ,@("MarkAllAsRead")
        ,@("FocusOnApp", "chrome.exe")
        ,@("SendKeysToAppInFocus", "^1")
        ,@("SleepMilliseconds", 100)
        ,@("SendKeysToAppInFocus", "gn")
    )
    ```

## Thumb buttons

<img src="./docs/thumb_buttons.png" width=260>

``` powershell
thumbButtons = @(
    ,@{
        description = "Mark all as read"
        iconPath = ".\mark_read.png"
        clickActions = @(
            ,@("MarkAllAsRead")
        )
    }
}
```
You can add maximum 7 thumb buttons and define actions performed when they are clicked. The formats of `iconPath` and `clickActions` are the same as before.

# Tools

## [tools/list_outlook_folders.bat](./tools/list_outlook_folders.bat)

Lists all the Outlook folder paths which the notifier can monitor. Copy one of the folder paths and paste it in your settings file.

## [tools/create_shortcut.bat](./tools/create_shortcut.bat)

This tool takes a settings file and creates a shortcut to run the notifier. If you want to create another notifier which monitors another Outlook folder, you just need to create a settings file and run this tool again.

## [tools/convert_image_to_ico.bat](./tools/convert_image_to_ico.bat)

Converts an image to `.ico` file. When you want to update the icon of the shortcut, run this tool.