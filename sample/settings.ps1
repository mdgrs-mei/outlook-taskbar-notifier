﻿# The encoding of this file must be UTF-8 with BOM.

$settings = @{

    outlook = @{
        # The outlook folder full path.
        folderPath = "\\your-email-addres@sample.com\folder-name"

        # The path to Outlook exe which is used for the FocusOnFolder action.
        exePath = "C:\Program Files\Microsoft Office\root\office16\OUTLOOK.EXE"
    }

    # Icon file used for the title bar. The path should be either a relative path from this settings file or a full path.
    iconPath = ".\icon.png"

    # Unread count is queried with this interval.
    updateUnreadCountIntervalInSeconds = 3

    # Overlay badge icon settings
    # Available WPF color names are listed here:
    # https://docs.microsoft.com/en-us/dotnet/api/system.windows.media.colors?view=net-5.0
    overlayIcon = @{
        backgroundColor = "DeepPink"
        textColor = "White"
    }

    # Flashes taskbar when unread email count increases.
    flashTaskbar = @{
        enable = $true                # Set $true or $false to toggle this feature.
        rateInMillisecond = 300       # Flash interval. Set 0 to use the system default rate.
        count = 2                     # Flash count
    }

    # Click Actions
    # The actions are executed sequentially when the app on the taskbar is clicked.

    # Basic setting
    clickActions = @(
        ,@("FocusOnFolder")                     # Open or focus on the folder in Outlook
    )

    <#
    # Open a new GitHub notifications page with firefox
    clickActions = @(
        ,@("MarkAllAsRead")                     # Mark all emails in the folder as read
        ,@("RunCommand", '"C:\Program Files\Mozilla Firefox\firefox.exe"', "-url", "https://github.com/notifications")   # Run executables with arguments
    )
    #>

    <#
    # Open GitHub notifications page which is pinned to tab1 on chrome
    clickActions = @(
        ,@("MarkAllAsRead")
        ,@("FocusOnApp", "chrome.exe")      # Focus on an application
        ,@("SendKeysToAppInFocus", "^1")    # Send keyboard input to the app in focus. The key string format follows Windows.Forms.SendKeys format: https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys?view=net-5.0
        ,@("SleepMilliseconds", 100)        # Wait for the specified milliseconds
        ,@("SendKeysToAppInFocus", "gn")    # GitHub notifications shortcut
    )
    #>

    # Thumb buttons
    # You can add max 7 buttons.
    thumbButtons = @(
        ,@{
            description = "Open unread email"
            iconPath = ".\open_mail.png"
            clickActions = @(
                ,@("OpenOldestUnread")
            )
        }
        ,@{
            description = "Mark all as read"
            iconPath = ".\mark_read.png"
            clickActions = @(
                ,@("MarkAllAsRead")
            )
        }
        ,@{
            description = "Jump to notifications page"
            iconPath = ".\web.png"
            clickActions = @(
                ,@("RunCommand", '"C:\Program Files\Mozilla Firefox\firefox.exe"', "-url", "https://github.com/notifications")
            )
        }
        ,@{
            description = "Toggle Do Not Disturb"
            iconPath = ".\notifications_off.png"
            clickActions = @(
                ,@("ToggleDoNotDisturb")
            )
        }
    )
}