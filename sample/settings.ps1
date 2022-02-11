# The encoding of this file must be UTF-8 with BOM.

$settings = @{

    outlook = @{
        # The outlook folder full path.
        folderPath = "\\your-email-addres@sample.com\folder-name"

        # The path to Outlook exe which is used for the FocusOnFolder action.
        exePath = "C:\Program Files\Microsoft Office\root\office16\OUTLOOK.EXE"
    }

    # Icon file used for the title bar. The path should be either a relative path from this settings file or a full path.
    iconPath = ".\icons\icon.png"

    # Unread count is queried with this interval.
    updateUnreadCountIntervalInSeconds = 3

    # Overlay badge icon settings
    # Available WPF color names are listed here:
    # https://docs.microsoft.com/en-us/dotnet/api/system.windows.media.colors?view=net-5.0
    overlayIcon = @{
        backgroundColor = "DeepPink"
        textColor = "White"         # If you want to hide the unread number, set this to "Transparent".
    }

    # Show a summary of unread items during a mouse over.
    unreadItemsSummary = @{
        enable = $true               # Set $true or $false to toggle this feature.
        maxItemCount = 10
        maxItemCharacterCount = 26
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

    <#
    # Open an unread email if it exists. Open the folder if there is no unread email.
    clickActions = @(
        ,@("Or", @("OpenNewestUnread"), @("FocusOnFolder"))
    )
    #>

    # Thumb buttons
    # You can add max 7 buttons.
    thumbButtons = @(
        ,@{
            description = "Open unread email"
            iconPath = ".\icons\open_mail.png"
            clickActions = @(
                ,@("OpenOldestUnread")
            )
        }
        ,@{
            description = "Mark all as read"
            iconPath = ".\icons\mark_read.png"
            clickActions = @(
                ,@("MarkAllAsRead")
            )
        }
        ,@{
            description = "Jump to notifications page"
            iconPath = ".\icons\web.png"
            clickActions = @(
                ,@("RunCommand", '"C:\Program Files\Mozilla Firefox\firefox.exe"', "-url", "https://github.com/notifications")
            )
        }
        ,@{
            description = "Toggle Do Not Disturb"
            iconPath = ".\icons\notifications_off.png"
            clickActions = @(
                ,@("ToggleDoNotDisturb")
            )
        }
    )
}
