# process.ps1 must be included before executing this file.

class ActionGenerator
{
    $actionTable = @{}

    [void] Init($outlookFolder, $window)
    {
        $this.actionTable = @{
            "FocusOnFolder" = {
                $outlookFolder.Focus()
                $true
            }.GetNewClosure()

            "OpenNewestUnread" = {
                $outlookFolder.OpenNewestUnread()
                $unreadCount = $outlookFolder.GetUnreadCount()
                $window.UpdateUnreadCount($unreadCount)
                $true
            }.GetNewClosure()

            "OpenOldestUnread" = {
                $outlookFolder.OpenOldestUnread()
                $unreadCount = $outlookFolder.GetUnreadCount()
                $window.UpdateUnreadCount($unreadCount)
                $true
            }.GetNewClosure()

            "MarkAllAsRead" = {
                $outlookFolder.MarkAllAsRead()
                $unreadCount = $outlookFolder.GetUnreadCount()
                $window.UpdateUnreadCount($unreadCount)
                $true
            }.GetNewClosure()

            "ToggleDoNotDisturb" = {
                $window.ToggleDoNotDisturb()
                $unreadCount = $outlookFolder.GetUnreadCount()
                $window.UpdateUnreadCount($unreadCount)
                $true
            }.GetNewClosure()

            "FocusOnApp" = {
                param($appName)
                FocusApp $appName
            }

            "SendKeysToAppInFocus" = {
                param($key)
                SendKeysToActiveApp $key
                $true
            }

            "SleepMilliseconds" = {
                param($millisec)
                Start-Sleep -Milliseconds $millisec
                $true
            }

            "RunCommand" = {
                if ($args.Length -gt 1)
                {
                    Start-Process $args[0] -ArgumentList $args[1..($args.Length-1)] -NoNewWindow
                }
                else
                {
                    Start-Process $args[0] -NoNewWindow
                }
                $true
            }

            "RunCommandAndWait" = {
                if ($args.Length -gt 1)
                {
                    Start-Process $args[0] -ArgumentList $args[1..($args.Length-1)] -NoNewWindow -Wait
                }
                else
                {
                    Start-Process $args[0] -NoNewWindow -Wait
                }
                $true
            }
        }
    }

    [void] Term()
    {
    }

    [Object] CreateActionSequence($actionSettings)
    {
        $class = $this

        $block = {
            $class.ExecuteActions($actionSettings)
        }.GetNewClosure()

        return $block
    }

    [void] ExecuteActions($actionSettings)
    {
        try
        {
            foreach ($actionSetting in $actionSettings)
            {
                $actionSetting = @($actionSetting)
                Write-Host "Action:"
                Write-Host $actionSetting

                $actionName = $actionSetting[0]
                $actionArgs = $actionSetting[1..($actionSetting.Count-1)]
                $block = $this.actionTable[$actionName]
                if ($block)
                {
                    $success = Invoke-Command $block -ArgumentList $actionArgs
                    if (-not $success)
                    {
                        return
                    }
                }
            }    
        }
        catch
        {
            Write-Host "Action failed. [$PSItem]"
        }
    }
}

