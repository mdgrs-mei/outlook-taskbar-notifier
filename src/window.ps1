# process.ps1 must be included before executing this file.
# "Add-Type -AssemblyName PresentationFramework" must be called before including this file.

$srcDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$globalDoNotDisturbFilePath = Join-Path (Split-Path $srcDir -Parent) "do_not_disturb"

. (Join-Path $srcDir "flash_window.ps1")

class Window
{
    $window
    $mainWindowHandle
    $settings
    $timer
    $onClicked
    $lastUnreadCount = 0
    $isFlashing = $false
    $skipOnClicked = $false
    $doNotDisturb = $false

    [void] Init($xamlPath, $title, $settings)
    {
        $this.settings = $settings

        $xaml = [xml](Get-Content $xamlPath)
        $nodeReader = (New-Object System.Xml.XmlNodeReader $xaml)
        $this.window = [System.Windows.Markup.XamlReader]::Load($nodeReader)
        $this.window.Title = $title

        $iconPath = $settings.GetFullPath($settings.iconPath)
        if ($iconPath)
        {
            $this.window.Icon = $iconPath
        }

        # Start with Normal window to make Windows draw preview window.
        $this.window.WindowState = [System.Windows.WindowState]::Normal

        $thisInstance = $this

        $this.window.add_Loaded({
            $thisInstance.OnLoaded()
        }.GetNewClosure())

        $this.window.add_ContentRendered({
            $thisInstance.OnContentRendered()
        }.GetNewClosure())

        $this.window.add_StateChanged({
            $thisInstance.OnStateChanged()
        }.GetNewClosure())
    }

    [void] Term()
    {
    }

    [void] InitJumpList()
    {
        $jumpList = New-Object System.Windows.Shell.JumpList

        $jumpTask = New-Object System.Windows.Shell.JumpTask
        $jumpTask.Title = "Open settings location"
        $jumpTask.ApplicationPath = "explorer.exe"
        $jumpTask.IconResourcePath = "explorer.exe"
        $jumpTask.Arguments = "/select,`"{0}`"" -f $this.settings.path
        $jumpList.JumpItems.Add($jumpTask)

        $jumpTask = New-Object System.Windows.Shell.JumpTask
        $jumpTask.Title = "Run with debug console"
        $jumpTask.Arguments = "-ExecutionPolicy Bypass -File `"{0}`" `"{1}`" -SkipJumpList" -f $this.settings.notifierPath, $this.settings.path
        $jumpList.JumpItems.Add($jumpTask)

        $jumpList.Apply()
    }

    [void] OnLoaded()
    {
        $this.mainWindowHandle = (New-Object System.Windows.Interop.WindowInteropHelper($this.window)).Handle
    }

    [void] OnContentRendered()
    {
        # Immediately minimize the window after the thumbnail is rendered.
        $this.window.WindowState = [System.Windows.WindowState]::Minimized
    }

    [void] OnStateChanged()
    {
        if ($this.window.WindowState -eq [System.Windows.WindowState]::Minimized)
        {
            return
        }
        
        $this.isFlashing = $false

        if ($this.skipOnClicked)
        {
            $this.skipOnClicked = $false
        }
        elseif ($this.onClicked)
        {
            $this.onClicked.Invoke()
        }

        $this.window.WindowState = [System.Windows.WindowState]::Minimized
    }

    [void] SetOnClickedFunction($block)
    {
        $this.onClicked = $block
    }

    [void] SetTaskbarItemInfoDescription($text)
    {
        $this.window.TaskbarItemInfo.Description = $text
    }

    [Object] AddThumbButton($thumbButtonSetting)
    {
        $button = New-Object System.Windows.Shell.ThumbButtonInfo
        $button.Description = $thumbButtonSetting.description
        $button.DismissWhenClicked = $true

        $iconPath = $this.settings.GetFullPath($thumbButtonSetting.iconPath)
        if ($iconPath)
        {
            $button.ImageSource = $iconPath
        }

        $this.window.TaskbarItemInfo.ThumbButtonInfos.Add($button)
        return $button
    }

    [void] ShowDialog()
    {
        $this.window.ShowDialog()
    }

    [void] UpdateUnreadCount($unreadCount)
    {
        if ($this.doNotDisturb)
        {
            return
        }
        if ($unreadCount -eq $this.lastUnreadCount)
        {
            return
        }

        if ($this.settings.flashTaskbar.enable)
        {
            if (($unreadCount -gt 0) -and ($unreadCount -gt $this.lastUnreadCount))
            {
                $this.Flash()
            }
            elseif (($unreadCount -eq 0) -and $this.isFlashing)
            {
                $this.ClearFlash()
            }
        }
        $this.lastUnreadCount = $unreadCount

        if ($unreadCount -eq 0)
        {
            $content = ""
        }
        elseif ($unreadCount -lt 0)
        {
            $content = "E"
        }
        else
        {
            $content = [Math]::Min($unreadCount, 99).ToString()
        }
        $this.UpdateOverlayIcon($content)
    }

    [void] UpdateOverlayIcon($content)
    {
        if (-not $content)
        {
            $this.window.TaskbarItemInfo.Overlay = $null
            return
        }

        $dpi = 96
        $iconParameters = $this.GetOverlayIconParameters()
        $iconParameters.Text = $content

        $bitmap = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($iconParameters.IconSize, $iconParameters.IconSize, $dpi, $dpi, [System.Windows.Media.PixelFormats]::Default)
        $rect = New-Object System.Windows.Rect 0, 0, $iconParameters.IconSize, $iconParameters.IconSize
        $control = New-Object System.Windows.Controls.ContentControl
        $control.ContentTemplate = $this.window.Resources["OverlayIcon"]
        $control.Content = [PSCustomObject]$iconParameters
        $control.Arrange($rect)
        $bitmap.Render($control)
        $this.window.TaskbarItemInfo.Overlay = $bitmap
    }

    [Object] GetOverlayIconParameters()
    {
        $parameters = @{}
        $parameters.IconSize = 20.0
        if ($this.settings.overlayIcon.size)
        {
            $parameters.IconSize = $this.settings.overlayIcon.size
        }
        $parameters.FontSize = $parameters.IconSize * 0.7

        $parameters.LineWidth = 1.0
        if ($this.settings.overlayIcon.lineWidth)
        {
            $parameters.LineWidth = $this.settings.overlayIcon.lineWidth
        }

        if ($this.doNotDisturb)
        {
            $parameters.Color = "Gray"
            $parameters.TextColor = "White"
        }
        else
        {
            $parameters.Color = $this.settings.overlayIcon.backgroundColor
            $parameters.TextColor = $this.settings.overlayIcon.textColor
        }

        return $parameters
    }

    [void] ToggleDoNotDisturb()
    {
        $this.doNotDisturb = (-not $this.doNotDisturb)
        if ($this.doNotDisturb)
        {
            $this.UpdateOverlayIcon("D")
            $this.lastUnreadCount = -2 # force update after recovery
            if ($this.isFlashing)
            {
                # call clear twice as a workaround for not cleared issue first time at launch
                $this.ClearFlash()
                $this.ClearFlash()
            }
        }
    }

    [void] ToggleGlobalDoNotDisturb()
    {
        $isGlobalDoNotDisturb = Test-Path $script:globalDoNotDisturbFilePath
        $isGlobalDoNotDisturb = -not $isGlobalDoNotDisturb
        if ($isGlobalDoNotDisturb)
        {
            New-Item $script:globalDoNotDisturbFilePath -Force | Out-Null
        }
        else
        {
            Remove-item $script:globalDoNotDisturbFilePath -Force | Out-Null
        }

        $this.ReferToGlobalDoNotDisturb()
    }

    [void] ReferToGlobalDoNotDisturb()
    {
        $isGlobalDoNotDisturb = Test-Path $script:globalDoNotDisturbFilePath
        if ($isGlobalDoNotDisturb -ne $this.doNotDisturb)
        {
            $this.ToggleDoNotDisturb()
        }
    }

    [void] StartTimerFunction($block, $intervalInSeconds)
    {
        if ($this.timer)
        {
            $this.timer.Stop()
        }
        $this.timer = New-Object System.Windows.Threading.DispatcherTimer
        $this.timer.interval = New-Object TimeSpan(0, 0, $intervalInSeconds)
        $this.timer.add_tick($block)
        $this.timer.Start()
    }

    [void] Flash()
    {
        FlashWindow $this.mainWindowHandle $this.settings.flashTaskbar.rateInMillisecond $this.settings.flashTaskbar.count
        $this.isFlashing = $true
    }

    [void] ClearFlash()
    {
        $this.skipOnClicked = $true
        ShowWindow $this.mainWindowHandle
        $this.isFlashing = $false
    }
}
