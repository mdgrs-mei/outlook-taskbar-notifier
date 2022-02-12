# process.ps1, settings_helper.ps1 must be included before executing this file.
# "Add-Type -AssemblyName PresentationFramework" must be called before including this file.

. (Join-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) "flash_window.ps1")

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

        $iconPath = GetFullPathFromSettingsRelativePath $settings $settings.iconPath
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

        $iconPath = GetFullPathFromSettingsRelativePath $this.settings $thumbButtonSetting.iconPath
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

        $iconSize = $this.window.Resources["OverlayIconSize"]
        $dpi = 96
        if ($this.doNotDisturb)
        {
            $backgroundColor = "Gray"
            $textColor = "White"
        }
        else
        {
            $backgroundColor = $this.settings.overlayIcon.backgroundColor
            $textColor = $this.settings.overlayIcon.textColor
        }

        $bitmap = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($iconSize, $iconSize, $dpi, $dpi, [System.Windows.Media.PixelFormats]::Default)
        $rect = New-Object System.Windows.Rect 0, 0, $iconSize, $iconSize
        $control = New-Object System.Windows.Controls.ContentControl
        $control.ContentTemplate = $this.window.Resources["OverlayIcon"]
        $control.content = [PSCustomObject]@{
            Color = $backgroundColor
            TextColor = $textColor
            Text = $content
        }
        $control.Arrange($rect)
        $bitmap.Render($control)
        $this.window.TaskbarItemInfo.Overlay = $bitmap
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
