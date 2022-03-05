# process.ps1 must be included before executing this file.

class OutlookFolder
{
    $folderPath
    $folderName
    $outlook
    $folder

    [String] Init($folderPath)
    {
        $this.folderPath = $folderPath
        $this.folderName = $folderPath.Substring($folderPath.LastIndexOf("\")+1)
        return $this.InitOutlook()
    }

    [void] Term()
    {
        if (-not $this.IsOutlookValid())
        {
            return
        }
        [system.runtime.interopservices.marshal]::releasecomobject($this.outlook)
    }

    [String] InitOutlook()
    {
        if (-not $this.IsOutlookValid())
        {
            $this.outlook = New-Object -ComObject Outlook.Application
            if (-not $this.IsOutlookValid())
            {
                return "Failed to get Outlook."
            }
        }

        $namespace = $this.outlook.GetNamespace("MAPI")

        # find the folder in normal folders first
        $this.folder = [OutlookFolder]::FindFolder($namespace.Folders, $this.folderPath)

        # find the folder in Search Folders if failed
        if (-not $this.IsFolderValid())
        {
            foreach ($store in $namespace.Stores)
            {
                $searchFolders = $store.GetSearchFolders()
                $this.folder = [OutlookFolder]::FindFolder($searchFolders, $this.folderPath)
                if ($this.IsFolderValid())
                {
                    break
                }
            }
        }

        if (-not $this.IsFolderValid())
        {
            if (-not $this.folder)
            {
                return "Failed to find folder [{0}]." -f $this.folderPath
            }
            return "Folder is not valid [{0}]." -f $this.folderPath
        }
        return ""
    }

    static [Object] FindFolder($folders, $folderPath)
    {
        foreach ($folder in $folders)
        {
            if ($folder.FolderPath -and ($folder.FolderPath.ToString() -eq $folderPath))
            {
                return $folder
            }

            $f = [OutlookFolder]::FindFolder($folder.Folders, $folderPath)
            if ($f)
            {
                return $f
            }
        }
        return $null
    }

    [boolean] IsOutlookValid()
    {
        return $this.outlook -and $this.outlook.Name
    }

    [boolean] IsFolderValid()
    {
        return $this.folder.Name
    }

    [String] GetName()
    {
        return $this.folderName
    }

    [String] InitOutlookIfNotValid()
    {
        if ((-not $this.IsOutlookValid()) -or (-not $this.IsFolderValid()))
        {
            return $this.InitOutlook()
        }
        return ""
    }

    [int] GetUnreadCount()
    {
        $errorUnreadCount = -1
        if (-not $this.IsFolderValid())
        {
            return $errorUnreadCount
        }

        try 
        {
            $items = $this.GetUnreadItems()
            if (-not $items)
            {
                return 0;
            }
            return $items.Count
        }
        catch
        {
            Write-Host "GetUnreadCount failed. [$PSItem]"
            return $errorUnreadCount
        }
    }

    [Object] GetUnreadItems()
    {
        if (-not $this.IsFolderValid())
        {
            return $null
        }
        return $this.folder.Items.Restrict("[UnRead] = True")
    }

    [String] GetUnreadItemsSummary($maxItemCount, $maxItemCharacterCount)
    {
        $summary = ""
        try 
        {
            $items = $this.GetUnreadItems()
            if (-not $items)
            {
                return $summary
            }
            $items.Sort("[ReceivedTime]", $true)

            $titleMaxLength = $maxItemCharacterCount
            for ($i = 0; $i -lt $items.Count; $i++)
            {
                if ($i -eq $maxItemCount)
                {
                    $summary += "..."
                    break
                }

                $item = $items.Item($i+1)
                $title = $item.Subject
                if ($title.Length -gt $titleMaxLength)
                {
                    $title = $title.SubString(0, $titleMaxLength) + "..."
                }
                $itemStr = "{0}`n" -f $title
                $summary += $itemStr
            }
        }
        catch
        {
            Write-Host "GetUnreadItemsSummary failed. [$PSItem]"
        }
        return $summary
    }

    [void] MarkAllAsRead()
    {
        try 
        {
            $items = $this.GetUnreadItems()
            if (-not $items)
            {
                return
            }
            for ($i = $items.Count; $i -gt 0; --$i)
            {
                $items[$i].Unread = $false
            }
        }
        catch
        {
            Write-Host "MarkAllAsRead failed. [$PSItem]"
        }
    }

    [void] Focus()
    {
        if (-not $this.IsFolderValid())
        {
            return
        }

        try 
        {
            $explorer = $this.outlook.ActiveExplorer()
            if ($explorer)
            {
                $explorer.Activate()
                $explorer.CurrentFolder = $this.folder

                $explorer.ClearSearch()
                $explorer.ClearSelection()
                $view = $explorer.CurrentView
                if ($view)
                {
                    # Reset the selection to top.
                    $view.Apply()
                }
            }
            else
            {
                $this.folder.Display()
            }

            FocusApp "outlook.exe"
        }
        catch
        {
            Write-Host "Focus on folder failed. [$PSItem]"
        }
    }

    [boolean] OpenNewestUnread()
    {
        try 
        {
            $items = $this.GetUnreadItems()
            if (-not $items)
            {
                return $false
            }
            if ($items.Count -eq 0)
            {
                return $false
            }
            $items.Sort("[ReceivedTime]")
            $items[$items.Count].Display()
            FocusApp "outlook.exe"
            return $true
        }
        catch
        {
            Write-Host "OpenNewestUnread failed. [$PSItem]"
            return $false
        }
    }

    [boolean] OpenOldestUnread()
    {
        try 
        {
            $items = $this.GetUnreadItems()
            if (-not $items)
            {
                return $false
            }
            if ($items.Count -eq 0)
            {
                return $false
            }
            $items.Sort("[ReceivedTime]")
            $items[1].Display()
            FocusApp "outlook.exe"
            return $true
        }
        catch
        {
            Write-Host "OpenOldestUnread failed. [$PSItem]"
            return $false
        }
    }
}