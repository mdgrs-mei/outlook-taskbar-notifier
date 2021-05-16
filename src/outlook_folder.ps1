# process.ps1 must be included before executing this file.

class OutlookFolder
{
    $folderPath
    $outlookExePath
    $outlook
    $folder

    [String] Init($folderPath, $outlookExePath)
    {
        $this.folderPath = $folderPath
        $this.outlookExePath = $outlookExePath
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
        $this.folder = [OutlookFolder]::FindFolder($namespace.Folders, $this.folderPath)

        if (-not $this.IsFolderValid())
        {
            if (-not $this.folder)
            {
                return "Failed to find folder [{0}]." -f $this.folderPath
            }
            if (-not $this.folder.Parent)
            {
                return "Root folder is not supported [{0}]." -f $this.folderPath
            }
            return "Folder is not valid [{0}]." -f $this.folderPath
        }
        return ""
    }

    static [Object] FindFolder($folders, $folderPath)
    {
        foreach ($folder in $folders)
        {
            if ($folder.FolderPath.ToString() -eq $folderPath)
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
        return $this.folder.Name -and $this.folder.Parent
    }

    [String] GetName()
    {
        return $this.folder.Name
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
            # get folder again to refresh UnreadItemCount
            $this.folder = [OutlookFolder]::FindFolder($this.folder.Parent.Folders, $this.folderPath)
            if (-not $this.IsFolderValid())
            {
                return $errorUnreadCount
            }
            return $this.folder.UnreadItemCount
        }
        catch
        {
            Write-Host "GetUnreadCount failed. [$PSItem]"
            return $errorUnreadCount
        }
    }

    [void] MarkAllAsRead()
    {
        if (-not $this.IsFolderValid())
        {
            return
        }

        try 
        {
            $items = $this.folder.Items.Restrict("[UnRead] = True")
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

        $folderPathArg = "outlook:" + $this.folderPath
        $folderPathArg = '"' + $folderPathArg + '"'
        Start-Process $this.outlookExePath -Wait -ArgumentList "/recycle", "/select", $folderPathArg
        FocusApp "outlook.exe"
    }
}