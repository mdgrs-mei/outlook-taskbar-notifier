$outlook = New-Object -ComObject Outlook.Application
if (-not $outlook.Name)
{
    "Failed to open Outlook."
    return
}

function PrintFolders($folders)
{
    foreach ($folder in $folders)
    {
        # Exclude the root folders
        if ($folder.Parent.GetType().Name -ne "NameSpaceClass")
        {
            $folder.FolderPath
        }
        PrintFolders $folder.Folders
    }
}

$namespace = $outlook.GetNamespace("MAPI")
PrintFolders $namespace.Folders

[system.runtime.interopservices.marshal]::releasecomobject($outlook) | Out-Null