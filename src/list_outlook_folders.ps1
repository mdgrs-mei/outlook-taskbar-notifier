$outlook = New-Object -ComObject Outlook.Application
if (-not $outlook.Name)
{
    "Failed to open Outlook."
    return
}

function IsTemporaryFolder($folder)
{
    return $folder.Name.Contains("MS-OLK")
}

function PrintFolders($folders)
{
    foreach ($folder in $folders)
    {
        if (-not (IsTemporaryFolder $folder))
        {
            $folder.FolderPath
            PrintFolders $folder.Folders
        }
    }
}

$namespace = $outlook.GetNamespace("MAPI")
PrintFolders $namespace.Folders

foreach ($store in $namespace.Stores)
{
    $searchFolders = $store.GetSearchFolders()
    PrintFolders $searchFolders
}

[system.runtime.interopservices.marshal]::releasecomobject($outlook) | Out-Null
