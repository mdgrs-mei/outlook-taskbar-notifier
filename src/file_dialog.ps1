Add-Type -AssemblyName System.Windows.Forms

function GetInputFilePathWithDialog($title, $filter, $dir)
{
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $title
    $dialog.Filter = $filter
    $dialog.InitialDirectory = $dir
    $ret = $dialog.ShowDialog()
    if ($ret -eq [System.Windows.Forms.DialogResult]::OK)
    {
        return $dialog.FileName
    }
    return ""
}

function GetOutputFilePathWithDialog($title, $ext, $filter, $dir)
{
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = $title
    $dialog.DefaultExt = $ext
    $dialog.Filter = $filter
    $dialog.InitialDirectory = $dir
    $ret = $dialog.ShowDialog()
    if ($ret -eq [System.Windows.Forms.DialogResult]::OK)
    {
        return $dialog.FileName
    }
    return ""
}
