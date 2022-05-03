Add-Type -AssemblyName System.Windows.Forms

function CreateUi()
{
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Available Folder Paths"
    $form.Width = 400
    $form.Height = 400
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $table = New-Object System.Windows.Forms.TableLayoutPanel
    $table.ColumnCount = 1
    $table.RowCount = 2
    $table.Dock = "Fill"
    $form.Controls.Add($table)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Dock = "Fill"
    $listBox.IntegralHeight = $false
    $listbox.Add_SelectedValueChanged({$button.Text = "Copy Path to Clipboard"})
    $table.Controls.Add($listBox, 0, 0)
    $table.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Copy Path to Clipboard"
    $button.Anchor = "Left, Right"
    $button.Margin = 10
    $button.BackColor = [System.Drawing.Color]::DarkSeaGreen
    $button.Add_Click({
        Set-Clipboard -Value ('"{0}"' -f $listBox.SelectedItem)
        if ($listBox.SelectedItem)
        {
            $button.Text = "Copied!"
        }
    })
    $table.Controls.Add($button, 1, 0)
    $table.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

    $form, $listBox, $button
}

function IsTemporaryFolder($folder)
{
    return $folder.Name.Contains("MS-OLK")
}

function AddToListBox($folders)
{
    foreach ($folder in $folders)
    {
        if (-not (IsTemporaryFolder $folder))
        {
            $listBox.Items.Add($folder.FolderPath) | out-null
            AddToListBox $folder.Folders
        }
    }
}

$form, $listBox, $button = CreateUi

$outlook = New-Object -ComObject Outlook.Application
if (-not $outlook.Name)
{
    Read-Host "Failed to open Outlook. Press any key to close"
    return
}

$namespace = $outlook.GetNamespace("MAPI")
AddToListBox $namespace.Folders

foreach ($store in $namespace.Stores)
{
    try
    {
        $searchFolders = $store.GetSearchFolders()
    }
    catch
    {
        # catch access rights errors
        Write-Host $PSItem
        continue
    }
    AddToListBox $searchFolders
}

$form.ShowDialog() | Out-Null

[system.runtime.interopservices.marshal]::releasecomobject($outlook) | Out-Null

