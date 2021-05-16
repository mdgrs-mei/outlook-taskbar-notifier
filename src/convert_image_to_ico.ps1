Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$scriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$rootDir = Split-Path $scriptDir -Parent

$sourceImagePath = $args[0]
$iconPath = $args[1]

if ((-not $sourceImagePath) -or (-not $iconPath))
{
    . (Join-Path $scriptDir "file_dialog.ps1")

    $sourceImagePath = GetInputFilePathWithDialog "Select an image" "Image Files|*.png;*.bmp;*.tif;*.gif;*.jpg;*.tiff" $rootDir
    if (-not $sourceImagePath)
    {
        return
    }
    $sourceImageDir = Split-Path $sourceImagePath -Parent
    $iconPath = GetOutputFilePathWithDialog "Specify output ico path" ".ico" "Icon Files|*.ico" $sourceImageDir
    if (-not $iconPath)
    {
        return
    }
}

# https://docs.microsoft.com/en-us/previous-versions/ms997538(v=msdn.10)?redirectedfrom=MSDN
function CreateIconFromImage($image)
{
    $memoryStream = New-Object System.IO.MemoryStream
    $writer = New-Object System.IO.BinaryWriter($memoryStream);

    # ICONDIR
    $writer.Write([Int16]0) # reserved
    $writer.Write([Int16]1) # resource type 1 for icons
    $writer.Write([Int16]1); # number of images

    # ICONDIRENTRY
    $w = $image.Width
    if ($w -ge 256)
    {
        $w = 0 # 0 means 256 pixel
    }
    $writer.Write([Byte]$w)
    $h = $image.Height;
    if ($h -ge 256)
    {
        $h = 0
    }
    $writer.Write([Byte]$h)
    $writer.Write([Byte]0) #Number of colors in image (0 if >=8bpp)
    $writer.Write([Byte]0) # Reserved ( must be 0)
    $writer.Write([Int16]0) # Color Planes
    $writer.Write([Int16]0) # Bits per pixel
    $imageSizePos = $memoryStream.Position

    $writer.Write([Int]0) # How many bytes in this resource?
    $imageStart = [Int]$memoryStream.Position + 4
    $writer.Write([Int]$imageStart) # Where in the file is this image?

    # Image data
    $image.Save($memoryStream, [System.Drawing.Imaging.ImageFormat]::Png)
    $imageSize = [Int]$memoryStream.Position - $imageStart
    $memoryStream.Seek($imageSizePos, [System.IO.SeekOrigin]::Begin) | Out-Null
    $writer.Write([Int]$imageSize)
    $memoryStream.Seek(0, [System.IO.SeekOrigin]::Begin) | Out-Null

    return New-Object System.Drawing.Icon($memoryStream);
}

$sourceImage = [Drawing.Image]::FromFile($sourceImagePath)
$icon = CreateIconFromImage $sourceImage
$fileStream = [IO.File]::Create($iconPath)
$icon.Save($fileStream) 
$fileStream.Close()
$icon.Dispose()
