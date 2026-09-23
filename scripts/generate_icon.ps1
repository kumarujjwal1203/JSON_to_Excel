Add-Type -AssemblyName System.Drawing

$srcPath = 'C:\Users\ADMIN\.gemini\antigravity\brain\efdd7786-082b-4c52-97fb-eaea07742ec0\.user_uploaded\media_1790008771607.jpg'
$icoPath = 'c:\Users\ADMIN\Videos\GSTIN\src\GSTJsonToExcel\app_icon.ico'
$pngPath = 'c:\Users\ADMIN\Videos\GSTIN\src\GSTJsonToExcel\app_logo.png'

if (-not (Test-Path $srcPath)) {
    Write-Error "Source image not found: $srcPath"
    exit 1
}

$source = [System.Drawing.Bitmap]::FromFile($srcPath)
$sizes = @(256, 128, 64, 48, 32, 16)
$pngDataList = [System.Collections.Generic.List[psobject]]::new()

foreach ($size in $sizes) {
    $bmp = New-Object System.Drawing.Bitmap $size, $size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.DrawImage($source, 0, 0, $size, $size)
    $g.Dispose()

    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bytes = $ms.ToArray()
    $ms.Dispose()
    $bmp.Dispose()

    $pngDataList.Add([pscustomobject]@{
        Size = $size
        Bytes = $bytes
    })
}

$fs = [System.IO.File]::Create($icoPath)
$bw = New-Object System.IO.BinaryWriter $fs

$bw.Write([uint16]0)
$bw.Write([uint16]1)
$bw.Write([uint16]$pngDataList.Count)

$dataOffset = 6 + (16 * $pngDataList.Count)

foreach ($item in $pngDataList) {
    $w = if ($item.Size -ge 256) { [byte]0 } else { [byte]$item.Size }
    $h = if ($item.Size -ge 256) { [byte]0 } else { [byte]$item.Size }
    $bw.Write($w)
    $bw.Write($h)
    $bw.Write([byte]0)
    $bw.Write([byte]0)
    $bw.Write([uint16]1)
    $bw.Write([uint16]32)
    $bw.Write([uint32]$item.Bytes.Length)
    $bw.Write([uint32]$dataOffset)
    $dataOffset += $item.Bytes.Length
}

foreach ($item in $pngDataList) {
    $bw.Write($item.Bytes)
}

$bw.Flush()
$bw.Close()
$fs.Close()

# Also save 512x512 PNG
$logoBmp = New-Object System.Drawing.Bitmap 512, 512
$gLogo = [System.Drawing.Graphics]::FromImage($logoBmp)
$gLogo.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$gLogo.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$gLogo.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$gLogo.Clear([System.Drawing.Color]::Transparent)
$gLogo.DrawImage($source, 0, 0, 512, 512)
$gLogo.Dispose()
$logoBmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
$logoBmp.Dispose()

$source.Dispose()

# Copy to installer directory as well
Copy-Item $icoPath 'c:\Users\ADMIN\Videos\GSTIN\installer\app_icon.ico' -Force
Copy-Item $pngPath 'c:\Users\ADMIN\Videos\GSTIN\installer\app_logo.png' -Force

Get-Item $icoPath, $pngPath, 'c:\Users\ADMIN\Videos\GSTIN\installer\app_icon.ico' | Select-Object FullName, Length
