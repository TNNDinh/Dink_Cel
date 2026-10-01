Add-Type -AssemblyName System.Drawing

$sourcePath = Join-Path $PSScriptRoot 'icon_dink_cell.png'
$iconPath = Join-Path $PSScriptRoot 'DinkCel.ico'
$sizes = @(16, 24, 32, 48, 64, 128, 256)
$frames = New-Object 'System.Collections.Generic.List[byte[]]'
$source = [System.Drawing.Image]::FromFile($sourcePath)

try {
    foreach ($size in $sizes) {
        $bitmap = [System.Drawing.Bitmap]::new($size, $size,
            [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $memory = [System.IO.MemoryStream]::new()
        $frameWriter = [System.IO.BinaryWriter]::new($memory)
        try {
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.DrawImage($source, 0, 0, $size, $size)

            $maskStride = [int]([Math]::Ceiling($size / 32.0) * 4)
            $frameWriter.Write([uint32]40)
            $frameWriter.Write([int32]$size)
            $frameWriter.Write([int32]($size * 2))
            $frameWriter.Write([uint16]1)
            $frameWriter.Write([uint16]32)
            $frameWriter.Write([uint32]0)
            $frameWriter.Write([uint32]($size * $size * 4 + $maskStride * $size))
            $frameWriter.Write([int32]0)
            $frameWriter.Write([int32]0)
            $frameWriter.Write([uint32]0)
            $frameWriter.Write([uint32]0)

            $rectangle = [System.Drawing.Rectangle]::new(0, 0, $size, $size)
            $bits = $bitmap.LockBits($rectangle,
                [System.Drawing.Imaging.ImageLockMode]::ReadOnly,
                [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
            try {
                $rowBytes = New-Object byte[] ($size * 4)
                for ($y = $size - 1; $y -ge 0; $y--) {
                    $pointer = [System.IntPtr]::Add($bits.Scan0, $y * $bits.Stride)
                    [System.Runtime.InteropServices.Marshal]::Copy(
                        $pointer, $rowBytes, 0, $rowBytes.Length)
                    $frameWriter.Write($rowBytes)
                }
            }
            finally {
                $bitmap.UnlockBits($bits)
            }
            $frameWriter.Write((New-Object byte[] ($maskStride * $size)))
            $frameWriter.Flush()
            $frames.Add($memory.ToArray())
        }
        finally {
            $frameWriter.Dispose()
            $memory.Dispose()
            $graphics.Dispose()
            $bitmap.Dispose()
        }
    }
}
finally {
    $source.Dispose()
}

$stream = [System.IO.File]::Create($iconPath)
$writer = [System.IO.BinaryWriter]::new($stream)
try {
    $writer.Write([uint16]0)
    $writer.Write([uint16]1)
    $writer.Write([uint16]$sizes.Count)
    $offset = 6 + 16 * $sizes.Count
    for ($i = 0; $i -lt $sizes.Count; $i++) {
        $writer.Write([byte]($sizes[$i] -band 255))
        $writer.Write([byte]($sizes[$i] -band 255))
        $writer.Write([byte]0)
        $writer.Write([byte]0)
        $writer.Write([uint16]1)
        $writer.Write([uint16]32)
        $writer.Write([uint32]$frames[$i].Length)
        $writer.Write([uint32]$offset)
        $offset += $frames[$i].Length
    }
    foreach ($frame in $frames) {
        $writer.Write($frame)
    }
}
finally {
    $writer.Dispose()
}

Write-Output "Created $iconPath"
