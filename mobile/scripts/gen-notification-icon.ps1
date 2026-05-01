$ErrorActionPreference = 'Stop'
$out = Join-Path (Split-Path $PSScriptRoot -Parent) 'assets\notification-icon.png'
Add-Type -AssemblyName System.Drawing
$bmp = New-Object System.Drawing.Bitmap 96, 96
$bmp.MakeTransparent()
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.Clear([System.Drawing.Color]::Transparent)
$w = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
$g.FillEllipse($w, 14, 14, 68, 68)
$g.Dispose()
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Host "Written $out"
