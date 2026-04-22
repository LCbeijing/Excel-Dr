$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

$repo = Split-Path -Parent $PSScriptRoot
$screenshot = Join-Path $repo 'docs\images\app-screenshot.png'
$out = Join-Path $repo 'docs\images\cover.png'

if (-not (Test-Path $screenshot)) {
    throw "Screenshot not found: $screenshot"
}

$canvasW = 1600
$canvasH = 900
$bmp = New-Object System.Drawing.Bitmap $canvasW, $canvasH
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

function New-RoundedRectPath([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $r * 2
    $path.AddArc($x, $y, $d, $d, 180, 90)
    $path.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $path.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $path.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

function Fill-RoundedRect($graphics, $brush, [float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $path = New-RoundedRectPath $x $y $w $h $r
    $graphics.FillPath($brush, $path)
    $path.Dispose()
}

function Draw-RoundedRect($graphics, $pen, [float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $path = New-RoundedRectPath $x $y $w $h $r
    $graphics.DrawPath($pen, $path)
    $path.Dispose()
}

$bgRect = New-Object System.Drawing.Rectangle 0,0,$canvasW,$canvasH
$bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($bgRect,[System.Drawing.Color]::FromArgb(255,239,244,255),[System.Drawing.Color]::FromArgb(255,236,253,245),18)
$g.FillRectangle($bgBrush, $bgRect)

$glowBlue = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(85,191,219,254))
$glowGreen = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(80,187,247,208))
$g.FillEllipse($glowBlue, 1180, 38, 250, 250)
$g.FillEllipse($glowGreen, 18, 760, 120, 120)

$shellBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(220,255,255,255))
Fill-RoundedRect $g $shellBrush 54 78 1492 744 40

$titleFont = New-Object System.Drawing.Font('Segoe UI', 52, [System.Drawing.FontStyle]::Bold)
$subFont = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Regular)
$h1Font = New-Object System.Drawing.Font('Segoe UI', 28, [System.Drawing.FontStyle]::Bold)
$bodyFont = New-Object System.Drawing.Font('Segoe UI', 17, [System.Drawing.FontStyle]::Regular)
$metaFont = New-Object System.Drawing.Font('Consolas', 13, [System.Drawing.FontStyle]::Regular)

$white = [System.Drawing.Brushes]::White
$darkBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,15,23,42))
$blueBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,147,197,253))
$muted = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,71,85,105))
$accentBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,59,130,246))

$g.FillRectangle($accentBrush, 118, 204, 10, 236)
$g.DrawString('Excel-Dr', $titleFont, $darkBrush, 152, 206)
$g.DrawString('PRECISION WORKBOOK REPAIR', $subFont, $blueBrush, 156, 286)
$g.DrawString('Repair hidden', $h1Font, $darkBrush, 152, 394)
$g.DrawString('workbook junk', $h1Font, $darkBrush, 152, 436)

$body1 = 'Real product UI'
$body2 = 'Portable Windows EXE'
$body3 = 'Safe cleanup for hidden drawing junk'
$g.DrawString($body1, $bodyFont, $muted, 156, 530)
$g.DrawString($body2, $bodyFont, $muted, 156, 572)
$g.DrawString($body3, $bodyFont, $muted, (New-Object System.Drawing.RectangleF(156, 618, 250, 64)))
$g.DrawString('FORMULAS / VALID IMAGES PRESERVED', $metaFont, $blueBrush, 156, 704)

$cardShadow = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30,15,23,42))
Fill-RoundedRect $g $cardShadow 572 124 966 652 30
$cardBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,248,250,252))
Fill-RoundedRect $g $cardBrush 552 104 966 652 30

$shot = [System.Drawing.Image]::FromFile($screenshot)
$shotRect = New-Object System.Drawing.Rectangle 580,132,920,604
$clip = New-RoundedRectPath 580 132 920 604 24
$oldClip = $g.Clip
$g.SetClip($clip)
$g.DrawImage($shot, $shotRect)
$g.Clip = $oldClip
$clip.Dispose()

$overlayBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(14,255,255,255))
Fill-RoundedRect $g $overlayBrush 580 132 920 604 24
$framePen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255,203,213,225), 2)
Draw-RoundedRect $g $framePen 580 132 920 604 24

$labelBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,15,23,42))
Fill-RoundedRect $g $labelBrush 920 748 220 50 18
$labelFont = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$g.DrawString('Real Product UI', $labelFont, $white, 962, 762)

$shot.Dispose()
$bgBrush.Dispose()
$glowBlue.Dispose()
$glowGreen.Dispose()
$shellBrush.Dispose()
$darkBrush.Dispose()
$blueBrush.Dispose()
$muted.Dispose()
$accentBrush.Dispose()
$cardShadow.Dispose()
$cardBrush.Dispose()
$overlayBrush.Dispose()
$framePen.Dispose()
$labelBrush.Dispose()
$titleFont.Dispose()
$subFont.Dispose()
$h1Font.Dispose()
$bodyFont.Dispose()
$metaFont.Dispose()
$labelFont.Dispose()

$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose()
$bmp.Dispose()

Write-Output $out
