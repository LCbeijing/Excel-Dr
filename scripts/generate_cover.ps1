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
$bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    $bgRect,
    [System.Drawing.Color]::FromArgb(255,239,244,255),
    [System.Drawing.Color]::FromArgb(255,236,253,245),
    25
)
$g.FillRectangle($bgBrush, $bgRect)

$glowBlue = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(85,191,219,254))
$glowGreen = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(80,187,247,208))
$g.FillEllipse($glowBlue, 1080, 70, 320, 320)
$g.FillEllipse($glowGreen, 112, 650, 220, 220)

$shellBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(216,255,255,255))
Fill-RoundedRect $g $shellBrush 90 90 1420 720 42

$leftBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    (New-Object System.Drawing.Rectangle 144,144,560,610),
    [System.Drawing.Color]::FromArgb(255,15,23,42),
    [System.Drawing.Color]::FromArgb(255,17,24,39),
    90
)
Fill-RoundedRect $g $leftBrush 144 144 560 610 36

$titleFont = New-Object System.Drawing.Font('Segoe UI', 44, [System.Drawing.FontStyle]::Bold)
$subFont = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Regular)
$h1Font = New-Object System.Drawing.Font('Segoe UI', 30, [System.Drawing.FontStyle]::Bold)
$bodyFont = New-Object System.Drawing.Font('Segoe UI', 17, [System.Drawing.FontStyle]::Regular)
$pillFont = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)

$white = [System.Drawing.Brushes]::White
$blueBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,147,197,253))
$muted = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,203,213,225))

$g.DrawString('Excel-Dr', $titleFont, $white, 190, 230)
$g.DrawString('PRECISION WORKBOOK REPAIR', $subFont, $blueBrush, 194, 300)
$g.DrawString('Repair hidden junk inside Excel / WPS workbooks', $h1Font, $white, 190, 380)

$body1 = 'Built around the real product interface, for finance, operations, order, and warehouse teams.'
$body2 = 'Conservative cleanup for abnormal hidden objects without bluntly breaking formulas or styles.'
$g.DrawString($body1, $bodyFont, $muted, (New-Object System.Drawing.RectangleF(194, 455, 450, 60)))
$g.DrawString($body2, $bodyFont, $muted, (New-Object System.Drawing.RectangleF(194, 510, 450, 80)))

$pill1 = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,37,99,235))
$pill2 = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,22,163,74))
$pill3 = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,245,158,11))
Fill-RoundedRect $g $pill1 194 632 166 58 16
Fill-RoundedRect $g $pill2 376 632 152 58 16
Fill-RoundedRect $g $pill3 544 632 116 58 16
$g.DrawString('Portable EXE', $pillFont, $white, 222, 648)
$g.DrawString('Real UI', $pillFont, $white, 414, 648)
$g.DrawString('Open', $pillFont, [System.Drawing.Brushes]::Black, 573, 648)

$cardShadow = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(28,15,23,42))
Fill-RoundedRect $g $cardShadow 814 190 660 510 34
$cardBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,248,250,252))
Fill-RoundedRect $g $cardBrush 796 172 660 510 34

$shot = [System.Drawing.Image]::FromFile($screenshot)
$shotRect = New-Object System.Drawing.Rectangle 826,202,600,450
$clip = New-RoundedRectPath 826 202 600 450 24
$oldClip = $g.Clip
$g.SetClip($clip)
$g.DrawImage($shot, $shotRect)
$g.Clip = $oldClip
$clip.Dispose()

$overlayBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(26,255,255,255))
Fill-RoundedRect $g $overlayBrush 826 202 600 450 24
$framePen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255,203,213,225), 2)
Draw-RoundedRect $g $framePen 826 202 600 450 24

$labelBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,15,23,42))
Fill-RoundedRect $g $labelBrush 1018 692 220 54 18
$labelFont = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$g.DrawString('Actual Product UI', $labelFont, $white, 1048, 707)

$shot.Dispose()
$bgBrush.Dispose()
$glowBlue.Dispose()
$glowGreen.Dispose()
$shellBrush.Dispose()
$leftBrush.Dispose()
$blueBrush.Dispose()
$muted.Dispose()
$pill1.Dispose()
$pill2.Dispose()
$pill3.Dispose()
$cardShadow.Dispose()
$cardBrush.Dispose()
$overlayBrush.Dispose()
$framePen.Dispose()
$labelBrush.Dispose()
$titleFont.Dispose()
$subFont.Dispose()
$h1Font.Dispose()
$bodyFont.Dispose()
$pillFont.Dispose()
$labelFont.Dispose()

$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose()
$bmp.Dispose()

Write-Output $out
