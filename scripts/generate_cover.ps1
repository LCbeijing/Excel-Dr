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

$leftBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    (New-Object System.Drawing.Rectangle 110,132,390,610),
    [System.Drawing.Color]::FromArgb(255,15,23,42),
    [System.Drawing.Color]::FromArgb(255,17,24,39),
    90
)
Fill-RoundedRect $g $leftBrush 110 132 390 610 32

$titleFont = New-Object System.Drawing.Font('Segoe UI', 46, [System.Drawing.FontStyle]::Bold)
$subFont = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Regular)
$h1Font = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold)
$bodyFont = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Regular)
$pillFont = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)

$white = [System.Drawing.Brushes]::White
$blueBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,147,197,253))
$muted = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,203,213,225))

$g.DrawString('Excel-Dr', $titleFont, $white, 150, 210)
$g.DrawString('PRECISION WORKBOOK REPAIR', $subFont, $blueBrush, 154, 282)
$g.DrawString('Fix workbook junk', $h1Font, $white, 154, 362)

$body1 = 'Real product UI'
$body2 = 'Portable Windows EXE'
$body3 = 'Keeps formulas, styles, and valid images intact'
$body4 = 'Targets hidden drawing junk only'
$g.DrawString($body1, $bodyFont, $muted, 154, 440)
$g.DrawString($body2, $bodyFont, $muted, 154, 476)
$g.DrawString($body3, $bodyFont, $muted, (New-Object System.Drawing.RectangleF(154, 520, 300, 56)))
$g.DrawString($body4, $bodyFont, $muted, (New-Object System.Drawing.RectangleF(154, 586, 300, 32)))

$pill1 = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,37,99,235))
$pill2 = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,22,163,74))
$pill3 = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,245,158,11))
Fill-RoundedRect $g $pill1 154 644 136 54 16
Fill-RoundedRect $g $pill2 302 644 128 54 16
Fill-RoundedRect $g $pill3 442 644 62 54 16
$g.DrawString('Portable', $pillFont, $white, 180, 658)
$g.DrawString('Real UI', $pillFont, $white, 328, 658)
$g.DrawString('EXE', $pillFont, [System.Drawing.Brushes]::Black, 456, 658)

$cardShadow = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30,15,23,42))
Fill-RoundedRect $g $cardShadow 542 124 996 652 30
$cardBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,248,250,252))
Fill-RoundedRect $g $cardBrush 522 104 996 652 30

$shot = [System.Drawing.Image]::FromFile($screenshot)
$shotRect = New-Object System.Drawing.Rectangle 550,132,940,604
$clip = New-RoundedRectPath 550 132 940 604 24
$oldClip = $g.Clip
$g.SetClip($clip)
$g.DrawImage($shot, $shotRect)
$g.Clip = $oldClip
$clip.Dispose()

$overlayBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(14,255,255,255))
Fill-RoundedRect $g $overlayBrush 550 132 940 604 24
$framePen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255,203,213,225), 2)
Draw-RoundedRect $g $framePen 550 132 940 604 24

$labelBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,15,23,42))
Fill-RoundedRect $g $labelBrush 920 748 220 50 18
$labelFont = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$g.DrawString('Real Product UI', $labelFont, $white, 962, 762)

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
