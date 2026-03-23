Add-Type -AssemblyName System.Drawing

$blue = [System.Drawing.Color]::FromArgb(21, 100, 192)
$green = [System.Drawing.Color]::FromArgb(74, 162, 81)
$sizes = @(16, 32, 80)
$assetsDir = "$PSScriptRoot\assets"

function New-Icon($name, $drawFn) {
    foreach ($s in $sizes) {
        $bmp = New-Object System.Drawing.Bitmap($s, $s)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.Clear([System.Drawing.Color]::Transparent)
        & $drawFn $g $s
        $g.Dispose()
        $path = "$assetsDir\$name-$s.png"
        $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
        Write-Host "  $path"
    }
}

# --- Remove Images: blue photo frame with green X ---
Write-Host "Remove Images..."
New-Icon "icon-remove-images" {
    param($g, $s)
    $p = $s / 80.0
    $pen = New-Object System.Drawing.Pen($blue, [Math]::Max(2, 3 * $p))
    $rect = New-Object System.Drawing.RectangleF((8*$p), (8*$p), (48*$p), (40*$p))
    $g.DrawRectangle($pen, $rect.X, $rect.Y, $rect.Width, $rect.Height)
    # mountain inside
    $points = @(
        (New-Object System.Drawing.PointF((12*$p), (42*$p))),
        (New-Object System.Drawing.PointF((24*$p), (28*$p))),
        (New-Object System.Drawing.PointF((34*$p), (36*$p))),
        (New-Object System.Drawing.PointF((42*$p), (24*$p))),
        (New-Object System.Drawing.PointF((52*$p), (42*$p)))
    )
    $thinPen = New-Object System.Drawing.Pen($blue, [Math]::Max(1, 2 * $p))
    $g.DrawLines($thinPen, $points)
    $thinPen.Dispose()
    $pen.Dispose()
    # Green X overlay (bottom-right)
    $xPen = New-Object System.Drawing.Pen($green, [Math]::Max(2, 4 * $p))
    $xPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $xPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $g.DrawLine($xPen, (48*$p), (50*$p), (72*$p), (74*$p))
    $g.DrawLine($xPen, (72*$p), (50*$p), (48*$p), (74*$p))
    $xPen.Dispose()
}

# --- Remove Attachments: blue paperclip with green X ---
Write-Host "Remove Attachments..."
New-Icon "icon-remove-attachments" {
    param($g, $s)
    $p = $s / 80.0
    $pen = New-Object System.Drawing.Pen($blue, [Math]::Max(2, 3 * $p))
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    # Paperclip shape using arcs and lines
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc((18*$p), (6*$p), (24*$p), (24*$p), 180, 180)
    $path.AddLine((42*$p), (18*$p), (42*$p), (50*$p))
    $path.AddArc((24*$p), (38*$p), (18*$p), (18*$p), 0, 180)
    $path.AddLine((24*$p), (50*$p), (24*$p), (24*$p))
    $path.AddArc((24*$p), (14*$p), (12*$p), (12*$p), 180, -180)
    $path.AddLine((36*$p), (20*$p), (36*$p), (44*$p))
    $g.DrawPath($pen, $path)
    $path.Dispose()
    $pen.Dispose()
    # Green X (bottom-right)
    $xPen = New-Object System.Drawing.Pen($green, [Math]::Max(2, 4 * $p))
    $xPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $xPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $g.DrawLine($xPen, (48*$p), (50*$p), (72*$p), (74*$p))
    $g.DrawLine($xPen, (72*$p), (50*$p), (48*$p), (74*$p))
    $xPen.Dispose()
}

# --- Keep 2 Replies: two overlapping speech bubbles ---
Write-Host "Keep Replies..."
New-Icon "icon-keep-replies" {
    param($g, $s)
    $p = $s / 80.0
    # Back bubble (blue, filled)
    $brush1 = New-Object System.Drawing.SolidBrush($blue)
    $path1 = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path1.AddArc((22*$p), (6*$p), (52*$p), (40*$p), 180, 360)
    $path1.AddLine((56*$p), (46*$p), (62*$p), (58*$p))
    $path1.AddLine((62*$p), (58*$p), (50*$p), (46*$p))
    $path1.CloseFigure()
    $g.FillPath($brush1, $path1)
    $path1.Dispose()
    $brush1.Dispose()
    # Front bubble (green, filled)
    $brush2 = New-Object System.Drawing.SolidBrush($green)
    $path2 = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path2.AddArc((6*$p), (20*$p), (52*$p), (40*$p), 180, 360)
    $path2.AddLine((24*$p), (60*$p), (18*$p), (72*$p))
    $path2.AddLine((18*$p), (72*$p), (30*$p), (60*$p))
    $path2.CloseFigure()
    $g.FillPath($brush2, $path2)
    $path2.Dispose()
    $brush2.Dispose()
    # "2" text in front bubble
    $fontSize = [Math]::Max(7, 22 * $p)
    $font = New-Object System.Drawing.Font("Arial", $fontSize, [System.Drawing.FontStyle]::Bold)
    $whiteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $textRect = New-Object System.Drawing.RectangleF((6*$p), (18*$p), (52*$p), (42*$p))
    $g.DrawString("2", $font, $whiteBrush, $textRect, $sf)
    $sf.Dispose()
    $whiteBrush.Dispose()
    $font.Dispose()
}

# --- Clean All: blue circle with green checkmark ---
Write-Host "Clean All..."
New-Icon "icon-clean-all" {
    param($g, $s)
    $p = $s / 80.0
    # Blue circle
    $pen = New-Object System.Drawing.Pen($blue, [Math]::Max(2, 4 * $p))
    $g.DrawEllipse($pen, (8*$p), (8*$p), (64*$p), (64*$p))
    $pen.Dispose()
    # Green checkmark
    $checkPen = New-Object System.Drawing.Pen($green, [Math]::Max(2, 5 * $p))
    $checkPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $checkPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $checkPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    $g.DrawLine($checkPen, (22*$p), (40*$p), (34*$p), (54*$p))
    $g.DrawLine($checkPen, (34*$p), (54*$p), (58*$p), (26*$p))
    $checkPen.Dispose()
}

# --- Partial Reply: blue curved arrow left ---
Write-Host "Partial Reply..."
New-Icon "icon-reply" {
    param($g, $s)
    $p = $s / 80.0
    $pen = New-Object System.Drawing.Pen($blue, [Math]::Max(2, 4 * $p))
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    # Curved arrow body
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc((14*$p), (20*$p), (50*$p), (40*$p), 180, -150)
    $g.DrawPath($pen, $path)
    $path.Dispose()
    $pen.Dispose()
    # Arrowhead
    $brush = New-Object System.Drawing.SolidBrush($blue)
    $arrowPts = @(
        (New-Object System.Drawing.PointF((10*$p), (40*$p))),
        (New-Object System.Drawing.PointF((26*$p), (28*$p))),
        (New-Object System.Drawing.PointF((26*$p), (52*$p)))
    )
    $g.FillPolygon($brush, $arrowPts)
    $brush.Dispose()
}

# --- Partial Reply All: double curved arrows in blue ---
Write-Host "Partial Reply All..."
New-Icon "icon-reply-all" {
    param($g, $s)
    $p = $s / 80.0
    # Back arrow (lighter)
    $pen1 = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(140, 21, 100, 192), [Math]::Max(2, 3.5 * $p))
    $pen1.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen1.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $path1 = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path1.AddArc((24*$p), (20*$p), (50*$p), (40*$p), 180, -150)
    $g.DrawPath($pen1, $path1)
    $path1.Dispose()
    $pen1.Dispose()
    $brush1 = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(140, 21, 100, 192))
    $arrowPts1 = @(
        (New-Object System.Drawing.PointF((20*$p), (40*$p))),
        (New-Object System.Drawing.PointF((36*$p), (28*$p))),
        (New-Object System.Drawing.PointF((36*$p), (52*$p)))
    )
    $g.FillPolygon($brush1, $arrowPts1)
    $brush1.Dispose()
    # Front arrow (full blue)
    $pen2 = New-Object System.Drawing.Pen($blue, [Math]::Max(2, 3.5 * $p))
    $pen2.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen2.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $path2 = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path2.AddArc((8*$p), (20*$p), (50*$p), (40*$p), 180, -150)
    $g.DrawPath($pen2, $path2)
    $path2.Dispose()
    $pen2.Dispose()
    $brush2 = New-Object System.Drawing.SolidBrush($blue)
    $arrowPts2 = @(
        (New-Object System.Drawing.PointF((4*$p), (40*$p))),
        (New-Object System.Drawing.PointF((20*$p), (28*$p))),
        (New-Object System.Drawing.PointF((20*$p), (52*$p)))
    )
    $g.FillPolygon($brush2, $arrowPts2)
    $brush2.Dispose()
}

# --- Partial Forward: green arrow right ---
Write-Host "Partial Forward..."
New-Icon "icon-forward" {
    param($g, $s)
    $p = $s / 80.0
    $pen = New-Object System.Drawing.Pen($green, [Math]::Max(2, 4 * $p))
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    # Curved arrow body (mirrored)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc((16*$p), (20*$p), (50*$p), (40*$p), 0, 150)
    $g.DrawPath($pen, $path)
    $path.Dispose()
    $pen.Dispose()
    # Arrowhead pointing right
    $brush = New-Object System.Drawing.SolidBrush($green)
    $arrowPts = @(
        (New-Object System.Drawing.PointF((70*$p), (40*$p))),
        (New-Object System.Drawing.PointF((54*$p), (28*$p))),
        (New-Object System.Drawing.PointF((54*$p), (52*$p)))
    )
    $g.FillPolygon($brush, $arrowPts)
    $brush.Dispose()
}

Write-Host "All icons generated!"
