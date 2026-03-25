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

# --- Clean All: broom with green sparkles ---
Write-Host "Clean All..."
New-Icon "icon-clean-all" {
    param($g, $s)
    $p = $s / 80.0
    # Broom handle (diagonal)
    $handlePen = New-Object System.Drawing.Pen($blue, [Math]::Max(2, 3.5 * $p))
    $handlePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $handlePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $g.DrawLine($handlePen, (56*$p), (8*$p), (22*$p), (50*$p))
    $handlePen.Dispose()
    # Broom bristles (fan shape at bottom)
    $bristlePen = New-Object System.Drawing.Pen($blue, [Math]::Max(1.5, 2.5 * $p))
    $bristlePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $bristlePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $g.DrawLine($bristlePen, (22*$p), (50*$p), (8*$p), (72*$p))
    $g.DrawLine($bristlePen, (22*$p), (50*$p), (14*$p), (68*$p))
    $g.DrawLine($bristlePen, (22*$p), (50*$p), (18*$p), (74*$p))
    $g.DrawLine($bristlePen, (22*$p), (50*$p), (28*$p), (74*$p))
    $g.DrawLine($bristlePen, (22*$p), (50*$p), (32*$p), (72*$p))
    $g.DrawLine($bristlePen, (22*$p), (50*$p), (38*$p), (72*$p))
    $bristlePen.Dispose()
    # Green sparkles (3 four-pointed stars, bigger)
    $sparkBrush = New-Object System.Drawing.SolidBrush($green)
    # Sparkle 1 (top-right, large)
    $cx = 66*$p; $cy = 16*$p; $r = 8*$p
    $g.FillPolygon($sparkBrush, @(
        (New-Object System.Drawing.PointF($cx, ($cy - $r))),
        (New-Object System.Drawing.PointF(($cx + $r*0.3), ($cy - $r*0.3))),
        (New-Object System.Drawing.PointF(($cx + $r), $cy)),
        (New-Object System.Drawing.PointF(($cx + $r*0.3), ($cy + $r*0.3))),
        (New-Object System.Drawing.PointF($cx, ($cy + $r))),
        (New-Object System.Drawing.PointF(($cx - $r*0.3), ($cy + $r*0.3))),
        (New-Object System.Drawing.PointF(($cx - $r), $cy)),
        (New-Object System.Drawing.PointF(($cx - $r*0.3), ($cy - $r*0.3)))
    ))
    # Sparkle 2 (right, medium)
    $cx = 70*$p; $cy = 40*$p; $r = 6*$p
    $g.FillPolygon($sparkBrush, @(
        (New-Object System.Drawing.PointF($cx, ($cy - $r))),
        (New-Object System.Drawing.PointF(($cx + $r*0.3), ($cy - $r*0.3))),
        (New-Object System.Drawing.PointF(($cx + $r), $cy)),
        (New-Object System.Drawing.PointF(($cx + $r*0.3), ($cy + $r*0.3))),
        (New-Object System.Drawing.PointF($cx, ($cy + $r))),
        (New-Object System.Drawing.PointF(($cx - $r*0.3), ($cy + $r*0.3))),
        (New-Object System.Drawing.PointF(($cx - $r), $cy)),
        (New-Object System.Drawing.PointF(($cx - $r*0.3), ($cy - $r*0.3)))
    ))
    # Sparkle 3 (smaller)
    $cx = 54*$p; $cy = 32*$p; $r = 4.5*$p
    $g.FillPolygon($sparkBrush, @(
        (New-Object System.Drawing.PointF($cx, ($cy - $r))),
        (New-Object System.Drawing.PointF(($cx + $r*0.3), ($cy - $r*0.3))),
        (New-Object System.Drawing.PointF(($cx + $r), $cy)),
        (New-Object System.Drawing.PointF(($cx + $r*0.3), ($cy + $r*0.3))),
        (New-Object System.Drawing.PointF($cx, ($cy + $r))),
        (New-Object System.Drawing.PointF(($cx - $r*0.3), ($cy + $r*0.3))),
        (New-Object System.Drawing.PointF(($cx - $r), $cy)),
        (New-Object System.Drawing.PointF(($cx - $r*0.3), ($cy - $r*0.3)))
    ))
    $sparkBrush.Dispose()
}

# --- Keep Selection Only: crop frame around text ---
Write-Host "Keep Selection..."
New-Icon "icon-keep-selection" {
    param($g, $s)
    $p = $s / 80.0
    # Faded text lines outside crop area
    $fadePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(50, 21, 100, 192), [Math]::Max(1, 2 * $p))
    $fadePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $fadePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $g.DrawLine($fadePen, (16*$p), (10*$p), (64*$p), (10*$p))
    $g.DrawLine($fadePen, (16*$p), (18*$p), (54*$p), (18*$p))
    $g.DrawLine($fadePen, (16*$p), (62*$p), (64*$p), (62*$p))
    $g.DrawLine($fadePen, (16*$p), (70*$p), (48*$p), (70*$p))
    $fadePen.Dispose()
    # Bold text lines inside crop
    $boldPen = New-Object System.Drawing.Pen($blue, [Math]::Max(2, 3 * $p))
    $boldPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $boldPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $g.DrawLine($boldPen, (16*$p), (32*$p), (64*$p), (32*$p))
    $g.DrawLine($boldPen, (16*$p), (40*$p), (58*$p), (40*$p))
    $g.DrawLine($boldPen, (16*$p), (48*$p), (64*$p), (48*$p))
    $boldPen.Dispose()
    # Green crop corners (thick L-shapes)
    $cropPen = New-Object System.Drawing.Pen($green, [Math]::Max(2, 3.5 * $p))
    $cropPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $cropPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $cornerLen = 12 * $p
    # Top-left
    $g.DrawLine($cropPen, (8*$p), (24*$p), (8*$p), (24*$p + $cornerLen))
    $g.DrawLine($cropPen, (8*$p), (24*$p), (8*$p + $cornerLen), (24*$p))
    # Top-right
    $g.DrawLine($cropPen, (72*$p), (24*$p), (72*$p), (24*$p + $cornerLen))
    $g.DrawLine($cropPen, (72*$p), (24*$p), (72*$p - $cornerLen), (24*$p))
    # Bottom-left
    $g.DrawLine($cropPen, (8*$p), (56*$p), (8*$p), (56*$p - $cornerLen))
    $g.DrawLine($cropPen, (8*$p), (56*$p), (8*$p + $cornerLen), (56*$p))
    # Bottom-right
    $g.DrawLine($cropPen, (72*$p), (56*$p), (72*$p), (56*$p - $cornerLen))
    $g.DrawLine($cropPen, (72*$p), (56*$p), (72*$p - $cornerLen), (56*$p))
    $cropPen.Dispose()
}

Write-Host "All icons generated!"
