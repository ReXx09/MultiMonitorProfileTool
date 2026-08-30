param(
    [string]$OutputPath = "$PSScriptRoot\MultiMonitorProfileTool.ico",
    [int]$Size = 256
)

Add-Type -AssemblyName System.Drawing

function Create-MonitorIcon {
    param(
        [int]$Size = 256
    )
    
    $bitmap = New-Object System.Drawing.Bitmap($Size, $Size)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    
    $graphics.Clear([System.Drawing.Color]::White)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
    
    # Farben
    $colorPrimary = [System.Drawing.Color]::FromArgb(255, 59, 130, 246)        # Bright Blue
    $colorSecondary = [System.Drawing.Color]::FromArgb(255, 34, 197, 94)       # Bright Green
    $colorAccent = [System.Drawing.Color]::FromArgb(255, 249, 115, 22)         # Orange
    $colorBg = [System.Drawing.Color]::FromArgb(255, 241, 245, 249)            # Very light gray
    $colorDark = [System.Drawing.Color]::FromArgb(255, 30, 41, 59)             # Dark slate
    $colorInner = [System.Drawing.Color]::FromArgb(255, 191, 219, 254)         # Light blue
    
    # Hintergrund
    $bgBrush = New-Object System.Drawing.SolidBrush($colorBg)
    $graphics.FillRectangle($bgBrush, 0, 0, $Size, $Size)
    
    $brushPrimary = New-Object System.Drawing.SolidBrush($colorPrimary)
    $brushSecondary = New-Object System.Drawing.SolidBrush($colorSecondary)
    $brushAccent = New-Object System.Drawing.SolidBrush($colorAccent)
    $brushInner = New-Object System.Drawing.SolidBrush($colorInner)
    $brushDark = New-Object System.Drawing.SolidBrush($colorDark)
    
    $penDark = New-Object System.Drawing.Pen($colorDark, 2)
    
    # Skalierung
    $centerX = $Size / 2
    $centerY = $Size / 2
    $scale = $Size / 256
    
    # ==== Hauptmonitor (zentral, groß, deutlich) ====
    $monW = 140 * $scale
    $monH = 95 * $scale
    $monX = $centerX - $monW / 2
    $monY = $centerY - $monH / 2 - 10 * $scale
    
    # Hauptmonitor-Rahmen
    $monRect = New-Object System.Drawing.Rectangle(
        [int]$monX, [int]$monY, [int]$monW, [int]$monH
    )
    $graphics.FillRectangle($brushPrimary, $monRect)
    $graphics.DrawRectangle($penDark, $monRect)
    
    # Hauptmonitor-Display (innerer Bereich)
    $dispX = $monX + 4 * $scale
    $dispY = $monY + 4 * $scale
    $dispW = $monW - 8 * $scale
    $dispH = $monH - 18 * $scale
    $dispRect = New-Object System.Drawing.Rectangle(
        [int]$dispX, [int]$dispY, [int]$dispW, [int]$dispH
    )
    $graphics.FillRectangle($brushInner, $dispRect)
    
    # Ständer - zwei Beine
    $legW = 6 * $scale
    $legX1 = $monX + $monW * 0.3
    $legX2 = $monX + $monW * 0.7
    $legY = $monY + $monH
    $legBottomY = $legY + 18 * $scale
    
    $leg1 = New-Object System.Drawing.Rectangle(
        [int]$legX1, [int]$legY, [int]$legW, [int](18 * $scale)
    )
    $leg2 = New-Object System.Drawing.Rectangle(
        [int]$legX2, [int]$legY, [int]$legW, [int](18 * $scale)
    )
    $graphics.FillRectangle($brushDark, $leg1)
    $graphics.FillRectangle($brushDark, $leg2)
    
    # Basis (Fuß)
    $baseW = 56 * $scale
    $baseH = 5 * $scale
    $baseX = $centerX - $baseW / 2
    $baseY = $legBottomY - $baseH / 2
    $baseRect = New-Object System.Drawing.Rectangle(
        [int]$baseX, [int]$baseY, [int]$baseW, [int]$baseH
    )
    $graphics.FillRectangle($brushDark, $baseRect)
    
    # ==== Kleine Monitore (oben links & rechts) ====
    $smallW = 40 * $scale
    $smallH = 27 * $scale
    
    # Oben rechts (grün)
    $small1X = $centerX + 65 * $scale
    $small1Y = $monY - 8 * $scale
    $small1Rect = New-Object System.Drawing.Rectangle(
        [int]$small1X, [int]$small1Y, [int]$smallW, [int]$smallH
    )
    $graphics.FillRectangle($brushSecondary, $small1Rect)
    $graphics.DrawRectangle($penDark, $small1Rect)
    
    # Oben links (orange)
    $small2X = $centerX - 105 * $scale
    $small2Y = $monY - 8 * $scale
    $small2Rect = New-Object System.Drawing.Rectangle(
        [int]$small2X, [int]$small2Y, [int]$smallW, [int]$smallH
    )
    $graphics.FillRectangle($brushAccent, $small2Rect)
    $graphics.DrawRectangle($penDark, $small2Rect)
    
    # ==== Verbindungslinien (Vernetzung darstellen) ====
    $linePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(180, 148, 163, 184), 2)
    $linePen.DashStyle = [System.Drawing.Drawing2D.DashStyle]::Dot
    
    # Linie von rechts zu Mitte
    $graphics.DrawLine($linePen, 
        [int]($small1X + $smallW / 2), 
        [int]($small1Y + $smallH), 
        [int]($monX + $monW / 2), 
        [int]$monY
    )
    
    # Linie von links zu Mitte
    $graphics.DrawLine($linePen,
        [int]($small2X + $smallW / 2),
        [int]($small2Y + $smallH),
        [int]($monX + $monW / 2),
        [int]$monY
    )
    
    # Ressourcen freigeben
    $graphics.Dispose()
    return $bitmap
}

try {
    Write-Host "Erstelle verbessertes Monitor-Icon (256 x 256 px)..."
    Write-Host "Design: Zentral großer Monitor + kleine Nebemonitore + Vernetzungslinien"
    $bitmap = Create-MonitorIcon -Size $Size
    
    $icon = [System.Drawing.Icon]::FromHandle($bitmap.GetHicon())
    $fileStream = [System.IO.FileStream]::new($OutputPath, [System.IO.FileMode]::Create)
    $icon.Save($fileStream)
    $fileStream.Close()
    $icon.Dispose()
    $bitmap.Dispose()
    
    Write-Host "Icon erfolgreich erstellt: $OutputPath" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "Fehler beim Erstellen des Icons: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.Exception.StackTrace -ForegroundColor Red
    exit 1
}
