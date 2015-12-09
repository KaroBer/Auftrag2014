Module blChart
    Dim nutzbareHoehe, nutzbareBreite As Double
    Dim platzFuerBalkenH As Double = 0

    Sub BalkenGrafik(ByRef canvas As Canvas, ByVal daten As Dictionary(Of String, Double),
                     ByVal Headline As String,
                     ByVal AchsenTextX As String, ByVal AchsenTextY As String,
                     Optional ByVal MinFontSizeHeadline As Integer = 20,
                     Optional ByVal raster As Integer = 10)

        Dim breite As Integer = canvas.Width : Dim hoehe As Integer = canvas.Height
        Dim dLinks, dRechts, dUnten, dOben As Double
        dLinks = breite * 0.05
        dRechts = dLinks
        dUnten = hoehe * 0.075
        dOben = hoehe * 0.035

        nutzbareHoehe = hoehe - dOben - dUnten
        nutzbareBreite = breite - dLinks - dRechts

        Dim achseOben, achseUnten, achseRechts As Double
        achseOben = dOben
        achseUnten = dOben + nutzbareHoehe
        achseRechts = dLinks + nutzbareBreite

        Dim headFontSize As Integer = raster * 2
        If headFontSize < MinFontSizeHeadline Then headFontSize = MinFontSizeHeadline

        AchsenZeichnen(canvas, dLinks, raster, achseUnten, achseOben, achseRechts)
        LinienRasterZeichnen(canvas, raster * 3, dLinks, achseRechts, achseUnten, achseOben, Brushes.Silver)

        BeschriftungEinfuegen(canvas, Headline, headFontSize,
                              AchsenTextX, AchsenTextY, 12,
                              dLinks, dOben, achseRechts, achseUnten, raster)

        BalkenEinzeichnen(canvas, daten, dLinks, achseUnten, raster, Brushes.White)

    End Sub

    Private Sub BalkenEinzeichnen(ByRef canvas As Canvas, daten As Dictionary(Of String, Double),
                                  ByVal dlinks As Integer,
                                  ByVal achseUnten As Integer, ByVal raster As Integer,
                                  ByVal balkenTextColor As Brush)

        ' Die Balken berechnen und einzeichen
        Dim anzahlBalken As Integer = daten.Count
        Dim balkenHoehe As Integer = CInt(platzFuerBalkenH / anzahlBalken - raster)
        If balkenHoehe <= raster * 2 Then
            raster = CInt(raster * 0.6)
            balkenHoehe = CInt(platzFuerBalkenH / anzahlBalken - raster)
        End If

        Dim einheit As Double = (nutzbareBreite - raster) / maxValue(daten)
        Dim startOben As Integer = achseUnten - (anzahlBalken * (balkenHoehe + raster))


        Dim i As Integer = 0
        For Each b In daten
            i += 1
            Dim balkenLaenge = b.Value * einheit
            addRectWithShadow(canvas, balkenLaenge, balkenHoehe,
                              nextColor(i), 1, dlinks + 1,
                              startOben, 2, 3, 3, shadowColor)

            Dim lbl As New Label
            lbl.Padding = New Thickness(0)
            lbl.Content = b.Key.ToString + " | " + ZweiStellenHintermKomma(b.Value.ToString)
            If balkenHoehe >= 10 + CInt(raster * 0.7) Then
                lbl.FontSize = balkenHoehe - CInt(raster * 0.7)
            Else
                lbl.FontSize = 10
            End If
            lbl.FontFamily = New FontFamily("Consolas")
            lbl.FontWeight = FontWeights.Bold
            lbl.Foreground = balkenTextColor
            lbl.Background = Brushes.Transparent
            Dim lblBreite As Double = MeasureTextSize(lbl.Content, lbl.FontSize, lbl.FontFamily.ToString).Width
            canvas.SetTop(lbl, startOben + CInt(raster * 0.2))
            canvas.SetLeft(lbl, dlinks + raster)
            canvas.SetZIndex(lbl, 4)
            canvas.Children.Add(lbl)

            startOben = startOben + balkenHoehe + raster
        Next
    End Sub

    Private Sub BeschriftungEinfuegen(ByRef canvas As Canvas,
                                      ByVal headline As String, ByVal fontSze As Integer,
                                      ByVal achsenTextX As String, ByVal achsenTextY As String,
                                      ByVal achsenTextFontSize As Integer,
                                      ByVal dlinks As Double, ByVal dOben As Double,
                                      ByVal achseRechts As Double, ByVal achseUnten As Double,
                                      ByVal raster As Integer)

        ' Hellgrauer Schatten der Headline
        Dim headw As New Label
        headw.Foreground = Brushes.LightGray
        headw.FontSize = fontSze : headw.FontWeight = FontWeights.Bold
        headw.Content = headline
        canvas.SetLeft(headw, CInt(dlinks) + raster - 1)
        canvas.SetTop(headw, CInt(dOben - raster + 1))
        canvas.SetZIndex(headw, 1)
        canvas.Children.Add(headw)

        ' Die Headline selbst in Schwarz
        Dim head As New Label
        head.FontSize = fontSze : head.FontWeight = FontWeights.Bold
        head.Content = headline
        canvas.SetLeft(head, CInt(dlinks) + raster)
        canvas.SetTop(head, CInt(dOben - raster))
        canvas.SetZIndex(head, 2)
        canvas.Children.Add(head)

        Dim headH As Double = MeasureTextSize(head.Content, head.FontSize, head.FontFamily.ToString).Height
        platzFuerBalkenH = nutzbareHoehe - headH - raster

        Dim textX As New Label
        textX.FontSize = achsenTextFontSize : textX.FontWeight = FontWeights.Bold
        textX.Content = achsenTextX
        Dim textXwidth As Double = MeasureTextSize(achsenTextX, achsenTextFontSize, textX.FontFamily.ToString).Width
        Dim textXheight As Double = MeasureTextSize(achsenTextX, achsenTextFontSize, textX.FontFamily.ToString).Height
        canvas.SetLeft(textX, achseRechts - CInt(textXwidth + 20))
        canvas.SetTop(textX, CInt(achseUnten - raster \ 2))
        canvas.Children.Add(textX)

        Dim lbltextY As New Label
        Dim rt As New RotateTransform(270)
        lbltextY.RenderTransform = rt

        lbltextY.FontSize = achsenTextFontSize : lbltextY.FontWeight = FontWeights.Bold
        lbltextY.Content = achsenTextY

        Dim textYwidth As Double = MeasureTextSize(achsenTextY, achsenTextFontSize, lbltextY.FontFamily.ToString).Width

        canvas.SetLeft(lbltextY, dlinks - 25)
        canvas.SetTop(lbltextY, dOben + 20 + textYwidth)
        canvas.Children.Add(lbltextY)
    End Sub

    Private Sub LinienRasterZeichnen(ByRef canvas As Canvas, ByVal abstand As Integer, ByVal dLinks As Double,
                                     ByVal achseRechts As Double, ByVal achseUnten As Double, ByVal achseOben As Double,
                                     ByVal lineColor As Brush, Optional ByVal lineWidth As Double = 1)

        ' Linienraster horizontal
        Dim anz As Integer = CInt(nutzbareHoehe / abstand)
        For i = anz To 2 Step -1
            Dim l As New Line()
            l.Stroke = lineColor : l.StrokeThickness = lineWidth
            l.X1 = dLinks - 6 : l.Y1 = achseUnten - (abstand * (anz - i + 1))
            l.X2 = achseRechts : l.Y2 = achseUnten - (abstand * (anz - i + 1))
            canvas.Children.Add(l)
        Next

        ' Linienraster vertikal
        anz = CInt(nutzbareBreite / abstand)
        For i = 1 To anz
            Dim l1 As New Line()
            l1.Stroke = lineColor : l1.StrokeThickness = lineWidth
            l1.X1 = dLinks + i * abstand : l1.Y1 = achseOben + abstand - 10
            l1.X2 = dLinks + i * abstand : l1.Y2 = achseUnten + 6
6:          canvas.Children.Add(l1)
        Next
    End Sub

    Private Sub AchsenZeichnen(ByRef canvas As Canvas, ByVal dLinks As Double, ByVal raster As Double,
                               ByVal achseUnten As Double, ByVal achseoben As Double,
                               ByVal achseRechts As Double, Optional ByVal dicke As Integer = 1)

        Dim yAchse As New Line()
        yAchse.Stroke = Brushes.Black : yAchse.StrokeThickness = dicke
        yAchse.X1 = dLinks : yAchse.Y1 = achseoben
        yAchse.X2 = dLinks : yAchse.Y2 = achseUnten + raster
        canvas.Children.Add(yAchse)

        ' Pfeilspitze
        Dim L1 As New Line()
        L1.Stroke = Brushes.Black : L1.StrokeThickness = dicke
        L1.X1 = dLinks - raster / 2 : L1.Y1 = achseoben + raster
        L1.X2 = dLinks : L1.Y2 = achseoben
        canvas.Children.Add(L1)

        Dim L2 As New Line()
        L2.Stroke = Brushes.Black : L2.StrokeThickness = dicke
        L2.X1 = dLinks : L2.Y1 = achseoben
        L2.X2 = dLinks + raster / 2 : L2.Y2 = achseoben + raster
        canvas.Children.Add(L2)

        Dim xAchse As New Line()
        xAchse.Stroke = Brushes.Black : xAchse.StrokeThickness = dicke
        xAchse.X1 = dLinks - raster : xAchse.Y1 = achseUnten
        xAchse.X2 = achseRechts : xAchse.Y2 = achseUnten
        canvas.Children.Add(xAchse)

        ' Pfeilspitze
        Dim L3 As New Line()
        L3.Stroke = Brushes.Black : L3.StrokeThickness = dicke
        L3.X1 = achseRechts - raster : L3.Y1 = achseUnten - raster / 2
        L3.X2 = achseRechts : L3.Y2 = achseUnten
        canvas.Children.Add(L3)

        Dim L4 As New Line()
        L4.Stroke = Brushes.Black : L4.StrokeThickness = dicke
        L4.X1 = achseRechts : L4.Y1 = achseUnten
        L4.X2 = achseRechts - raster : L4.Y2 = achseUnten + raster / 2
        canvas.Children.Add(L4)

    End Sub

    Private Sub addRectWithShadow(ByRef c As Canvas, ByVal width As Double, ByVal height As Double,
                                  ByVal fillColor As Brush, ByVal frameThick As Integer,
                                  ByVal posLeft As Integer, ByVal posTop As Integer, ByVal zIdx As Integer,
                                  ByVal shadowDX As Integer, ByVal shadowDY As Integer, ByVal shadowColor As Brush)

        c.Children.Add(NewRect(width, height, fillColor, frameThick, posLeft, posTop, zIdx))
        c.Children.Add(NewRect(width, height, Brushes.DarkGray, 0, posLeft + shadowDX, posTop - shadowDY, 0, 6))
    End Sub

    Private Function NewRect(ByVal width As Double, ByVal height As Double, ByVal fillColor As Brush,
                             ByVal frameThick As Integer,
                             ByVal posLeft As Integer, ByVal posTop As Integer, ByVal zIdx As Integer,
                             Optional ByVal blurIT As Integer = 0)

        Dim r As New Rectangle()
        r.Width = width : r.Height = height : r.Fill = fillColor
        r.Stroke = New SolidColorBrush(Colors.Black) : r.StrokeThickness = frameThick
        Canvas.SetLeft(r, posLeft)
        Canvas.SetTop(r, posTop)
        Canvas.SetZIndex(r, zIdx)

        If blurIT > 0 Then
            Dim blur As New Effects.BlurEffect()
            blur.Radius = blurIT
            r.Effect = blur
        End If

        Return r
    End Function

    Private Function shadowColor()
        Dim myBrush As New LinearGradientBrush()
        myBrush.GradientStops.Add(New GradientStop(Colors.LightGray, 0.0))
        myBrush.GradientStops.Add(New GradientStop(Colors.Gray, 0.5))
        myBrush.GradientStops.Add(New GradientStop(Colors.DarkGray, 1.0))

        Return myBrush
    End Function

    Private Function nextColor(ByVal i As Integer) As Brush
        Dim b As New SolidColorBrush
        Dim max As Integer = 9
        If i > max Then i = i Mod max

        Select Case i
            Case Is = 1
                b = Brushes.Sienna
            Case Is = 2
                b = Brushes.SkyBlue
            Case Is = 3
                b = Brushes.MediumTurquoise
            Case Is = 4
                b = Brushes.SteelBlue
            Case Is = 5
                b = Brushes.Tomato
            Case Is = 6
                b = Brushes.Olive
            Case Is = 7
                b = Brushes.DarkSlateBlue
            Case Is = 8
                b = Brushes.Orange
            Case Is = 9
            Case Else
                b = Brushes.Moccasin

        End Select
        Return b
    End Function

    Public Function MeasureTextSize(ByVal text As String, ByVal fontSize As Double, ByVal typeFace As String) As Size

        typeFace = Replace(typeFace, "}", "")
        typeFace = Replace(typeFace, "{", "")

        Dim ft As FormattedText = New FormattedText(text, Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, New Typeface(typeFace), fontSize, Brushes.Black)
        Return New Size(ft.Width, ft.Height)

    End Function

    Private Function maxValue(ByVal daten As Dictionary(Of String, Double)) As Double
        Dim ret As Double
        For Each pair In daten
            If ret < pair.Value Then ret = pair.Value
        Next

        Return ret
    End Function
End Module
