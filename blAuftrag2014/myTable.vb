Public Class myColumn
    Public Enum alignTyp
        left = 1
        right = 2
        center = 3
        dezimal = 4
        dezWithPrefix = 5
        dezWithSuffix = 6
        euro = 7
    End Enum

    Property width As Integer
    Property align As alignTyp
    Property fill As String
    Property nachKomma As Integer
    Property prefix As String
    Property suffix As String

    Sub New()
        Me.width = 0 : Me.align = myColumn.alignTyp.left : Me.fill = " "
        Me.nachKomma = 2 : Me.prefix = "" : Me.suffix = ""
    End Sub

    Sub New(ByVal muster As myColumn)
        Me.nachKomma = muster.nachKomma
        Me.width = muster.width
        Me.align = muster.align
        Me.fill = muster.fill
        Me.prefix = muster.prefix
        Me.suffix = muster.suffix
    End Sub

    Sub New(ByVal width As Integer,
            Optional ByVal align As alignTyp = myColumn.alignTyp.left,
            Optional ByVal fill As String = " ",
            Optional ByVal NachKommaStellen As Integer = 2,
            Optional ByVal prefix As String = "",
            Optional ByVal suffix As String = "")

        Me.suffix = suffix
        Me.prefix = prefix
        Me.nachKomma = NachKommaStellen
        Me.width = width
        Me.align = align
        Me.fill = Mid(fill, 1, 1)
    End Sub

    Sub EigenschaftenAnlegen(ByVal s As String)
        s = Replace(s, " ", "")
        If s <> "" Then
            ' Eigenschaften aus String auslesen und anlegen
            Dim tmp As String = ""

            Dim x As Integer = InStr(s.ToLower, "euro")
            If x > 0 Then
                Me.align = myColumn.alignTyp.euro
                tmp = Mid(s, x, 4) : s = Replace(s, tmp, "")
            End If

            x = InStr(s.ToLower, "center")
            If x > 0 Then
                Me.align = myColumn.alignTyp.center
                tmp = Mid(s, x, 6) : s = Replace(s, tmp, "")
            End If

            x = InStr(s.ToLower, "right")
            If x > 0 Then
                Me.align = myColumn.alignTyp.right
                tmp = Mid(s, x, 5) : s = Replace(s, tmp, "")
            End If

            x = InStr(s.ToLower, "dezimal") + InStr(s.ToLower, "decimal")
            If x > 0 Then
                Me.align = myColumn.alignTyp.dezimal
                tmp = Mid(s, x, 7) : s = Replace(s, tmp, "")
            End If

            x = InStr(s.ToLower, "dezwithsuffix") + InStr(s.ToLower, "decwithsuffix")
            If x > 0 Then
                Me.align = myColumn.alignTyp.dezWithSuffix
                tmp = Mid(s, x, 13) : s = Replace(s, tmp, "")
            End If

            x = InStr(s.ToLower, "dezwithprefix") + InStr(s.ToLower, "dezwithpräfix") +
                InStr(s.ToLower, "decwithprefix") + InStr(s.ToLower, "decwithpräfix")
            If x > 0 Then
                Me.align = myColumn.alignTyp.dezWithPrefix
                tmp = Mid(s, x, 13) : s = Replace(s, tmp, "")
            End If

            x = InStr(s.ToLower, "fill")
            If x > 0 And s.Length >= x + 4 Then
                Me.fill = Mid(s, x + 4, 1)
                tmp = Mid(s, x, 5)
                s = Replace(s, tmp, "")
            End If

            x = InStr(s.ToLower, "nk")
            If x > 0 And s.Length >= x + 2 Then
                Try
                    Me.nachKomma = CInt(Mid(s, x + 2, 1))
                Catch ex As Exception
                    'keine Änderung, es bleibt bei Me.nachKomma = 2
                End Try
                tmp = Mid(s, x, 3)
                s = Replace(s, tmp, "")
            End If

            x = InStr(s.ToLower, "suffix")
            If x > 0 And s.Length >= x + 6 Then
                Me.suffix = Mid(s, x + 6)
                tmp = Mid(s, x, Me.suffix.Length + 6) : s = Replace(s, tmp, "")
            End If

            x = InStr(s.ToLower, "prefix") + InStr(s.ToLower, "präfix")
            If x > 0 And s.Length >= x + 6 Then
                Me.prefix = Mid(s, x + 6)
                tmp = Mid(s, x, Me.prefix.Length + 6) : s = Replace(s, tmp, "")
            End If


            Try
                Me.width = CInt(s)
            Catch ex As Exception
                ' wenn's nicht klappt, dann bleibts bei Me.width=0
            End Try
        End If
    End Sub

    Public Function ColumnInfo() As String()
        Dim cInfo(5) As String
        cInfo(0) = Me.width.ToString
        cInfo(1) = Me.align.ToString
        cInfo(2) = Me.fill.snText
        cInfo(3) = Me.nachKomma.ToString
        cInfo(4) = Me.prefix.snText
        cInfo(5) = Me.suffix.snText

        Return cInfo
    End Function

    Public Overrides Function toString() As String
        Return "Typ: myColumn" & vbCrLf & "width=" & Me.width & " align=" & Me.align.ToString & vbCrLf & "fill=" & Me.fill & " suffix=" & Me.suffix & vbCrLf & "NachkommaStellen=" & Me.nachKomma.ToString
    End Function
End Class

Public Class myTable
    Property A As myColumn
    Property B As myColumn
    Property C As myColumn
    Property D As myColumn
    Property E As myColumn
    Property F As myColumn

    Property AB As myColumn
    Property ABC As myColumn
    Property ABCD As myColumn
    Property ABCDE As myColumn
    Property ABCDEF As myColumn

    Property BC As myColumn
    Property BCD As myColumn
    Property BCDE As myColumn
    Property BCDEF As myColumn

    Property CD As myColumn
    Property CDE As myColumn
    Property CDEF As myColumn

    Property DE As myColumn
    Property DEF As myColumn

    Property EF As myColumn

    Property sb As System.Text.StringBuilder

    Private Sub uebrigeSpaltenAnlegen(ByVal spA As myColumn, ByVal spB As myColumn,
                                      Optional ByVal spC As myColumn = Nothing,
                                      Optional ByVal spD As myColumn = Nothing,
                                      Optional ByVal spE As myColumn = Nothing,
                                      Optional ByVal spF As myColumn = Nothing)

        Me.AB = New myColumn(spB)
        Me.AB.width = spA.width + spB.width

        If spC IsNot Nothing Then
            Me.ABC = New myColumn(spC)
            Me.ABC.width = spA.width + spB.width + spC.width

            Me.BC = New myColumn(spC)
            Me.BC.width = spB.width + spC.width
        End If

        If spC IsNot Nothing And spD IsNot Nothing Then
            Me.ABCD = New myColumn(spD)
            Me.ABCD.width = spA.width + spB.width + spC.width + spD.width

            Me.BCD = New myColumn(spD)
            Me.BCD.width = spB.width + spC.width + spD.width

            Me.CD = New myColumn(spD)
            Me.CD.width = spC.width + spD.width
        End If

        If spC IsNot Nothing And spD IsNot Nothing And spE IsNot Nothing Then
            Me.ABCDE = New myColumn(spE)
            Me.ABCDE.width = spA.width + spB.width + spC.width + spD.width + spE.width

            Me.BCDE = New myColumn(spE)
            Me.BCDE.width = spB.width + spC.width + spD.width + spE.width

            Me.CDE = New myColumn(spE)
            Me.CDE.width = spC.width + spD.width + spE.width

            Me.DE = New myColumn(spE)
            Me.DE.width = spD.width + spE.width
        End If

        If spC IsNot Nothing And spD IsNot Nothing And spE IsNot Nothing And spF IsNot Nothing Then
            Me.ABCDEF = New myColumn(spF)
            Me.ABCDEF.width = spA.width + spB.width + spC.width + spD.width + spE.width + spF.width

            Me.BCDEF = New myColumn(spF)
            Me.BCDEF.width = spB.width + spC.width + spD.width + spE.width + spF.width

            Me.CDEF = New myColumn(spF)
            Me.CDEF.width = spC.width + spD.width + spE.width + spF.width

            Me.DEF = New myColumn(spF)
            Me.DEF.width = spD.width + spE.width + spF.width

            Me.EF = New myColumn(spF)
            Me.EF.width = spE.width + spF.width

        End If
    End Sub

    Sub New(ByVal sA As String, ByVal sB As String, Optional ByVal sC As String = "",
            Optional ByVal sD As String = "", Optional ByVal sE As String = "", Optional ByVal sF As String = "")
        ' Die Spaltenbreite sollte immer der erste Wert sein.
        ' Leerzeichen dienen nur der besseren Lesbarkeit, sie werden in EigenschaftenAnlegen gelöscht
        ' Ausgewertet werden
        '               fill: Nur ein Zeichen ist erlaubt, etwa Fill. für einen Punkt als Füllzeichen.
        '              align: center, right, euro, dezimal, dezWithPrefix, dezWithSuffix
        '                     prefix, suffix
        '            WICHTIG: Präfix und Suffix MÜSSEN am ENDE des Spaltenstrings stehen!
        '                     Es darf nur entweder Präfix oder Suffix vorkommen.  
        '   NachkommaStellen: NK  --> Nur eine Ziffer ist erlaubt, etwa NK5

        Me.sb = New System.Text.StringBuilder
        Dim c34 As String = Chr(34)

        Me.A = New myColumn()
        Me.A.EigenschaftenAnlegen(sA)

        Me.B = New myColumn()
        Me.B.EigenschaftenAnlegen(sB)

        If sC <> "" Then
            Me.C = New myColumn()
            Me.C.EigenschaftenAnlegen(sC)
        End If

        If sD <> "" Then
            Me.D = New myColumn()
            Me.D.EigenschaftenAnlegen(sD)
        End If

        If sE <> "" Then
            Me.E = New myColumn()
            Me.E.EigenschaftenAnlegen(sE)
        End If

        If sF <> "" Then
            Me.F = New myColumn()
            Me.F.EigenschaftenAnlegen(sF)
        End If
        uebrigeSpaltenAnlegen(Me.A, Me.B, Me.C, Me.D, Me.E, Me.F)
    End Sub

    Sub New(ByVal spA As myColumn, ByVal spB As myColumn, Optional ByVal spC As myColumn = Nothing,
            Optional ByVal spD As myColumn = Nothing, Optional ByVal spE As myColumn = Nothing,
            Optional ByVal spF As myColumn = Nothing)

        Me.sb = New System.Text.StringBuilder
        Me.A = spA : Me.B = spB
        Me.C = spC : Me.D = spD : Me.E = spE
        uebrigeSpaltenAnlegen(spA, spB, spC, spD, spE, spF)
    End Sub

    Public Shared Function format(s As String, ByVal A As myColumn) As String

        ' kürzen wenn der Text nicht in die Spalte passt
        If s.Length > A.width Then s = Mid(s, 1, A.width)

        Dim ret As String = ""
        If A.align = myColumn.alignTyp.left Then
            ret = Mid(s + A.fill.Mal(A.width - s.Length), 1, A.width)

        ElseIf A.align = myColumn.alignTyp.right Then
            ret = Mid(A.fill.Mal(A.width - s.Length) + s, 1, A.width)

        ElseIf A.align = myColumn.alignTyp.center Then
            If A.width = s.Length Then
                ret = s
            Else
                Dim frei As Integer = A.width - s.Length
                ret = Mid(A.fill.Mal(frei \ 2) + s + A.fill.Mal(frei), 1, A.width)
            End If

        ElseIf A.align = myColumn.alignTyp.dezimal Or
               A.align = myColumn.alignTyp.dezWithPrefix Or
               A.align = myColumn.alignTyp.dezWithSuffix Or
               A.align = myColumn.alignTyp.euro Then

            'Eigentlich ein Fehler, weil ein String übergeben wird!
            ret = tryFormatAsDouble(A, s)
        End If

        Return ret
    End Function

    Public Shared Function format(ByVal d As Double, ByVal A As myColumn) As String
        d = Math.Round(d, A.nachKomma)
        Dim dString As String = d.ToString("0." + "0".Mal(A.nachKomma))

        Dim ret As String = ""
        If A.align = myColumn.alignTyp.left Or _
           A.align = myColumn.alignTyp.right Or _
           A.align = myColumn.alignTyp.center Then

            ' In String gewandelt ausrichten
            ret = format(dString, A)

        ElseIf A.align = myColumn.alignTyp.dezimal Then
            ret = formatAnpassen(A, dString)

        ElseIf A.align = myColumn.alignTyp.dezWithPrefix Then
            If Right(A.prefix, 1) <> " " Then A.prefix += " " ' Zwangsabstand zur Zahl
            dString = A.prefix + dString
            ret = formatAnpassen(A, dString)

        ElseIf A.align = myColumn.alignTyp.dezWithSuffix Then
            If Left(A.suffix, 1) <> " " Then A.suffix = " " + A.suffix ' Zwangsabstand zur Zahl
            dString += A.suffix
            ret = formatAnpassen(A, dString)

        ElseIf A.align = myColumn.alignTyp.euro Then
            dString += " Euro"
            ret = formatAnpassen(A, dString)
        End If

        Return ret
    End Function

    Public Shared Function formatAnpassen(ByVal A As myColumn, ByVal dString As String) As String
        If A.width < dString.Length Then A.width = dString.Length 'verschiebt die Tabelle, aber die Zahl bleibt korrekt
        Dim PlatzDavor As Integer = A.width - dString.Length

        Return A.fill.Mal(PlatzDavor) + dString
    End Function

    Public Shared Function tryFormatAsDouble(ByVal A As myColumn, ByVal s As String) As String
        Dim ret As String = ""

        Dim test As Double = 0
        Try
            'Testen ob ein Double-Wert als String übergeben wurde
            test = CDbl(s)
            'OK. dann in die richtige Funktion einspeisen
            ret = format(test, A)

        Catch ex As Exception
            ' Wurde kein Double-Wert übergeben, behandeln wie Rechtsbündig
            ret = Mid(A.fill.Mal(A.width - s.Length) + s, 1, A.width)
        End Try

        Return ret
    End Function

    Private Function infoColumnWidth(ByVal c As myColumn, ByVal minWidth As Integer, ByVal maxWidth As Integer) As Integer
        Dim ret As Integer = minWidth

        Dim cInfo() As String = c.ColumnInfo
        For i = 0 To 5
            If ret < cInfo(i).Length Then ret = cInfo(i).Length
        Next
        If ret < minWidth Then ret = minWidth
        If ret > maxWidth Then ret = maxWidth

        Return ret
    End Function

    Function Info(Optional ByVal c As String = "") As String
        Dim c34 As String = Chr(34)
        Dim sb As New System.Text.StringBuilder

        Dim minWidth As Integer = 6
        Dim maxWidth As Integer = 12
        Dim grA As Integer = infoColumnWidth(Me.A, minWidth, maxWidth) + 10  ' für "| width  : "
        Dim grB As Integer = infoColumnWidth(Me.B, minWidth, maxWidth) + 3   ' für "| "
        Dim grC As Integer = minWidth
        If Me.C IsNot Nothing Then grC = infoColumnWidth(Me.C, minWidth, maxWidth) + 3
        Dim grD As Integer = minWidth
        If Me.D IsNot Nothing Then grD = infoColumnWidth(Me.D, minWidth, maxWidth) + 3
        Dim grE As Integer = minWidth
        If Me.E IsNot Nothing Then grE = infoColumnWidth(Me.E, minWidth, maxWidth) + 3
        Dim grF As Integer = minWidth
        If Me.F IsNot Nothing Then grF = infoColumnWidth(Me.F, minWidth, maxWidth) + 3

        Dim t As New myTable(grA.ToString, grB.ToString, grC.tostring, grD.tostring, grE.tostring, grF.tostring)

        Dim na As String = "n/a"
        Dim cA() As String = Me.A.ColumnInfo
        Dim cB() As String = Me.B.ColumnInfo
        Dim cC() As String = {na, na, na, na, na, na}
        Dim cD() As String = {na, na, na, na, na, na}
        Dim cE() As String = {na, na, na, na, na, na}
        Dim cF() As String = {na, na, na, na, na, na}

        If Me.C IsNot Nothing Then cC = Me.C.ColumnInfo
        If Me.D IsNot Nothing Then cD = Me.D.ColumnInfo
        If Me.E IsNot Nothing Then cE = Me.E.ColumnInfo
        If Me.F IsNot Nothing Then cF = Me.F.ColumnInfo

        Dim w(5) As String
        w(0) = "| Width : " + cA(0)
        w(1) = "| Align : " + cA(1)
        w(2) = "| Fill  : " + cA(2)
        w(3) = "| Nachkomma: " + cA(3)
        w(4) = "| Prefix: " + cA(4)
        w(5) = "| Suffix: " + cA(5)

        Dim line As String = "|" + "-".Mal(t.A.width - 1) + "|" + "-".Mal(t.B.width - 1) + "|" + "-".Mal(t.C.width - 1) +
                             "|" + "-".Mal(t.D.width - 1) + "|" + "-".Mal(t.E.width - 1) + "|" + "-".Mal(t.F.width - 1)
        sb.AppendLine(c + " myTable.Info(" + c34 + c + c34 + ")")
        sb.AppendLine(c + line + "|")
        sb.AppendLine(c + "| A".f(t.A) + "| B".f(t.B) + "| C".f(t.C) + "| D".f(t.D) + "| E".f(t.E) + "| F".f(t.F) + "|")
        sb.AppendLine(c + line + "|")

        Dim nl As String = ""
        For i = 0 To 5
            nl = c + w(i).f(t.A)
            nl += ("| " + cB(i)).f(t.B) + ("| " + cC(i)).f(t.C)
            nl += ("| " + cD(i)).f(t.D) + ("| " + cE(i)).f(t.E)
            nl += ("| " + cF(i)).f(t.F) + "|"
            sb.AppendLine(nl)
        Next
        sb.AppendLine(c + line + "|")

        Dim ret As String = sb.ToString
        Clipboard.SetText(ret)

        Return ret
    End Function

    Public Sub append(Optional ByVal s As String = "")
        Me.sb.Append(s)
    End Sub

    Public Sub newLine(Optional ByVal s As String = "")
        Me.sb.AppendLine(s)
    End Sub

    Public Sub clearAllLines()
        Me.sb.Clear()
    End Sub

    Public Overrides Function toString() As String
        Return Me.sb.ToString
    End Function

End Class