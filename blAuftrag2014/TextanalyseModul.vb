Option Strict Off

Module TextAnalyseModul
    Function selectNumbers(ByVal s As String) As String()
        If Not containsDigit(s) Then Return Nothing

        Dim max As Integer = 5
        Dim erg(max) As String
        Dim i As Integer = -1
        Dim numString(2) As String

        s = numStringSaeubern(s)

        Do
            i += 1
            If i > max Then
                max = max + 10
                ReDim erg(max)
            End If
            numString = selectFirstNum(s)
            erg(i) = numString(0)
            s = numString(1)

            If containsDigit(s, False) = False Then Exit Do
            If i > 100 Then
                MsgBox("unsauberer Programmabruch nach 100 Durchläufen!!")
                Exit Do
            End If
        Loop

        Return erg
    End Function

    Function zaehleZeilenOhneZahl(ByVal zeilen() As String) As Integer
        If zeilen Is Nothing Then Return 0
        Dim anz As Integer = 0
        For Each z In zeilen
            If Not containsDigit(z) Then anz += 1
        Next
        Return anz
    End Function


    Private Function numStringSaeubern(ByVal s As String, Optional ByVal n As String = " ") As String
        s = s + "X"  ' wird unten wieder abgeschnitten
        Dim out As String = s
        Dim c() As Char = s.ToCharArray
        Dim pre As Char = CChar(" ")
        Dim changeflag As Boolean = False

        For i = 0 To c.Length - 2
            ' Array-Länge 10 : Index 0 ... 9 : '
            ' Ich guck ' noch einen Index voraus (i+1) also max 8 (length-2)
            ' Für jedes Element nachgucken, ob das nachfolgende Element identisch ist
            ' Falls ja, dann das aktuelle Element i ändern (in n, Vorgabe ist Leerzeichen)
            If c(i) = "+" Or c(i) = "-" Or c(i) = "." Then
                If c(i) = "+" Then
                    If c(i) = c(i + 1) Then
                        pre = c(i)
                        c(i) = CChar(n)
                        changeflag = True
                    Else
                        ' Beim letzten Wiederholungszeichen ist nur der Vorgänger dasselbe Zeichen
                        If c(i) = pre Then c(i) = CChar(n) : changeflag = True
                    End If
                End If

                If c(i) = "-" Then
                    If c(i) = c(i + 1) Then
                        pre = c(i)
                        c(i) = CChar(n)
                        changeflag = True
                    Else
                        ' Beim letzten Wiederholungszeichen ist nur der Vorgänger dasselbe Zeichen
                        If c(i) = pre Then c(i) = CChar(n) : changeflag = True
                    End If
                End If

                If c(i) = "." Then
                    If c(i) = c(i + 1) Then
                        pre = c(i)
                        c(i) = CChar(n)
                        changeflag = True
                    Else
                        ' Beim letzten Wiederholungszeichen ist nur der Vorgänger dasselbe Zeichen
                        If c(i) = pre Then c(i) = CChar(n) : changeflag = True
                    End If
                End If
            Else
                ' Ein anderes Zeichen als "+-."
                ' Alles außer NumberChars ersetzen
                If c(i) <> n Then
                    If Not isNumberChar(c(i)) Then c(i) = n : changeflag = True
                End If

                ' Auch die Kombination 5-5 oder 6+7 soll durch 5 5 und 6 7 erstetzt werden
                If isDigit(c(i), False) And i <= s.Length - 3 Then
                    If (c(i + 1) = "-" Or c(i + 1) = "+") And isDigit(c(i + 2), False) Then
                        c(i + 1) = " " : changeflag = True
                    End If
                End If
            End If
        Next

        If changeflag Then out = New String(c)
        ' Oben angehängtes Zeichen wieder abschneiden
        If out <> "" Then
            out = Mid(out, 1, out.Length - 1)
            ' Kein +, - oder . am Ende
            Dim tmp As String = Mid(out, out.Length)
            If tmp = "+" Or tmp = "-" Or tmp = "." Then
                Mid(out, out.Length) = n
            End If
        End If

        Return out
    End Function

    Private Function selectFirstNum(ByVal s As String) As String()
        ' Übergebenes s muss vorher gesäubert werden (numStringSaeubern)
        ' Es enthält dann nur noch die Zahlen und Leerzeichen
        If containsDigit(s) = False Then Return Nothing

        Dim erg(2) As String

        ' Erste Ziffer suchen ... dann erstes Leerzeichen dahinter suchen
        Dim c As String = ""
        Dim p1 As Integer = s.Length + 1
        Dim l1 As Integer = s.Length + 1

        For i = 1 To s.Length
            c = Mid(s, i, 1)
            If isDigit(c, True, True) Then
                If i < p1 Then p1 = i
            End If
            If c = " " Then
                ' Schon ein p1 gefunden? 
                If p1 < s.Length Then l1 = i : Exit For
            End If
        Next

        Dim numString As String = Mid(s, p1, l1 - p1)
        If isNumber(numString) Then
            erg(0) = numString
        Else
            erg(0) = ""
        End If
        erg(1) = Mid(s, l1 + 1)

        Return erg
    End Function

    Private Function containsDigit(ByVal s As String,
                                   Optional ByVal plusMinus As Boolean = True,
                                   Optional ByVal point As Boolean = False) As Boolean

        If s = "" Or s = Nothing Then Return False
        Dim tf As Boolean = False

        ' Auf Ziffern prüfen
        For Each c As Char In s.ToCharArray
            If isDigit(c, plusMinus, point) Then
                tf = True
                Exit For
            End If
        Next

        Return tf
    End Function

    Private Function isNumber(ByVal s As String) As Boolean
        Dim tf As Boolean = True

        ' Auf Ziffern, Komma und Punkt prüfen
        For Each c As Char In s.ToCharArray
            If Not isNumberChar(c) Then
                tf = False
                Exit For
            End If
        Next

        ' Komma darf nicht 2x vorkommen
        Dim x As Integer = InStr(s, ",")
        If tf And x > 0 And s.Length >= x Then
            If InStr(x + 1, s, ",") > 0 Then tf = False
            If InStr(x + 1, s, ".") > 0 Then tf = False
        End If

        ' Punkt darf nicht 2x vorkommen
        x = InStr(s, ".")

        If tf And x > 0 And s.Length >= x Then
            If InStr(x + 1, s, ".") > 0 Then tf = False
            If InStr(x + 1, s, ",") > 0 Then tf = False
        End If

        Return tf
    End Function

    Private Function isNumberChar(ByVal c As Char) As Boolean
        Return isDigit(c) Or isPlusMinus(c) Or isKomma(c) Or isPoint(c)
    End Function

    Private Function isPlusMinus(ByVal c As Char) As Boolean
        Dim tf As Boolean = False
        If c = CChar("+") Then tf = True
        If c = CChar("-") Then tf = True
        Return tf
    End Function

    Private Function isKomma(ByVal c As Char) As Boolean
        Dim tf As Boolean = False
        If c = CChar(",") Then tf = True
        Return tf
    End Function

    Private Function isPoint(ByVal c As Char) As Boolean
        Dim tf As Boolean = False
        If c = CChar(".") Then tf = True
        Return tf
    End Function

    Private Function isDigit(ByVal c As Char,
                             Optional ByVal plusMinus As Boolean = True,
                             Optional ByVal point As Boolean = False) As Boolean

        Dim tf As Boolean = False
        If Asc(c) >= 48 And Asc(c) <= 57 Then tf = True
        If plusMinus Then
            If Asc(c) = 43 Or Asc(c) = 45 Then tf = True ' + und -
        End If
        If point Then
            If c = CChar(".") Then tf = True
        End If
        Return tf
    End Function

End Module
