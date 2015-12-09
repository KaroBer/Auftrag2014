Option Strict On

Class AE
    Inherits blConst
    Property Text As String
    Property Datum As Date
    Property von As Date
    Property bis As Date
    Property Dauer As TimeSpan
    Property Wert As Double

    Public Sub New()
    End Sub

    Public Function getDauer() As TimeSpan
        Dim d As TimeSpan = Today - Today
        Try
            If Me.bis > Me.von Then d = Me.bis - Me.von
        Catch ex As Exception
            ' keine Meldung. Zurückgegeben wird "00:00:00"
        End Try
        Return d
    End Function

    Public Overrides Function ToString() As String
        Dim t As String = " | "
        Return Me.Text + t + Me.Datum.ToShortDateString + t + Me.von.ToShortTimeString + t + Me.bis.ToShortTimeString
    End Function

    Public Function ToStringZeitFixLength() As String
        Dim t As String = " | "
        Return StringFuellen(Me.Text, 28, True) + t + Me.Datum.ToShortDateString + t + Me.von.ToShortTimeString + t + Me.bis.ToShortTimeString + t
    End Function

    Public Function ToStringStueckFixLength() As String
        Dim t As String = " | "
        Return StringFuellen(Me.Text, 40, True) + t + Me.Datum.ToShortDateString + t + Me.Wert.ToString + t
    End Function


    Public Function toCSV() As String
        Dim t As String = CSVTRENNER
        Dim CSVBuilder As New System.Text.StringBuilder
        CSVBuilder.Append(Me.Text + t + Me.Datum.ToShortDateString + t + Me.von.ToShortTimeString + t + Me.bis.ToShortTimeString + t + Me.Dauer.ToString + t + Me.Wert.ToString + ZEILENENDE)
        Return CSVBuilder.ToString
    End Function

    Public Function CSVtoObject(ByVal CSVString As String) As AE
        Dim AE As New AE()
        Dim proptys() As String = Split(CSVString, CSVTRENNER)
        If proptys.Count > 5 Then
            Try
                AE.Text = proptys(0)
                AE.Datum = CDate(proptys(1))
                AE.von = CDate(proptys(2))
                AE.bis = CDate(proptys(3))
                AE.Dauer = AE.getDauer
                AE.Wert = CDbl(proptys(5))
            Catch ex As Exception
                MsgBox("AE.CSVtoObject: Fehler beim Umwandeln der AE-Eigenschaften." + vbCrLf + ex.ToString)
            End Try
        Else
            MsgBox("AE.CSVtoObject: AE hat zu wenige Eigenschaften")
            Return Nothing
        End If
        Return AE
    End Function


    'Hilfsfunktionen - hier war StringFuellen - das liegt aber jetzt in blHilfsModul2014.vb

End Class
