Option Strict On
Imports System.IO

Module DBLayerModule

    Function neuenAuftragSichern(ByVal a As Auftrag) As Boolean
        If a Is Nothing Then Return False

        ' Der Auftrag darf noch nicht existieren
        If File.Exists(a.DateiPfadUndName) Then
            MsgBox("neuenAuftragSichern: Der Auftrag existiert bereits!" + vbCrLf + a.DateiPfadUndName)
            Return False
        End If

        'ggf. Ausgabepfad anlegen
        If Not Directory.Exists(a.AusgabePfad) Then Directory.CreateDirectory(a.AusgabePfad)
        Return TextDateiSchreiben(a.DateiPfadUndName, a.toCSV, False)
    End Function

    Function geaendertenAuftragSichern(ByVal a As Auftrag) As Boolean
        'Der Auftrag muss bereits existieren
        If Not File.Exists(a.DateiPfadUndName) Then
            MsgBox("geaendertenAuftragSichern: Der Auftrag existiert nicht!" + vbCrLf + a.DateiPfadUndName)
            Return False
        End If

        'Geänderte Daten schreiben
        Return TextDateiSchreiben(a.DateiPfadUndName, a.toCSV, False)
    End Function

    Function AuftragEinLesen(ByVal Code As String) As Auftrag
        If Code.Length < 8 Then
            MsgBox("Das Einlesen von " + vbCrLf + Code + vbCrLf + "ist fehlgeschlagen!" + vbCrLf + "Der übergebene Code ist fehlerhaft!")
            Return Nothing
        End If

        Dim heft As String = Mid(Code, 1, 3)
        Dim ausgabe As String = Mid(Code, 4, 4)
        Dim DateiNameUndPfad As String = blConst.STANDARDPFAD + heft + "\" + heft + ausgabe + "\" + Code + blConst.DATEIENDUNG
        Dim a As New Auftrag

        If File.Exists(DateiNameUndPfad) Then
            Dim zeilen As New List(Of String)
            zeilen = TextZeilenLesen(DateiNameUndPfad, True)
            If zeilen.Count > 0 Then
                a = a.CSVtoObject(zeilen)
            End If
        Else
            MsgBox("Das Einlesen von " + vbCrLf + DateiNameUndPfad + vbCrLf + "ist fehlgeschlagen!" + vbCrLf + "Der Dateiname existiert nicht!")
        End If
        Return a
    End Function

    Function AuftragsDateiEinLesen(ByVal Code As String) As String
        If Code.Length < 8 Then
            MsgBox("Das Einlesen von " + vbCrLf + Code + vbCrLf + "ist fehlgeschlagen!" + vbCrLf + "Der übergebene Code ist fehlerhaft!")
            Return Nothing
        End If

        Dim heft As String = Mid(Code, 1, 3)
        Dim ausgabe As String = Mid(Code, 4, 4)
        Dim DateiNameUndPfad As String = blConst.STANDARDPFAD + heft + "\" + heft + ausgabe + "\" + Code + blConst.DATEIENDUNG
        Dim textDerAuftragsDatei As String = ""

        If File.Exists(DateiNameUndPfad) Then
            textDerAuftragsDatei = qTextDateiLesen(DateiNameUndPfad, True)
        Else
            MsgBox("Das Einlesen von " + vbCrLf + DateiNameUndPfad + vbCrLf + "ist fehlgeschlagen!" + vbCrLf + "Der Dateiname existiert nicht!")
        End If
        Return textDerAuftragsDatei
    End Function



    Function ListeDerUnterordner(ByVal Pfad As String) As List(Of String)
        Dim ListeDerOrdner As New List(Of String)
        If Directory.Exists(Pfad) Then
            'Ordner in Liste übernehmen
            Dim Ordner As String() = Directory.GetDirectories(Pfad)
            For Each o As String In Ordner
                ListeDerOrdner.Add(o)
            Next
        End If
        Return ListeDerOrdner
    End Function

    Function ListeDerDateiNamenImOrdner(ByVal Pfad As String, ByVal DateiEndung As String) As List(Of String)
        Dim ListeDerDateinamen As New List(Of String)
        If Directory.Exists(Pfad) Then
            'Dateien in Liste übernehmen
            Dim pFiles As String() = Directory.GetFiles(Pfad)
            For Each f As String In pFiles
                If f.EndsWith(DateiEndung.ToLower) Or f.EndsWith(DateiEndung.ToUpper) Then ListeDerDateinamen.Add(f)
            Next
        End If
        Return ListeDerDateinamen
    End Function

    Function TextPruefenUndSpeichern(
                                    ByVal nichtErlaubt As String,
                                    ByVal pfad As String,
                                    ByVal Text As String,
                                    ByVal Fehlermeldung As Boolean) As Boolean
        Dim ret As Boolean = False
        If Text <> nichtErlaubt Then
            ret = TextDateiSchreiben(pfad, Text, False)
        End If
        Return ret
    End Function

    Function TextDateiSchreiben(ByVal pfad As String, ByVal Text As String, ByVal fehlermeldung As Boolean) As Boolean
        Dim ret As Boolean = True
        Try
            File.WriteAllText(pfad, Text)
        Catch except As Exception
            ret = False
            If fehlermeldung Then MsgBox(except.Message & vbNewLine & "Fehler in TextDateiSchreiben()", MsgBoxStyle.Exclamation)
        End Try

        Return ret
    End Function

    Function TextDateiLesen(ByVal pfad As String, ByVal fehlermeldung As Boolean) As String
        Dim ret As String = ""
        If File.Exists(pfad) Then
            Try
                ret = File.ReadAllText(pfad)
            Catch except As Exception
                If fehlermeldung Then MsgBox(except.Message & vbNewLine & "Fehler in TextDateiLesen()", MsgBoxStyle.Exclamation)
            End Try
        End If
        Return ret
    End Function

    Function TextZeilenLesen(ByVal pfad As String, ByVal fehlermeldung As Boolean) As List(Of String)
        If File.Exists(pfad) Then
            Try
                Dim erg = File.ReadLines(pfad)
                Dim zeilen As New List(Of String)
                For Each z In erg
                    zeilen.Add(z.ToString)
                Next
                Return zeilen
            Catch except As Exception
                If fehlermeldung Then MsgBox(except.Message & vbNewLine & "Fehler in TextZeilenLesen()", MsgBoxStyle.Exclamation)
                Return Nothing
            End Try
        Else
            Return Nothing
        End If
    End Function

End Module
