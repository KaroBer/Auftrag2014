Option Strict Off


Imports iTextSharp.text.pdf


Class MainWindow
    'Allgemeine Variablen
    Dim AktHeft As Auftrag.Heft = Auftrag.Heft.dnp
    Dim AktAusgabe As String = "1405"
    Dim WithEvents AktAuftrag As New Auftrag

    Dim AEzeileStueckStatus_NEU As Boolean = True
    Dim AEzeileZeitStatus_NEU As Boolean = True
    Dim AEPapierkorb As New AE

    Dim EingabeStatus As String = "Neu"

    Dim ci As Integer = 0

    ' Für PDF-Laufzettel
    Dim Arbeitspfad As String = "D:\SkyDrive\Arbeit\dotnetpro\"
    Dim OriginalPDF As String = Arbeitspfad + "TestPDF\DNP_Jobsheet_2015_Formular.pdf"
    Dim Kopie As String = ""
    Dim maxA1 As Integer = 75
    Dim maxA2bis4 As Integer = 90

    Dim ProgPfade As String()
    Dim TBinhalteBeiProgStart As String()

    Dim WerRuftAnProc As New Process


    ' Allgemeine Hilfsfunktionen sind in Modul blHilfsModul2014.vb
    ' Anwendungsbezogene Hilfsfunktionen

    Private Sub kwNewsZahlGeaendert(ByVal zahl As Integer)
        lbl_Zahl.Content = zahl.ToString
    End Sub

    Private Function aktKWSo() As Date
        Dim heut As DayOfWeek = Today.DayOfWeek
        Dim So As Date

        ' Wenn Sonntag, dann wird heut = 0
        ' Wochenanfang war der Montag vorher, deshalb anders rechnen!
        If heut = 0 Then
            So = Today.AddDays(-1)
        Else
            So = Today.AddDays(7 - heut)
        End If
        Return So
    End Function

    Private Function aktKWMo() As Date
        Dim heut As DayOfWeek = Today.DayOfWeek
        Dim Mo As Date

        ' Wenn Sonntag, dann wird heut = 0
        ' Wochenanfang war der Montag vorher, deshalb anders rechnen!
        If heut = 0 Then
            Mo = Today.AddDays(-6)
        Else
            Mo = Today.AddDays(-(heut - 1))
        End If
        Return Mo
    End Function

    Private Function alleTboxInhalte() As String()
        Dim alleTexte(10) As String
        alleTexte(1) = tb_Liste1.Text
        alleTexte(2) = tb_Liste2.Text
        alleTexte(3) = tb_Liste3.Text
        alleTexte(4) = tb_Liste4.Text
        alleTexte(5) = tb_Daten1.Text
        alleTexte(6) = tb_Liste6.Text
        alleTexte(7) = tb_Termine1.Text
        alleTexte(8) = tb_Termine2.Text

        Return alleTexte
    End Function

    Private Sub IconsEinfuegen()
        img1.Source = erstesIconHolen(ProgPfade(1)).Source
        img2.Source = erstesIconHolen(ProgPfade(2)).Source
        img3.Source = erstesIconHolen(ProgPfade(3)).Source
        img4.Source = erstesIconHolen(ProgPfade(4)).Source
        img5.Source = erstesIconHolen(ProgPfade(5)).Source
        img6.Source = erstesIconHolen(ProgPfade(6)).Source
        img7.Source = erstesIconHolen(ProgPfade(7)).Source
        'img8.Source = erstesIconHolen(ProgPfade(8)).Source
        'img9.Source = erstesIconHolen(ProgPfade(9)).Source
        'img10.Source = erstesIconHolen(ProgPfade(10)).Source
        'img11.Source = erstesIconHolen(ProgPfade(11)).Source
        img12.Source = erstesIconHolen(ProgPfade(12)).Source
        img13.Source = erstesIconHolen(ProgPfade(13)).Source
        img14.Source = erstesIconHolen(ProgPfade(14)).Source
        img15.Source = erstesIconHolen(ProgPfade(15)).Source
        img16.Source = erstesIconHolen(ProgPfade(16)).Source
        img17.Source = erstesIconHolen(ProgPfade(17)).Source
        img18.Source = erstesIconHolen(ProgPfade(18)).Source
        img19.Source = erstesIconHolen(ProgPfade(19)).Source
        'img20.Source = erstesIconHolen(ProgPfade(20)).Source
    End Sub

    Private Sub AuftragLoeschenInPapierkorb()
        ' Angeklickten Auftrag löschen
        Dim lAuftrag As String = lb_auftragsliste.SelectedItem.ToString
        Dim melde As String = "Wollen Sie den Auftrag " & lAuftrag & " wirklich in den Papierkorb verschieben?" &
            vbCrLf & "Das Wiederherstellen kann nur manuell durchgeführt werden!" &
            vbCrLf & "Gibt es im Papierkorb eine Datei gleichen Namens, wird diese überschrieben."

        Dim antw As MessageBoxResult = MsgBox(melde, MsgBoxStyle.YesNo, "Wirklich löschen?")
        If antw = vbYes Then
            Dim lDatei = AktAuftrag.AusgabePfad & lAuftrag & blConst.DATEIENDUNG
            Dim lDateiInPapierkorb = blConst.PAPIERKORB & AktAuftrag.Code & blConst.DATEIENDUNG
            MsgBox(lDatei & vbCrLf & lDateiInPapierkorb)
            If System.IO.File.Exists(lDatei) Then
                My.Computer.FileSystem.MoveFile(lDatei, lDateiInPapierkorb, overwrite:=True)
            End If
        End If
        'Liste der Aufträge neu einlesen, dabei wird ein anderer Auftrag zum aktAuftrag
        ListeDerAuftraegeInListBoxEinlesen()
    End Sub

    Private Sub AuftragUmbenennen()
        ' Angeklickten Auftrag umbenennen
        Dim uAuftragCode As String = lb_auftragsliste.SelectedItem.ToString
        ' Feste Elemente des neuen Dateipfades
        Dim uPfad As String = AktAuftrag.AusgabePfad
        Dim uAlterName As String = Mid(uAuftragCode, 8)
        Dim uNeuerName As String = ""
        Dim uNeuerCode As String = ""
        Dim ok As Boolean = False

        Do
            uNeuerName = InputBox("Alter Name: " & uAlterName & vbCrLf & "Neuen Namen eingeben:", "Umbenennen", uAlterName)
            If uNeuerName = "" Then Exit Do
            ' Prüfen, ob es den Namen schon gibt
            For Each code In lb_auftragsliste.Items
                If InStr(code.tolower, uNeuerName.ToLower) = 0 Then
                    ' Den Namen gibt's noch nicht, also alles ok!
                    ok = True
                Else
                    ' Den Namen gibt's bereits, er kann nicht verwendet werden.
                    ok = False
                    MsgBox("Der eingegebene Name existiert bereits. Bitte auf 'ok' klicken und dann einen neuen Namen eingeben.")
                    Exit For
                End If
            Next
        Loop Until ok = True

        'MsgBox("wir sind raus aus der do-loop-schleife")

        ' Es könnte ein Leerstring eingegeben worden sein, dann das Umbenennen überspringen
        If ok And uNeuerName <> "" Then
            uNeuerCode = Mid(uAuftragCode, 1, 7) & uNeuerName
            Dim uNeuerCodeDat As String = uNeuerCode & blConst.DATEIENDUNG
            Dim uAlteDatei = uPfad & uAuftragCode & blConst.DATEIENDUNG
            Dim uNeueDatei = uPfad & uNeuerCodeDat

            ' MsgBox(uAlteDatei + vbCrLf + uNeuerName + vbCrLf + "Umbenennen ist noch in Arbeit!")
            ' Umbenennen durchführen
            If System.IO.File.Exists(uAlteDatei) Then
                My.Computer.FileSystem.RenameFile(uAlteDatei, uNeuerCodeDat)
            Else
                MsgBox("Die Datei " & vbCrLf & uAlteDatei & vbCrLf & "existiert nicht!",, "Umbenennen fehlgeschlagen!")
                ok = False
            End If
            ' In der Datei steht der alte Name, der muss überschrieben werden
            Dim inhalt As String = TextDateiLesen(uNeueDatei, True)
            Dim wo As Integer = InStr(inhalt, uAlterName)
            If wo > 0 Then
                ' Name austauschen
                inhalt = Replace(inhalt, uAlterName, uNeuerName, 1, 1)
                TextDateiSchreiben(uNeueDatei, inhalt, True)
            Else
                MsgBox("Umbenennen: Fataler Fehler: Der alte Name in der Datei " & uNeueDatei &
                       " konnte nicht ersetzt werden. Es ist ein manueller Eingriff notwendig! Nach dem Klick auf OK wird der Ordner wird geöffnet.")
                starteProg(uPfad)
            End If

        End If
        'Liste der Aufträge neu einlesen, dabei wird ein anderer Auftrag zum aktAuftrag
        ListeDerAuftraegeInListBoxEinlesen()
    End Sub

    Private Sub VerschiebeAuftragInNaechstesHeft()
        'Angeklickten Auftrag feststellen
        Dim vAuftrag As String = lb_auftragsliste.SelectedItem.ToString

        'Name der neu anzulegenden Datei
        Dim vZielAusgabe As String = HeftNrPlus1(AktAusgabe)
        Dim vZielHeft As String = AktHeft.ToString + vZielAusgabe

        'Name der zu löschenden Datei
        Dim aktPfad As String = blConst.STANDARDPFAD + AktHeft.ToString + "\" + AktHeft.ToString + AktAusgabe + "\"
        Dim zuLoeschen As String = aktPfad + AktHeft.ToString + AktAusgabe + AktAuftrag.Text + blConst.DATEIENDUNG
        Dim deletedName As String = AktHeft.ToString + AktAusgabe + AktAuftrag.Text + ".jetzt_in_" + vZielAusgabe

        'Pfade zu den Dateien
        Dim neuerPfad As String = blConst.STANDARDPFAD + AktHeft.ToString + "\" + AktHeft.ToString + vZielAusgabe + "\"
        Dim neueDatei As String = neuerPfad + AktHeft.ToString + vZielAusgabe + AktAuftrag.Text + blConst.DATEIENDUNG

        'MsgBox(neuerPfad) : Stop

        'Ersetzen in der ersten Zeile der Datei vorbereiten
        Dim zuErsetzen As String = AktHeft.ToString + blConst.CSVTRENNER + AktAusgabe + blConst.CSVTRENNER
        Dim neuFuerErsteZeile As String = AktHeft.ToString + blConst.CSVTRENNER + vZielAusgabe + blConst.CSVTRENNER

        'Alle Zeilen der alten Datei einlesen
        Dim alteZeilen As New List(Of String)
        alteZeilen = TextZeilenLesen(zuLoeschen, True)

        'Neue Datei zusammenbauen
        Dim neueZeilen As New List(Of String)
        For Each z In alteZeilen
            If z = alteZeilen.First Then
                neueZeilen.Add(Replace(z, zuErsetzen, neuFuerErsteZeile))
            Else
                neueZeilen.Add(z)
            End If
        Next

        Dim melde As String = ""
        melde += "von" + vbCrLf + zuLoeschen + vbCrLf
        melde += "nach" + vbCrLf + neueDatei + vbCrLf + "Inhalt" + vbCrLf
        For Each x In neueZeilen
            melde += x + vbCrLf
        Next

        Dim antw As MessageBoxResult = MsgBox(melde, MsgBoxStyle.YesNo, "Wirklich verschieben?")
        If antw = vbYes Then

            If Not System.IO.Directory.Exists(neuerPfad) Then
                System.IO.Directory.CreateDirectory(neuerPfad)
            End If

            ' Neue Datei anlegen
            If Not System.IO.File.Exists(neueDatei) Then
                System.IO.File.WriteAllLines(neueDatei, neueZeilen)
            Else
                MsgBox("Fehler: Die Datei " + vbCrLf + neueDatei + vbCrLf + "existiert bereits." + vbCrLf + "Es wird nichts verschoben.")
                Exit Sub
            End If

            ' Alte Datei umbenennen in .jetzt_in_1510
            If System.IO.File.Exists(zuLoeschen) And Not System.IO.File.Exists(aktPfad + deletedName) Then
                My.Computer.FileSystem.RenameFile(zuLoeschen, deletedName)
            Else
                MsgBox("Alte Datei kann nicht gelöscht werden." + vbCrLf + "Neue Datei wurde bereits angelegt.")
            End If
        End If

        'Jetzt noch die Liste der Aufträge neu laden
        lbl_HeftNrplus1_MouseDown("VerschiebeAuftrag", Nothing)
        lbl_HeftNrminus1_MouseDown("VerschiebeAuftrag", Nothing)

    End Sub

    Private Sub TabWechseln(ByRef ti As TabItem)
        ' Funktioniert nur aus Mouseup heraus, nicht bei Mousedown !!
        ti.Focus()
    End Sub

    Private Sub wumLabelHervorheben()
        lbl_wum.Background = Brushes.SeaShell
        lbl_wum.Foreground = Brushes.DarkRed
        lbl_wum.BorderBrush = Brushes.Red
        lbl_dnp.Background = Brushes.Transparent
        lbl_dnp.Foreground = Brushes.SeaShell
        lbl_dnp.BorderBrush = Brushes.DarkRed
        lbl_com.Background = Brushes.Transparent
        lbl_com.Foreground = Brushes.Green
        lbl_com.BorderBrush = Brushes.DarkRed
    End Sub

    Private Sub dnpLabelHervorheben()
        lbl_dnp.Background = Brushes.SeaShell
        lbl_dnp.Foreground = Brushes.DarkRed
        lbl_dnp.BorderBrush = Brushes.Red
        lbl_wum.BorderBrush = Brushes.DarkRed
        lbl_wum.Background = Brushes.Transparent
        lbl_wum.Foreground = Brushes.SeaShell
        lbl_com.Background = Brushes.Transparent
        lbl_com.Foreground = Brushes.Green
        lbl_com.BorderBrush = Brushes.DarkRed
    End Sub

    Private Sub comLabelHervorheben()
        lbl_com.Background = Brushes.SeaShell
        lbl_com.Foreground = Brushes.DarkRed
        lbl_com.BorderBrush = Brushes.Red
        lbl_wum.BorderBrush = Brushes.DarkRed
        lbl_wum.Background = Brushes.Transparent
        lbl_wum.Foreground = Brushes.SeaShell
        lbl_dnp.Background = Brushes.Transparent
        lbl_dnp.Foreground = Brushes.SeaShell
        lbl_dnp.BorderBrush = Brushes.DarkRed
    End Sub

    Private Function HeftNrPlus1(ByVal nr As String) As String
        Dim jahr As Integer = 0
        Dim monat As Integer = 0
        Try
            jahr = CInt(Mid(nr, 1, 2))
            monat = CInt(Mid(nr, 3, 2))
        Catch ex As Exception

        End Try
        If monat = 12 Then
            monat = 1 : jahr = jahr + 1
        Else
            monat = monat + 1
        End If
        Dim m As String = "0" + monat.ToString
        If m.Length > 2 Then m = Mid(m, 2, 2)
        Return jahr.ToString + m
    End Function

    Private Function HeftNrMinus1(ByVal nr As String) As String
        Dim jahr As Integer = 0
        Dim monat As Integer = 0
        Try
            jahr = CInt(Mid(nr, 1, 2))
            monat = CInt(Mid(nr, 3, 2))
        Catch ex As Exception

        End Try
        If monat = 1 Then
            monat = 12 : jahr = jahr - 1
        Else
            monat = monat - 1
        End If
        Dim m As String = "0" + monat.ToString
        If m.Length > 2 Then m = Mid(m, 2, 2)
        Return jahr.ToString + m
    End Function

    Private Function kompletterPfadZumAktuellenHeft() As String
        Return blConst.STANDARDPFAD + AktHeft.ToString + "\" + AktHeft.ToString + AktAusgabe + "\"
    End Function

    Private Function kompletterPfadZumHeft(ByVal Heft As String, Ausgabe As String) As String
        Return blConst.STANDARDPFAD + Heft + "\" + Heft + Ausgabe + "\"
    End Function

    Public Function standardTexte() As String
        Dim t As String = ez_tbtext.Text
        ' Texte per Doppelklick abrufen
        If t = "Red 2" Then t = "Schlussred"
        If t = "Zweites Überarbeiten" Then t = "Red 2"
        If t = "Erstes Überarbeiten" Then t = "Zweites Überarbeiten"
        If t = "" Then t = "Erstes Überarbeiten"
        Return t
    End Function

    Function zeileInStatistikDateiTextTauschen(ByVal tDateiInhalt As String, muster As String, neuerText As String, Optional ByVal ObenEinfuegen_wenn_nicht_drin As Boolean = True) As String
        'Voraussetzung: Ganze Zeilen werden getauscht. Die Zeilen enden mit vbcrlf
        '               Das Muster muss nicht die komplette Zeile enthalten.
        '               Anwendungsfall: Zeile 20131120:90
        '                   austauschen gegen 20131120:145 
        '                 dabei ist muster = "20131120"
        '                   und neuer Text = "20131120: 145" + vbcrlf

        'Testen, ob das Datum schon in der Datei ist
        Dim tmp As String = ""
        If InStr(tDateiInhalt, muster) < 1 Then
            'nicht drin, ggf. oben dranhängen
            If ObenEinfuegen_wenn_nicht_drin Then tDateiInhalt = neuerText + tDateiInhalt

        Else ' Datum ist schon drin

            'direkt im String tDateiInhalt austauschen
            'alten Wert greifen
            Dim start As Integer = InStr(tDateiInhalt, muster)
            If start > 0 Then
                Dim ende As Integer = InStr(start, tDateiInhalt, vbCrLf) + 2
                Dim alt As String = Mid(tDateiInhalt, start, ende - start)
                tDateiInhalt = Replace(tDateiInhalt, alt, neuerText)
            End If
        End If

        Return tDateiInhalt
    End Function

    Private Function WertAusStatistikDateiTextHolen(ByVal tDateiInhalt As String, ByVal key As String) As String
        Dim ReturnWert As String = ""
        Dim Datei As String = tDateiInhalt
        Dim pTag As Integer = InStr(Datei, key)

        If pTag > 0 Then
            ' ist schon drin
            Dim pEnde As Integer = InStr(pTag, Datei, vbCrLf)
            Dim zeile As String = Mid(Datei, pTag, pEnde - pTag) + vbCrLf
            ReturnWert = ZweiStellenHintermKomma(WertAusStatistikZeileHolen(zeile, key))
        Else
            ' ist noch nicht drin
            ReturnWert = "0,00"
        End If
        Return ReturnWert
    End Function

    Private Function DatenFuerChartHolen(ByVal twm As String) As Dictionary(Of String, Double)
        ' t = Daten aus TagesStatistikdatei
        ' m = Monat
        ' w = Woche
        ' r = Abrechnungen summiert dnp und wum aus Liste3
        ' com = Daten für Jobgrafik in der com
        '       Funktioniert auch mit anderen Daten

        Dim daten As New Dictionary(Of String, Double)
        Dim Trenner As String = ":"
        Dim key As String = ""
        Dim value As Double = 0


        Select Case twm
            Case Is = "t"
                Dim tagStatistik As New List(Of String)
                tagStatistik = TextZeilenLesen(blConst.TAGESDATEIPFAD, False)

                For i = 0 To 9
                    key = Mid(tagStatistik.Item(i), 5, 4)
                    value = CDbl(Mid(tagStatistik.Item(i), 10))
                    daten.Add(key, value)
                Next

            Case Is = "w"
                Dim kwStatistik As New List(Of String)
                kwStatistik = TextZeilenLesen(blConst.KWDATEIPFAD, False)

                For i = 0 To 9
                    key = Mid(kwStatistik.Item(i), 6, 2)
                    value = CDbl(Mid(kwStatistik.Item(i), 9))
                    daten.Add(key, value)
                Next

            Case Is = "m"
                Dim mStatistik As New List(Of String)
                mStatistik = TextZeilenLesen(blConst.MONATSDATEIPFAD, False)

                For i = 0 To 9
                    key = Mid(mStatistik.Item(i), 1, 7)
                    value = CDbl(Mid(mStatistik.Item(i), 9))
                    daten.Add(key, value)
                Next

            Case Is = "r"
                Dim rStatistik As New List(Of String)
                rStatistik = TextZeilenLesen(blConst.STANDARDPFAD + "Liste3" + blConst.DATEIENDUNG, False)

                For i = 0 To 9
                    If rStatistik.Item(i).Length > 15 Then
                        key = Mid(rStatistik.Item(i), 1, 10)
                        value = CDbl(Mid(rStatistik.Item(i), 12))
                        daten.Add(key, value)
                    Else
                        ' leere Zeile wird übersprungen
                    End If
                Next

            Case Is = "com"
                Dim comStatistik As New List(Of String)
                ' Die Daten dafür stehen im Textfeld tb_Daten1
                ' und haben die Form: key <tab> prozentwert

                Dim zeilen() As String
                Try
                    zeilen = Split(tb_Daten1.Text, vbCrLf)

                    Dim tb As Integer = 0 ' Tab
                    Dim ws As Integer = 0 ' Whitespace
                    For Each z As String In zeilen
                        z = Replace(z, "%", "")
                        tb = InStr(z, Chr(9))
                        If tb < 1 Then tb = InStr(z, (Chr(32)))

                        If tb > 0 Then
                            key = Mid(z, 1, tb - 1)
                            Dim num() As String
                            num = TextAnalyseModul.selectNumbers(Mid(z, tb + 1))
                            If num(1) <> "" Then
                                value = CDbl(num(1))
                            ElseIf num(0) <> "" Then
                                value = CDbl(num(0))
                            Else
                                value = 0
                            End If

                            'MsgBox(key + ": " + value.ToString)
                            Try
                                daten.Add(key, value)
                            Catch ex As Exception
                                ' Es fehlt wohl ein eindeutiger Schlüssel
                                daten.Add(key + value.ToString, value)
                            End Try

                        Else
                            ' leere Zeilen überspringen
                        End If
                    Next
                Catch ex As Exception
                    daten.Add("kDv", 0)
                    Return daten
                    Exit Function
                End Try

            Case Else
        End Select

        If daten.Count = 0 Then daten.Add("kDv", 0)
        Return daten
    End Function

    Private Function WertAusStatistikDateiHolen(ByVal pfad As String, ByVal key As String) As String
        Dim ReturnWert As String = ""

        Dim Datei As String = TextDateiLesen(pfad, False)
        Dim pTag As Integer = InStr(Datei, key)

        If pTag > 0 Then
            ' ist schon drin
            Dim pEnde As Integer = InStr(pTag, Datei, vbCrLf)
            Dim zeile As String = Mid(Datei, pTag, pEnde - pTag) + vbCrLf
            ReturnWert = ZweiStellenHintermKomma(WertAusStatistikZeileHolen(zeile, key))
        Else
            ' ist noch nicht drin
            ReturnWert = "0,00"
        End If
        Return ReturnWert
    End Function

    Private Function WertAusStatistikZeileHolen(ByVal data As String, ByVal key As String) As String
        ' Die Daten haben die Form    20131121:117,75 oder 2013_11:508 oder 2013:13210
        ' Der Schlüssel hat die Form  20131121:       oder 2013_11:    oder 2013:
        ' Am Zeilenende steht VbCrLf (2 Zeichen lang)
        Dim ret As String = ""
        Dim start As Integer = InStr(data, key) + key.Length
        Dim ende As Integer = InStr(start, data, vbCrLf)
        If start > 0 Then
            ret = Mid(data, start, ende - start)
            Return ret
        Else
            Return "-1"
        End If
    End Function

    Private Function DurchschnittBerechnen(ByVal StatistikDatei As String, ByVal AnzahlEinträge As Integer, Optional ByVal OhneWE As Boolean = False) As Double
        Dim summe As Double = 0.0
        Dim schnitt As Double = 0.0
        Dim AlleZeilen As List(Of String) = TextZeilenLesen(StatistikDatei, False)
        If AnzahlEinträge > AlleZeilen.Count - 1 Then AnzahlEinträge = AlleZeilen.Count - 1

        Dim Zeile As String = ""
        Dim pos As Integer = 0
        Dim dateString As String = ""
        Dim datum As Date = Today
        Dim Wert As Double = 0
        Dim WEKorrektur As Integer = 0

        For i = 1 To AnzahlEinträge
            ' Der jüngste Eintrag wird nicht mit in den Durchschnitt aufgenommen! Deshalb i=1 und nicht i=0
            Zeile = AlleZeilen(i)
            pos = InStr(Zeile, ":")

            If pos > 1 And pos < Zeile.Length Then
                Wert = CDbl(Replace(Mid(Zeile, pos + 1), vbCrLf, ""))
                summe = summe + Wert

                If OhneWE Then
                    dateString = Mid(Zeile, 7, 2) + "." + Mid(Zeile, 5, 2) + "." + Mid(Zeile, 1, 4)
                    Try
                        datum = CDate(dateString)
                    Catch ex As Exception
                        datum = Today
                    End Try
                    If IstWE(datum) Or istFixFeiertagInBayern(datum) Then WEKorrektur = WEKorrektur + 1
                End If
            End If
        Next
        schnitt = Math.Round(summe / (AnzahlEinträge - WEKorrektur), 0)

        Return schnitt
    End Function

    Private Function BasisHeft(ByVal d As Date, Optional ByVal delta As Integer = 3) As String
        Dim Basis As String = AktAusgabe
        Dim y As Integer = d.Year - 2000
        Dim m As Integer = d.Month + delta
        If m > 12 Then
            m = m - 12
            y = y + 1
        End If
        Basis = y.ToString + NullVorWert(m.ToString, 2)
        Return Basis
    End Function

    Private Function RelevanteListeFuerAuswertung(ByVal von As Date) As List(Of String)
        Dim Basis As String = BasisHeft(von)
        Dim trenner As String = "-".Mal(50)

        ' Container für die Heftsummen vorher, aktuell, nachher
        Dim dnp_00_Name As String = ""
        Dim dnp_01_Name As String = ""
        Dim dnp_02_Name As String = ""

        Dim wum_00_Name As String = ""
        Dim wum_01_Name As String = ""
        Dim wum_02_Name As String = ""

        Dim com_00_Name As String = ""
        Dim com_01_Name As String = ""
        Dim com_02_Name As String = ""

        Dim liste As New List(Of String)
        ' dnp Ausgabe vorher
        liste = ListeDerAuftraegeInListeEinlesen("dnp", HeftNrMinus1(Basis), liste)
        ' dnp, aktuelle Ausgabe 
        liste = ListeDerAuftraegeInListeEinlesen("dnp", Basis, liste)
        'dnp-Ausgabe nacher
        liste = ListeDerAuftraegeInListeEinlesen("dnp", HeftNrPlus1(Basis), liste)
        liste.Add("dnp" + trenner)

        ' Wum-Ausgabe vorher
        liste = ListeDerAuftraegeInListeEinlesen("wum", HeftNrMinus1(Basis), liste)
        ' wum, aktuelle Ausgabe
        liste = ListeDerAuftraegeInListeEinlesen("wum", Basis, liste)
        ' Wum-Ausgabe nachher
        liste = ListeDerAuftraegeInListeEinlesen("wum", HeftNrPlus1(Basis), liste)
        liste.Add("wum" + trenner)

        ' com Ausgabe vorher
        liste = ListeDerAuftraegeInListeEinlesen("com", HeftNrMinus1(Basis), liste)
        ' com, aktuelle Ausgabe 
        liste = ListeDerAuftraegeInListeEinlesen("com", Basis, liste)
        ' com-Ausgabe nacher
        liste = ListeDerAuftraegeInListeEinlesen("com", HeftNrPlus1(Basis), liste)

        Return liste
    End Function

    Private Function NewsZaehlenKW(ByVal heft As Auftrag.Heft, ByVal von As Date, ByVal bis As Date,
                                   Optional ByVal RelevanteListe As List(Of String) = Nothing)

        Dim trenner As String = "-".Mal(50)
        Dim anz As Integer = 0
        Dim liste As New List(Of String)

        If RelevanteListe Is Nothing Then
            liste = RelevanteListeFuerAuswertung(von)
        Else
            liste = RelevanteListe
        End If

        Dim countDNPnews As Integer = 0
        Dim countWUMnews As Integer = 0

        Dim codeDict As New Dictionary(Of String, Double)
        'Liste durchgucken, Werte sammeln
        For Each kode As String In liste
            If Mid(kode, 4) <> trenner Then
                'Auftrag einlesen
                Dim a As New Auftrag
                a = AuftragEinLesen(kode) : If a Is Nothing Then Return anz
                Dim tmp = a.Kunde.ToString + a.Ausgabe
                For Each ae As AE In a.AEs
                    If ae.Datum >= von And ae.Datum <= bis Then
                        If CBool(CInt(a.Art = Auftrag.Typ.st) And InStr(a.Text.ToLower, "news")) Then
                            If a.Kunde = Auftrag.Heft.dnp Then
                                countDNPnews += 1
                            Else
                                countWUMnews += 1
                            End If
                        End If
                    End If
                Next
            End If
        Next
        If heft = Auftrag.Heft.wum Then anz = countWUMnews Else anz = countDNPnews

        Return anz
    End Function

    Private Sub AuswertungVonBis(ByVal von As Date, ByVal bis As Date, Optional ByVal kurz As Boolean = False, Optional ByVal anhaengen As Boolean = False, Optional ByVal Text As String = "")
        Dim Basis As String = BasisHeft(von)
        Dim summe As Double = 0

        Dim trenner As String = "-".Mal(50)
        Dim sDNPnews As New System.Text.StringBuilder
        Dim sWUMnews As New System.Text.StringBuilder
        Dim countDNPnews As Integer = 0
        Dim countWUMnews As Integer = 0

        ' Container für die Heftsummen vorher, aktuell, nachher
        Dim dnp_00_Name As String = ""
        Dim dnp_01_Name As String = ""
        Dim dnp_02_Name As String = ""
        Dim dnp_00_Summe As Double = 0
        Dim dnp_01_Summe As Double = 0
        Dim dnp_02_Summe As Double = 0

        Dim wum_00_Name As String = ""
        Dim wum_00_Summe As Double = 0
        Dim wum_01_Name As String = ""
        Dim wum_01_Summe As Double = 0
        Dim wum_02_Name As String = ""
        Dim wum_02_Summe As Double = 0

        Dim com_00_Name As String = ""
        Dim com_00_Summe As Double = 0
        Dim com_01_Name As String = ""
        Dim com_01_Summe As Double = 0
        Dim com_02_Name As String = ""
        Dim com_02_Summe As Double = 0

        Dim HeftAusgabe_Summe_NichtZugeordnet As Double = 0

        Dim liste As New List(Of String)
        liste = RelevanteListeFuerAuswertung(von)

        ' dnp Ausgabe vorher
        dnp_00_Name = "dnp" + HeftNrMinus1(Basis)
        ' dnp, aktuelle Ausgabe 
        dnp_01_Name = "dnp" + Basis
        'dnp-Ausgabe nacher
        dnp_02_Name = "dnp" + HeftNrPlus1(Basis)

        ' Wum-Ausgabe vorher
        wum_00_Name = "wum" + HeftNrMinus1(Basis)
        ' wum, aktuelle Ausgabe
        wum_01_Name = "wum" + Basis
        ' Wum-Ausgabe nachher
        wum_02_Name = "wum" + HeftNrPlus1(Basis)

        ' com Ausgabe vorher
        com_00_Name = "com" + HeftNrMinus1(Basis)
        ' com, aktuelle Ausgabe 
        com_01_Name = "com" + Basis
        ' com-Ausgabe nacher
        com_02_Name = "com" + HeftNrPlus1(Basis)

        Dim DNPzwiSumme As Double = 0
        Dim WUMzwiSumme As Double = 0
        Dim COMzwiSumme As Double = 0

        Dim codeDict As New Dictionary(Of String, Double)
        'Liste durchgucken, Werte sammeln
        For Each kode As String In liste
            If Mid(kode, 4) <> trenner Then
                'Auftrag einlesen
                Dim a As New Auftrag
                a = AuftragEinLesen(kode) : If a Is Nothing Then Exit Sub
                Dim tmp = a.Kunde.ToString + a.Ausgabe
                For Each ae As AE In a.AEs
                    If ae.Datum >= von And ae.Datum <= bis Then
                        summe = summe + ae.Wert
                        If codeDict.ContainsKey(kode) Then
                            codeDict(kode) = codeDict.Item(kode) + ae.Wert
                        Else
                            codeDict.Add(kode, ae.Wert)
                        End If

                        ' Den ae.Wert einem Heft+Ausgabe zuordnen und summieren
                        Select Case tmp
                            Case Is = dnp_00_Name
                                dnp_00_Summe += ae.Wert
                            Case Is = dnp_01_Name
                                dnp_01_Summe += ae.Wert
                            Case Is = dnp_02_Name
                                dnp_02_Summe += ae.Wert
                            Case Is = wum_00_Name
                                wum_00_Summe += ae.Wert
                            Case Is = wum_01_Name
                                wum_01_Summe += ae.Wert
                            Case Is = wum_02_Name
                                wum_02_Summe += ae.Wert
                            Case Is = com_00_Name
                                com_00_Summe += ae.Wert
                            Case Is = com_01_Name
                                com_01_Summe += ae.Wert
                            Case Is = com_02_Name
                                com_02_Summe += ae.Wert
                            Case Else
                                HeftAusgabe_Summe_NichtZugeordnet += ae.Wert
                        End Select

                        If CBool(CInt(a.Art = Auftrag.Typ.st) And InStr(a.Text.ToLower, "news")) Then
                            If a.Kunde = Auftrag.Heft.dnp Then
                                sDNPnews.Append(ae.Datum.ToShortDateString + "  " + ae.Text + vbCrLf)
                                countDNPnews += 1
                            Else
                                sWUMnews.Append(ae.Datum.ToShortDateString + "  " + ae.Text + vbCrLf)
                                countWUMnews += 1
                            End If
                        End If
                    End If
                Next
            Else
                codeDict.Add(kode, 0.0)
                If Mid(kode, 1, 3) = "dnp" Then
                    DNPzwiSumme = summe
                ElseIf Mid(kode, 1, 3) = "wum" Then
                    WUMzwiSumme = summe - DNPzwiSumme
                End If
            End If
        Next
        COMzwiSumme = summe - DNPzwiSumme - WUMzwiSumme

        Dim tmp_sum As Double
        Dim t As New myTable("11 fill .",
                             "14 euro fill .",
                             "10 right fill .",
                             "15 euro fill .")
        With t
            If anhaengen Then .append(tb_Statistik.Text)
            .newLine("Zusammenfassung vom " + von.ToShortDateString + " bis zum " + bis.ToShortDateString + " (Basis: " + Basis + ")")
            If Text <> "" Then .newLine(Text)
            .newLine(trenner)

            .ABC.align = myColumn.alignTyp.left

            ' Verteilung der Monatswerte auf die Hefte
            If dnp_00_Summe > 0 Then .append((dnp_00_Name & " ").f(.A) + dnp_00_Summe.f(.B))
            If wum_00_Summe > 0 Then .append(vbCrLf + (wum_00_Name & " ").f(.A) + wum_00_Summe.f(.B))
            If com_00_Summe > 0 Then .append(vbCrLf + (com_00_Name & " ").f(.A) + com_00_Summe.f(.B))
            tmp_sum = dnp_00_Summe + wum_00_Summe + com_00_Summe
            If tmp_sum > 0 Then .newLine(" Summe:".f(.C) + tmp_sum.f(.D))
            tmp_sum = 0

            If dnp_01_Summe > 0 Then .append((dnp_01_Name & " ").f(.A) + dnp_01_Summe.f(.B))
            If wum_01_Summe > 0 Then .append(vbCrLf + (wum_01_Name & " ").f(.A) + wum_01_Summe.f(.B))
            If com_01_Summe > 0 Then .append(vbCrLf + (com_01_Name & " ").f(.A) + com_01_Summe.f(.B))
            tmp_sum = dnp_01_Summe + wum_01_Summe + com_01_Summe
            If tmp_sum > 0 Then .newLine(" Summe:".f(.C) + tmp_sum.f(.D))
            tmp_sum = 0

            If dnp_02_Summe > 0 Then .append((dnp_02_Name & " ").f(.A) + dnp_02_Summe.f(.B))
            If wum_02_Summe > 0 Then .append(vbCrLf + (wum_02_Name & " ").f(.A) + wum_02_Summe.f(.B))
            If com_02_Summe > 0 Then .append(vbCrLf + (com_02_Name & " ").f(.A) + com_02_Summe.f(.B))
            tmp_sum = dnp_02_Summe + wum_02_Summe + com_02_Summe
            If tmp_sum > 0 Then .newLine(" Summe:".f(.C) + tmp_sum.f(.D))
            tmp_sum = 0

            .newLine(trenner)
            tmp_sum = dnp_00_Summe + dnp_01_Summe + dnp_02_Summe + wum_00_Summe + wum_01_Summe + wum_02_Summe + com_00_Summe + com_01_Summe + com_02_Summe
            .newLine("Insgesamt".f(.ABC) + tmp_sum.f(.D))
            tmp_sum = 0

            .newLine()
            If Not kurz Then
                Dim key As String = ""
                Dim wert As Double = 0.0
                For Each d As KeyValuePair(Of String, Double) In codeDict
                    key = d.Key : wert = d.Value
                    If Mid(key, 4) <> trenner Then
                        .newLine((key + " ").f(.ABC) + wert.f(.D))
                    ElseIf Mid(key, 1, 3) = "dnp" Then
                        .newLine(trenner)
                        .newLine("Zwischensumme dnp ".f(.ABC) + DNPzwiSumme.f(.D))
                        .newLine()

                    ElseIf Mid(key, 1, 3) = "wum" Then
                        .newLine(trenner)
                        .newLine("Zwischensumme wum ".f(.ABC) + WUMzwiSumme.f(.D))
                        .newLine()
                    End If
                Next
                .newLine(trenner)
                .newLine("Zwischensumme com ".f(.ABC) + COMzwiSumme.f(.D))
                .newLine()

                .newLine(trenner)
                .newLine("Gesamtsumme ".f(.ABC) + summe.f(.D))
                .newLine("=".Mal(.ABCD.width))

                .newLine(vbCrLf.Mal(2))
                .newLine("--dnp-News--(" + countDNPnews.ToString + " Stück)--")
                .newLine(sDNPnews.ToString)
                .newLine("--wum-News--(" + countWUMnews.ToString + " Stück)--")
                .newLine(sWUMnews.ToString)
            End If
            tb_Statistik.Text = .toString
        End With

        tb_Statistik.CaretIndex = 1
        ti_Statistik.Focus()
    End Sub

    Private Sub AuswertungHeftAusgabe(ByVal code As String)
        Dim summe As Double = 0
        Dim s As New System.Text.StringBuilder
        Dim trenner As String = xMal("-", 50) + vbCrLf
        Dim Heft As String = Mid(code, 1, 3)
        Dim Ausgabe As String = Replace(code, Heft, "")

        Dim sNews As New System.Text.StringBuilder
        Dim CountNews As Integer = 0

        s.Append(vbCrLf + "Details zur Abrechnung " + code)
        s.Append(vbCrLf + trenner)
        s.Append(StringFuellen("Artikel-Code", 30))
        s.Append(StringFuellen("Art", 6))
        s.Append(StringFuellen("Anz", 8))
        s.Append(StringFuellen("Betrag", 10))
        s.Append(vbCrLf + trenner)

        Dim liste As New List(Of String)
        liste = ListeDerAuftraegeInListeEinlesen(Heft, Ausgabe, liste)

        Dim codeDict As New Dictionary(Of String, Double)
        'Liste durchgucken, Werte sammeln
        For Each kode As String In liste
            'Auftrag einlesen
            Dim a As New Auftrag
            a = AuftragEinLesen(kode)
            If a IsNot Nothing Then
                summe = summe + a.AktWert
                s.Append(a.toDetailString + vbCrLf)

                If CBool(CInt(a.Art = Auftrag.Typ.st) And InStr(a.Text.ToLower, "news")) Then
                    CountNews = CInt(a.Bisher)
                    For Each AE As AE In a.AEs
                        sNews.Append(AE.Datum.ToShortDateString + "  " + AE.Text + vbCrLf)
                    Next
                End If
            End If
        Next

        s.Append(trenner)
        s.Append(StringFuellen("Gesamtsumme ", 40, False, " ", False))
        s.Append(StringFuellen(StellenHintermKomma(summe, 2), 10, vonVorne:=True))
        s.Append(" Euro" + vbCrLf)
        s.Append(xMal("=", 50) + vbCrLf)
        'MsgBox(s.ToString)

        s.Append(vbCrLf + vbCrLf)
        s.Append("--" + Heft + "-News--(" + CountNews.ToString + " Stück)--" + vbCrLf)
        s.Append(sNews.ToString + vbCrLf)

        tb_Statistik.Text = s.ToString
        ti_Statistik.Focus()

    End Sub

    'Verhalten des Programms / Eigene Events
    Private Sub ProgrammStart()
        TBinhalteBeiProgStart = alleTboxInhalte()

        Me.Title = "blAuftrag 2015 | 4.12.15 /bl - Bessere Kontraste - VS2015"
        'Me.Title = "blAuftrag 2015 | 25.11.15 /bl - WerRuftAn start/ende - VS2015"
        'Me.Title = "blAuftrag 2015 | 19.11.15 /bl - Datum rechts oben - VS2015"
        'Me.Title = "blAuftrag 2015 | 18.11.15 /bl - NewsZahl anzeigen - VS2015"
        'Me.Title = "blAuftrag 2015 | 5.11.15 /bl - Umbenennen & Löschen - VS2015"
        'Me.Title = "blAuftrag 2015 | 5.11.15 /bl - PDFLaufzettel: Clear - VS2015"
        'Me.Title = "blAuftrag 2015 | 15.10.15 /bl - Suche news:Thema - VS2015"
        'Me.Title = "blAuftrag 2015 | 14.8.15 /bl - ProgStarter - VS2015"
        Me.Top = 70
        Me.Left = 1520

        tb_Liste1.Text = TextDateiLesen(blConst.STANDARDPFAD + "Liste1" + blConst.DATEIENDUNG, False)
        tb_Liste2.Text = TextDateiLesen(blConst.STANDARDPFAD + "Liste2" + blConst.DATEIENDUNG, False)
        tb_Liste3.Text = TextDateiLesen(blConst.STANDARDPFAD + "Liste3" + blConst.DATEIENDUNG, False)
        tb_Liste4.Text = TextDateiLesen(blConst.STANDARDPFAD + "Liste4" + blConst.DATEIENDUNG, False)
        tb_Liste6.Text = TextDateiLesen(blConst.STANDARDPFAD + "Liste6" + blConst.DATEIENDUNG, False)
        tb_Daten1.Text = TextDateiLesen(blConst.STANDARDPFAD + "Liste5" + blConst.DATEIENDUNG, False)
        tb_Termine1.Text = TextDateiLesen(blConst.STANDARDPFAD + "Termine1" + blConst.DATEIENDUNG, False)
        tb_Termine2.Text = TextDateiLesen(blConst.STANDARDPFAD + "Termine2" + blConst.DATEIENDUNG, False)

        Dim acode As String = TextDateiLesen(blConst.STANDARDPFAD + blConst.AKTAUFTRAGDATEI + blConst.DATEIENDUNG, True)

        If acode.Length > 6 Then
            AktHeft = CType([Enum].Parse(GetType(Auftrag.Heft), Mid(acode, 1, 3)), Auftrag.Heft)
            If AktHeft = Auftrag.Heft.dnp Then
                dnpLabelHervorheben()
            ElseIf AktHeft = Auftrag.Heft.wum Then
                wumLabelHervorheben()
            ElseIf AktHeft = Auftrag.Heft.com Then
                comLabelHervorheben()
            End If
            AktAusgabe = Mid(acode, 4, 4)
            lbl_HeftNrAktuell.Content = AktAusgabe
            lbl_HeftNrminus1.Content = HeftNrMinus1(AktAusgabe)
            lbl_HeftNrplus1.Content = HeftNrPlus1(AktAusgabe)
        Else
            MsgBox("Fehler bei: " + AktHeft.ToString)
            ' Fehler beim Einlesen, Vorgaben weiterverwenden
        End If
        ListeDerAuftraegeInListBoxEinlesen()

        TagStatistikAendern(0.0, Today)
        KWStatistikAendern(0.0, Today)
        MonatStatistikAendern(0.0, Today)
        JahrStatistikAendern(0.0, Today)

        Dim v As String : Dim b As String
        b = HeftNrPlus1(AktAusgabe)
        v = HeftNrMinus1(HeftNrMinus1(HeftNrMinus1(AktAusgabe)))
        tb_Matrix_von.Text = v
        tb_Matrix_bis.Text = b

        PDFVorgabenEintragen()

        ProgPfade = Split(TextDateiLesen(blConst.PfadZuProgPfadeText, False), vbCrLf)
        IconsEinfuegen() ' Für die Programmstarter-Buttons

        kwNewsZahlGeaendert(NewsZaehlenKW(AktHeft, aktKWMo, aktKWSo))

        ' Tag, Datum
        Dim tag As String = WochentagDeutsch(Today.DayOfWeek)
        lbl_heute.Content = tag & ", " & Today.ToShortDateString

        WerRuftAnProc = Process.Start(blConst.PfadZuWERruftAN)
    End Sub

    Private Sub ProgrammEnde()

        TextBoxInhalteSichern()

        'Code des aktuellen Heftes für den nächsten Start speichern
        Dim b As Boolean = TextDateiSchreiben(blConst.STANDARDPFAD + blConst.AKTAUFTRAGDATEI + blConst.DATEIENDUNG, AktAuftrag.Code, True)

        ' ZIP-Backup anlegen, ein jüngeres Backup vom selben Tag wird dabei überschrieben
        Try
            Dim ZipDateiName As String = "blAuftrag2014_" + Today.Year.ToString + "-" + NullVorWert(Today.Month.ToString, 2) + "-" + NullVorWert(Today.Day.ToString, 2) + "-" + NullVorWert(Now.Hour.ToString, 2) + ".zip"
            Dim tf As Boolean = ZippeOrdnerSchnell(blConst.STANDARDPFAD, blConst.ZIPpfad + ZipDateiName, True)
        Catch ex As Exception
            MsgBox("Fehler beim Erstellen des Backups: " + vbCrLf + ex.ToString)
        End Try

        If Not WerRuftAnProc.HasExited Then WerRuftAnProc.Kill()
    End Sub

    Private Sub TextBoxInhalteSichern()
        Dim b As Boolean
        Dim no As String() = TBinhalteBeiProgStart
        b = TextPruefenUndSpeichern(no(1), blConst.STANDARDPFAD + "Liste1" + blConst.DATEIENDUNG,
                               tb_Liste1.Text, True)
        b = TextPruefenUndSpeichern(no(2), blConst.STANDARDPFAD + "Liste2" + blConst.DATEIENDUNG,
                               tb_Liste2.Text, True)
        b = TextPruefenUndSpeichern(no(3), blConst.STANDARDPFAD + "Liste3" + blConst.DATEIENDUNG,
                               tb_Liste3.Text, True)
        b = TextPruefenUndSpeichern(no(4), blConst.STANDARDPFAD + "Liste4" + blConst.DATEIENDUNG,
                               tb_Liste4.Text, True)
        b = TextPruefenUndSpeichern(no(5), blConst.STANDARDPFAD + "Liste5" + blConst.DATEIENDUNG,
                               tb_Daten1.Text, True)
        b = TextPruefenUndSpeichern(no(6), blConst.STANDARDPFAD + "Liste6" + blConst.DATEIENDUNG,
                               tb_Liste6.Text, True)
        b = TextPruefenUndSpeichern(no(7), blConst.STANDARDPFAD + "Termine1" + blConst.DATEIENDUNG,
                               tb_Termine1.Text, True)
        b = TextPruefenUndSpeichern(no(8), blConst.STANDARDPFAD + "Termine2" + blConst.DATEIENDUNG,
                               tb_Termine2.Text, True)
    End Sub

    Private Function ListeDerAuftraegeInListeEinlesen(ByVal Heft As String, Ausgabe As String, ByRef Liste As List(Of String)) As List(Of String)
        Dim pfad As String = kompletterPfadZumHeft(Heft, Ausgabe)
        Dim code As String = ""
        For Each z In qListeDerDateiNamenImOrdner(pfad, blConst.DATEIENDUNG)
            code = Replace(Mid(z, z.LastIndexOf("\") + 2), blConst.DATEIENDUNG, "")
            Liste.Add(code)
        Next
        Return Liste
    End Function

    Private Sub ListeDerAuftraegeInListBoxEinlesen()
        Dim pfad As String = ""
        pfad = kompletterPfadZumAktuellenHeft()

        lb_auftragsliste.Items.Clear()
        Dim code As String = ""
        For Each z In qListeDerDateiNamenImOrdner(pfad, blConst.DATEIENDUNG)
            code = Replace(Mid(z, z.LastIndexOf("\") + 2), blConst.DATEIENDUNG, "")
            lb_auftragsliste.Items.Add(code)
        Next
        If lb_auftragsliste.Items.Count > 0 Then lb_auftragsliste.SelectedIndex = 0
    End Sub

    Private Sub SucheNachAuftraegen(ByVal s As String)
        ' Extra: Suche nach News:Thema
        ' Soll alle Aufträge finden, in dessen Titel das Wort News vorkommt, dann aber 
        ' nur diejenigen anzeigen, in denen das Thema vorkommt. z.B. "News:Roboter"
        Dim newsSuche As Boolean = False
        Dim newsThema As String = ""
        If s.StartsWith("News:") Then
            newsSuche = True
            newsThema = Replace(s, "News:", "").ohneFührendeLeerzeichen
            s = "News"
        ElseIf s.StartsWith("news:")
            newsSuche = True
            newsThema = Replace(s, "news:", "").ohneFührendeLeerzeichen
            s = "news"
        End If

        ' Treffer werden in die Liste der Aufträge eingelesen (lb_auftragsliste)
        Dim pfad As String = ""
        Dim AlleHefte As New List(Of String)
        AlleHefte.Add("dnp")
        AlleHefte.Add("wum")
        AlleHefte.Add("com")
        Dim AlleAusgaben As New List(Of String)
        Dim code As String = ""

        lb_auftragsliste.Items.Clear()
        For Each h As String In AlleHefte
            pfad = blConst.STANDARDPFAD + h + "\"
            ' Neue Version mit QuickIO
            For Each z In qListeAllerDateienImOrdnerUndDessenUnterordnern(pfad, blConst.DATEIENDUNG)
                If InStr(z.ToUpper, s.ToUpper) > 0 Then
                    code = Replace(Mid(z, z.LastIndexOf("\") + 2), blConst.DATEIENDUNG, "")
                    lb_auftragsliste.Items.Add(code)
                End If
            Next
        Next

        If newsSuche Then
            Dim newsTrefferListe As New List(Of String)
            ' MsgBox("Anzahl News-Dateien: " & lb_auftragsliste.Items.Count.ToString)

            ' Alle Dateien öffnen und prüfen, ob das newsThema drin vorkommt
            For Each code In lb_auftragsliste.Items
                Dim AuftragsText As String = AuftragsDateiEinLesen(code)
                If InStr(AuftragsText.ToLower, newsThema.ToLower) > 0 Then
                    newsTrefferListe.Add(code)
                End If
            Next
            lb_auftragsliste.Items.Clear()
            For Each code In newsTrefferListe
                lb_auftragsliste.Items.Add(code)
            Next

            ' Die gefundenen Zeilen noch auslesen und die Ergebnisse anzeigen
            Dim x As Integer = 0
            Dim sb As New System.Text.StringBuilder
            For Each code In newsTrefferListe
                'sb.Append(code + vbCrLf)
                Dim AuftragsText As String = AuftragsDateiEinLesen(code)
                Dim heftAusg As String = Mid(code, 1, 7)
                For Each row In Split(AuftragsText, vbCrLf)
                    If InStr(row.ToLower, newsThema.ToLower) > 0 Then
                        x += 1
                        Dim textZeichen As Integer = InStr(row, blConst.CSVTRENNER) - 1
                        Dim text As String = Mid(row, 1, textZeichen)
                        Dim datum As String = Mid(row, textZeichen + 2, 10)
                        sb.Append(heftAusg & ": " & datum & " " & text & vbCrLf)
                    End If
                Next
            Next
            MsgBox(sb.ToString, Title:="Die Suche nach '" & newsThema & "' ergab " & x.ToString & " Treffer.")
        End If
    End Sub

    Private Sub AuftragAnzeigen(ByVal code As String)
        'MsgBox(code)
        Dim a As New Auftrag
        a = AuftragEinLesen(code)
        If a Is Nothing Then Exit Sub

        'Dies ist die einzige Stelle an der der AktAuftrag gesetzt und sofort angezeigt wird!
        AktAuftrag = a
        lbl_AuftragsString.Content = a.toString

        If a.Art = Auftrag.Typ.std Then
            Zeiterfassung.Visibility = Windows.Visibility.Visible
            Stückerfassung.Visibility = Windows.Visibility.Collapsed

            lb_AEzeilenZeit.Items.Clear()
            For Each x As AE In a.AEs
                lb_AEzeilenZeit.Items.Add(x.ToStringZeitFixLength)
            Next
        Else
            ' Stück oder Pauschal
            Zeiterfassung.Visibility = Windows.Visibility.Collapsed
            Stückerfassung.Visibility = Windows.Visibility.Visible

            lb_AEzeilenST.Items.Clear()
            For Each e As AE In a.AEs
                lb_AEzeilenST.Items.Add(e.ToStringStueckFixLength)
            Next
        End If
    End Sub

    Private Sub HeftOderAusgabeGeaendert()
        AktAusgabe = CStr(lbl_HeftNrAktuell.Content)
        ListeDerAuftraegeInListBoxEinlesen()
        If AktHeft <> Auftrag.Heft.com Then
            kwNewsZahlGeaendert(NewsZaehlenKW(AktHeft, aktKWMo, aktKWSo))
        End If
    End Sub

    Private Sub StueckAEInMaskeUebernehmen(ByVal AEidx As Integer)
        AEzeileStueckStatus_NEU = False : st_lblmodus.Content = "Ändern" : st_lblmodus.Foreground = Brushes.DarkRed
        Dim ae As AE = AktAuftrag.AEs(AEidx)
        st_tbtext.Text = ae.Text
        st_tbdatum.Text = ae.Datum.ToShortDateString
        st_tbwert.Text = ae.Wert.ToString
        ' Felder für Änderungen freigeben
        st_tbtext.IsEnabled = True
        st_tbdatum.IsEnabled = True
    End Sub

    Private Sub ZeitAEinMaskeUebernehmen(ByVal AEidx As Integer)
        AEzeileZeitStatus_NEU = False : ez_lblmodus.Content = "Ändern" : ez_lblmodus.Foreground = Brushes.DarkRed
        Dim ae As AE = AktAuftrag.AEs(AEidx)
        ez_tbtext.Text = ae.Text
        ez_tbdatum.Text = ae.Datum.ToShortDateString
        ez_tbvon.Text = ae.von.ToShortTimeString
        ez_tbbis.Text = ae.bis.ToShortTimeString
        ' Felder für Änderungen freigeben
        ez_tbtext.IsEnabled = True
        ez_tbdatum.IsEnabled = True
    End Sub

    Private Sub ZeitAEMaskeLeeren()
        ez_tbtext.Text = "" : ez_tbdatum.Text = "" : ez_tbvon.Text = "" : ez_tbbis.Text = "" : ez_lbldauer.Content = ""
        AEzeileZeitStatus_NEU = True : ez_lblmodus.Content = "Neu" : ez_lblmodus.Foreground = Brushes.DarkGreen
    End Sub

    Private Sub StueckAEMaskeLeeren()
        st_tbtext.Text = "" : st_tbdatum.Text = "" : st_tbwert.Text = ""
        AEzeileStueckStatus_NEU = True : st_lblmodus.Content = "Neu" : st_lblmodus.Foreground = Brushes.DarkGreen
    End Sub

    Private Function testArbeitsmarkt(ByVal a As Auftrag) As Boolean
        Dim s As String = a.Text.ToLower
        Dim ret As Boolean = False
        'Abkürzung: 
        If a.Art <> Auftrag.Typ.pauschal Then Return ret

        If InStr(s, "arbeitsmarkt") > 0 Or InStr(s, "stellenmarkt") > 0 Then ret = True
        'Schreibfehler abfangen
        If ret = False Then
            'MsgBox(a.Text + " - " + a.Satz.ToString)
            If a.Satz = 225 And (InStr(s, "arbeit") > 0 Or InStr(s, "stellen") > 0 Or InStr(s, "arbiet") > 0) Then
                ret = True
            ElseIf a.Satz = 225 And (InStr(s, "markt") > 0 Or InStr(s, "martk") > 0) Then
                ret = True
            End If
        End If
        Return ret
    End Function

    Private Function testBuecher(ByVal a As Auftrag) As Boolean
        Dim s As String = a.Text.ToLower
        Dim ret As Boolean = False
        'Abkürzung
        If a.Art <> Auftrag.Typ.st Then Return ret

        If InStr(s, "bücher") > 0 Or InStr(s, "buecher") > 0 Then ret = True
        Return ret
    End Function

    Private Function testNews(ByVal a As Auftrag) As Boolean
        Dim s As String = a.Text.ToLower
        Dim ret As Boolean = False
        If a.Art = Auftrag.Typ.st And InStr(s, "news") > 0 Then ret = True
        Return ret
    End Function

    Private Sub RGVorbereiten()
        Dim RedStd As Double = 0
        Dim RedWert As Double = 0
        Dim BuecherSeiten As Double = 0
        Dim BuecherWert As Double = 0
        Dim Stellenmarkt As Double = 0
        Dim NewsStueck As Integer = 0
        Dim NewsWert As Double = 0
        Dim newsliste As New System.Text.StringBuilder
        Dim sonstText As String = ""
        Dim sonstStueck As Double = 0
        Dim sonstWert As Double = 0
        Dim Summe As Double = 0
        Dim MwSt As Double = 0
        Dim Brutto As Double = 0

        'RG wird für die gerade angezeigten Aufträge gemacht
        For Each code As String In lb_auftragsliste.Items
            'Auftrag einlesen
            Dim a As New Auftrag
            a = AuftragEinLesen(code)
            If a Is Nothing Then Exit Sub

            If testBuecher(a) Then
                BuecherWert = a.AktWert
                BuecherSeiten = BuecherWert / a.Satz

            ElseIf testArbeitsmarkt(a) Then
                Stellenmarkt = a.AktWert

            ElseIf a.Art = Auftrag.Typ.std Then
                RedStd = RedStd + a.Bisher
                RedWert = RedStd * a.Satz

            ElseIf testNews(a) Then
                NewsStueck = a.AEs.Count
                NewsWert = a.AktWert
                For Each n As AE In a.AEs
                    newsliste.Append(n.Datum.ToShortDateString + "  " + n.Text + vbCrLf)
                Next
            Else
                ' Hier bleiben die selbstgeschriebenen Seiten über
                sonstStueck += Math.Round(a.AktWert / a.Satz, 2)
                sonstWert = a.AktWert
                sonstText += a.Text + " "
            End If
        Next
        Summe = Summe + RedWert + BuecherWert + Stellenmarkt + NewsWert + sonstWert
        MwSt = Summe * 0.07
        Brutto = Summe + MwSt

        ' Tabelle definieren
        Dim t As New myTable("14 fill .",
                             "16 center Fill .",
                             "1 fill .",
                             "14 euro fill .")
        With t
            Dim ABCDLinie As String = "-".Mal(.ABCD.width)
            .ABCD.align = myColumn.alignTyp.center
            .ABCD.fill = " "

            'Ergebnisse zusammenschreiben
            Dim tmp As String = AktHeft.ToString + " " + AktAusgabe + "  / Bernhard Lauer / " + Today.ToShortDateString
            .newLine(tmp.f(.ABCD))
            .newLine("-".Mal(.ABCD.width).f(.ABCD))

            If RedWert > 0 Then
                tmp = NumFormat(RedStd, 2, 2) + " Stunden "
                .newLine("Redaktion".f(.A) + tmp.f(.B) + "".f(.C) + RedWert.f(.D))
            End If

            If BuecherWert > 0 Then
                tmp = NumFormat(BuecherSeiten, 2, 2) + " Seiten "
                .newLine("Bücher".f(.A) + tmp.f(.B) + "".f(.C) + BuecherWert.f(.D))
            End If

            If Stellenmarkt > 0 Then
                .newLine("Arbeitsmarkt".f(.A) + " Pauschal ".f(.B) + "".f(.C) + Stellenmarkt.f(.D))
            End If

            If NewsWert > 0 Then
                tmp = NumFormat(NewsStueck, 0, 3) + " Stück "
                .newLine("Online-News".f(.A) + tmp.f(.B) + "".f(.C) + NewsWert.f(.D))
            End If

            If sonstWert > 0 Then
                tmp = NumFormat(sonstStueck, 2, 2) + " Seiten "
                .newLine(sonstText.f(.A) + tmp.f(.B) + "".f(.C) + sonstWert.f(.D))
            End If

            t.ABC.fill = " "
            t.D.fill = " "

            .newLine(ABCDLinie)
            .newLine("Nettosumme".f(.ABC) + Summe.f(.D))
            .newLine("7% MwSt".f(.ABC) + MwSt.f(.D))
            .newLine(ABCDLinie)
            .newLine("Bruttosumme".f(.ABC) + Brutto.f(.D))
            .newLine("=".Mal(.ABCD.width).f(.ABCD))

            .newLine("")
            .newLine(ABCDLinie)
            .newLine(("Liste der Online-News (" + NewsStueck.ToString + " Stück)").f(.ABCD))
            .newLine(ABCDLinie)
            .newLine(newsliste.ToString)

            Clipboard.SetText(.toString)
            tb_RGtexte.Text = .toString
        End With

        ti_Rechnung.Focus()
        bt_RGfürListe.Content = AktHeft.ToString + " " + AktAusgabe + " |   " + StellenHintermKomma(Summe, 2) + " | "
    End Sub

    Private Function TagStatistikPflege(ByVal DateiInhalt As String) As String
        ' Sorgt dafür, dass auch Tage eingefügt werden, an denen nichts gearbeitet wurde
        Dim datum As Date = Today
        Dim key As String
        Dim neueZeile As String

        For i = -31 To 0
            datum = Today.AddDays(i)
            key = datum.Year.ToString + NullVorWert(datum.Month.ToString, 2) + NullVorWert(datum.Day.ToString, 2) + ":"
            If CBool(InStr(DateiInhalt, key)) Then
                'MsgBox("ok: " + key)
                ' Alles OK.
            Else
                ' Kein Eintrag für diesen Tag, also oben anfügen
                neueZeile = key + "0" + vbCrLf
                DateiInhalt = neueZeile + DateiInhalt
            End If
        Next

        Return DateiInhalt
        MsgBox(DateiInhalt)
    End Function

    Private Function kwStatistikPflege(ByVal DateiInhalt As String) As String
        ' Die funktion trägt fehlende Wochen in der KW-Datai nach (mit Eintrag 0, z.B. 2015_12:0)

        Dim kw As Integer = Kalenderwoche(Today)
        Dim key As String
        Dim neueZeile As String

        ' Funktioniert nicht für die ersten vier Kalenderwochen, macht aber auch keinen Fehler
        For i = 4 To 0
            kw = kw - i : If kw < 1 Then Exit For

            key = Today.Year.ToString + "_" + NullVorWert(kw.ToString, 2) + ":"

            If CBool(InStr(DateiInhalt, key)) Then
                ' Alles OK.
            Else
                ' Kein Eintrag für diese KW, also oben anfügen
                neueZeile = key + "0" + vbCrLf
                DateiInhalt = neueZeile + DateiInhalt
            End If
        Next
        Return DateiInhalt
    End Function

    Private Function MonatsStatistikPflege(ByVal DateiInhalt As String) As String
        Dim key As String
        Dim neueZeile As String
        For i = -1 To 0
            If Today.Month > 1 Then
                key = Today.Year.ToString + "_" + NullVorWert(Today.AddMonths(i).Month.ToString, 2) + ":"
            Else
                key = (Today.Year - 1).ToString + "_" + NullVorWert(Today.AddMonths(i).Month.ToString, 2) + ":"
            End If
            If CBool(InStr(DateiInhalt, key)) Then
                ' Alles OK.
            Else
                ' Kein Eintrag für diesen Monat, also oben anfügen
                neueZeile = key + "0" + vbCrLf
                DateiInhalt = neueZeile + DateiInhalt
            End If
        Next
        Return DateiInhalt
    End Function

    Private Sub TagStatistikAendern(ByVal delta As Double, ByVal datum As Date)
        Dim TagesUmsaetze As String = TagStatistikPflege(TextDateiLesen(blConst.TAGESDATEIPFAD, False))
        '  MsgBox(TagesUmsaetze)

        Dim key As String = datum.Year.ToString + NullVorWert(datum.Month.ToString, 2) + NullVorWert(datum.Day.ToString, 2) + ":"
        Dim TagWertAlt As Double = 0
        Dim TagWertNeu As Double = 0

        TagWertAlt = CDbl(WertAusStatistikDateiTextHolen(TagesUmsaetze, key))
        TagWertNeu = TagWertAlt + delta

        'MsgBox(datum.ToShortDateString + vbCrLf + "Bisher : " + TagWertAlt.ToString + " + " + delta.ToString + " = " + TagWertNeu.ToString)

        'Neuen Tageswert in lbl_Tag.Content eintragen
        If datum.ToShortDateString = Today.ToShortDateString Then
            lbl_tag.Content = Math.Round(TagWertNeu, 0).ToString
        End If

        'MsgBox(datum.ToShortDateString + ": " + TagWertAlt.ToString + " --> " + TagWertNeu.ToString)

        'TagDatei anpassen
        Dim neueZeile As String = key + StellenHintermKomma(TagWertNeu, 2) + vbCrLf
        'MsgBox(key + " --> " + neueZeile)

        TagesUmsaetze = zeileInStatistikDateiTextTauschen(TagesUmsaetze, key, neueZeile, True)
        ' MsgBox(TagesUmsaetze)
        TextDateiSchreiben(blConst.TAGESDATEIPFAD, TagesUmsaetze, False)

        ' Durchschnitt berechnen und eintragen
        Dim schnitt As Double = DurchschnittBerechnen(blConst.TAGESDATEIPFAD, 31, True)
        Dim abw As Double = Math.Round(TagWertNeu - schnitt, 0)

        lbl_tagSchnitt.Content = "(" + abw.ToString + " | " + schnitt.ToString + ")"

        ' Werte für die vorangegangenen Tage vorher eintragen
        Dim lblText As String = " | "
        For i = 1 To 5
            datum = datum.AddDays(-1)
            key = datum.Year.ToString + NullVorWert(datum.Month.ToString, 2) + NullVorWert(datum.Day.ToString, 2) + ":"
            lblText = lblText + OhneNachkommaStellen(WertAusStatistikDateiTextHolen(TagesUmsaetze, key)) + " | "
        Next
        lbl_Vortage.Content = lblText

        lbl_x_Tag.Content = Today.ToShortDateString
        lbl_x_TagIst.Content = Math.Round(TagWertNeu, 0).ToString
        lbl_x_TagSchnitt.Content = Math.Round(schnitt, 0).ToString
        lbl_x_TagAbw1.Content = Math.Round(abw, 0).ToString
        If abw < 0 Then lbl_x_TagAbw1.Foreground = Brushes.Red Else lbl_x_TagAbw1.Foreground = Brushes.Green

    End Sub

    Private Sub KWStatistikAendern(ByVal delta As Double, ByVal datum As Date)
        Dim Umsaetze As String = kwStatistikPflege(TextDateiLesen(blConst.KWDATEIPFAD, False))
        Dim kw As Integer = Kalenderwoche(datum)
        Dim kwString As String = kw.ToString
        Dim key As String = datum.Year.ToString + "_" + NullVorWert(kwString, 2) + ":"
        Dim WertAlt As Double = CDbl(WertAusStatistikDateiTextHolen(Umsaetze, key))
        Dim WertNeu As Double = WertAlt + delta

        ' Label Kalenderwoche neu befüllen
        If kw = Kalenderwoche(Today) Then lbl_kw.Content = Math.Round(WertNeu, 0)

        Dim neueZeile As String = key + StellenHintermKomma(WertNeu, 2) + vbCrLf

        'KW-Datei anpassen
        Umsaetze = zeileInStatistikDateiTextTauschen(Umsaetze, key, neueZeile, True)
        TextDateiSchreiben(blConst.KWDATEIPFAD, Umsaetze, False)

        ' Durchschnitt berechnen und eintragen
        Dim schnitt As Double = DurchschnittBerechnen(blConst.KWDATEIPFAD, 8, False)
        Dim abw As Double = Math.Round(WertNeu - schnitt, 0)
        lbl_kwSchnitt.Content = "(" + abw.ToString + " | " + schnitt.ToString + ")"

        ' Werte für die vorangegangenen Wochen eintragen
        Dim lblText As String = " | "
        For i = 1 To 4
            datum = datum.AddDays(-7)
            key = datum.Year.ToString + "_" + Kalenderwoche(datum).ToString + ":"
            lblText = lblText + OhneNachkommaStellen(WertAusStatistikDateiTextHolen(Umsaetze, key)) + " | "
        Next
        lbl_Vorwochen.Content = lblText

        lbl_x_Woche.Content = "KW " + Kalenderwoche(Today).ToString
        lbl_x_WocheIst.Content = Math.Round(WertNeu, 0).ToString
        lbl_x_WocheSchnitt.Content = Math.Round(schnitt, 0).ToString
        lbl_x_WocheAbw1.Content = Math.Round(abw, 0).ToString
        If abw < 0 Then lbl_x_WocheAbw1.Foreground = Brushes.Red Else lbl_x_WocheAbw1.Foreground = Brushes.Green

        Dim heut As Integer = CInt(Today.DayOfWeek)
        If heut = 0 Then heut = 5
        Dim hr As Double = 0
        hr = WertNeu / heut * 5
        lbl_x_WocheHR.Content = Math.Round(hr, 0).ToString

        ' Abw2 sagt, in welche Richtung sich der Schnitt ändert
        Dim abw2 As Double = hr - schnitt
        lbl_x_WocheAbw2.Content = Math.Round(abw2, 0).ToString
        If abw2 < 0 Then lbl_x_WocheAbw2.Foreground = Brushes.Red Else lbl_x_WocheAbw2.Foreground = Brushes.Green
        If hr < schnitt Then lbl_x_WocheHR.Foreground = Brushes.Red Else lbl_x_WocheHR.Foreground = Brushes.Green
    End Sub

    Private Sub MonatStatistikAendern(ByVal delta As Double, ByVal datum As Date)
        Dim Umsaetze As String = MonatsStatistikPflege(TextDateiLesen(blConst.MONATSDATEIPFAD, False))
        Dim key As String = datum.Year.ToString + "_" + NullVorWert(datum.Month.ToString, 2) + ":"
        Dim WertAlt As Double = CDbl(WertAusStatistikDateiTextHolen(Umsaetze, key))
        Dim WertNeu As Double = WertAlt + delta

        ' Label Monat neu befüllen
        If datum.Month = Today.Month Then lbl_monat.Content = Math.Round(WertNeu, 0)

        Dim neueZeile As String = key + StellenHintermKomma(WertNeu, 2) + vbCrLf

        'Monatsdatei anpassen
        Umsaetze = zeileInStatistikDateiTextTauschen(Umsaetze, key, neueZeile, True)
        TextDateiSchreiben(blConst.MONATSDATEIPFAD, Umsaetze, False)

        ' Durchschnitt berechnen und eintragen
        Dim schnitt As Double = DurchschnittBerechnen(blConst.MONATSDATEIPFAD, 4, False)
        Dim abw As Double = Math.Round(WertNeu - schnitt, 0)
        lbl_monatSchnitt.Content = "(" + abw.ToString + " | " + schnitt.ToString + ")"

        ' Werte für die vorangegangenen Monate eintragen
        Dim lblText As String = " | "
        For i = 1 To 3
            datum = datum.AddMonths(-1)
            key = datum.Year.ToString + "_" + NullVorWert(datum.Month.ToString, 2) + ":"
            lblText = lblText + OhneNachkommaStellen(WertAusStatistikDateiTextHolen(Umsaetze, key)) + " | "
        Next
        lbl_Vormonate.Content = lblText

        lbl_x_Monat.Content = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Today.Month)
        lbl_x_MonatIst.Content = Math.Round(WertNeu, 0).ToString
        lbl_x_MonatSchnitt.Content = Math.Round(schnitt, 0).ToString
        lbl_x_MonatAbw1.Content = Math.Round(abw, 0).ToString
        If abw < 0 Then lbl_x_MonatAbw1.Foreground = Brushes.Red Else lbl_x_MonatAbw1.Foreground = Brushes.Green

        Dim heut As Integer = Today.Day
        Dim ultimo As Date = CDate("1." + Today.Month.ToString + ". " + Today.Year.ToString).AddMonths(1).AddDays(-1)
        Dim anzTageImMonat As Integer = ultimo.Day
        Dim hr As Double = WertNeu / heut * anzTageImMonat
        lbl_x_MonatHR.Content = Math.Round(hr, 0).ToString
        Dim abw2 As Double = hr - schnitt
        lbl_x_MonatAbw2.Content = Math.Round(abw2, 0).ToString
        If abw2 < 0 Then lbl_x_MonatAbw2.Foreground = Brushes.Red Else lbl_x_MonatAbw2.Foreground = Brushes.Green
        If hr < schnitt Then lbl_x_MonatHR.Foreground = Brushes.Red Else lbl_x_MonatHR.Foreground = Brushes.Green

    End Sub

    Private Sub JahrStatistikAendern(ByVal delta As Double, ByVal datum As Date)
        Dim Umsaetze As String = TextDateiLesen(blConst.JAHRESDATEIPFAD, False)
        Dim key As String = datum.Year.ToString + ":"
        Dim WertAlt As Double = CDbl(WertAusStatistikDateiTextHolen(Umsaetze, key))
        Dim WertNeu As Double = WertAlt + delta

        ' Label Jahr neu befüllen
        If datum.Year = Today.Year Then lbl_jahr.Content = Math.Round(WertNeu, 0)

        Dim neueZeile As String = key + StellenHintermKomma(WertNeu, 2) + vbCrLf

        'Monatsdatei anpassen
        Umsaetze = zeileInStatistikDateiTextTauschen(Umsaetze, key, neueZeile, True)
        TextDateiSchreiben(blConst.JAHRESDATEIPFAD, Umsaetze, False)

        ' Durchschnitt berechnen und eintragen
        Dim schnitt As Double = DurchschnittBerechnen(blConst.JAHRESDATEIPFAD, 3, False)
        Dim abw As Double = Math.Round(WertNeu - schnitt, 0)
        lbl_jahrSchnitt.Content = "(" + abw.ToString + " | " + schnitt.ToString + ")"

        ' Werte für die vorangegangenen Jahre eintragen
        Dim lblText As String = " | "
        For i = 1 To 3
            datum = datum.AddYears(-1)
            key = datum.Year.ToString + ":"
            lblText = lblText + OhneNachkommaStellen(WertAusStatistikDateiTextHolen(Umsaetze, key)) + " | "
        Next
        lbl_Vorjahre.Content = lblText

        lbl_x_Jahr.Content = Today.Year.ToString
        lbl_x_JahrIst.Content = Math.Round(WertNeu, 0).ToString
        lbl_x_JahrSchnitt.Content = Math.Round(schnitt, 0).ToString
        lbl_x_JahrAbw1.Content = Math.Round(abw, 0).ToString
        If abw < 0 Then lbl_x_JahrAbw1.Foreground = Brushes.Red Else lbl_x_JahrAbw1.Foreground = Brushes.Green

        Dim ts As TimeSpan
        ts = (Today - CDate("1.1." + Today.Year.ToString))
        Dim anzTageBisher As Integer = ts.Days

        Dim hr As Double = WertNeu / anzTageBisher * 365
        lbl_x_JahrHR.Content = Math.Round(hr, 0).ToString
        Dim abw2 As Double = hr - schnitt
        lbl_x_JahrAbw2.Content = Math.Round(abw2, 0).ToString
        If abw2 < 0 Then lbl_x_JahrAbw2.Foreground = Brushes.Red Else lbl_x_JahrAbw2.Foreground = Brushes.Green
        If hr < schnitt Then lbl_x_JahrHR.Foreground = Brushes.Red Else lbl_x_JahrHR.Foreground = Brushes.Green

    End Sub

    Private Sub AktAuftrag_ValueChanged(ByVal a As Auftrag, ByVal delta As Double, datum As Date) Handles AktAuftrag.ValueChanged
        'MsgBox(a.Code + ": " + delta.ToString + " (" + datum.ToShortDateString + ")")
        TagStatistikAendern(delta, datum)

        'KW korrigieren / hinzufügen
        KWStatistikAendern(delta, datum)

        'Monat korrigieren / hinzufügen
        MonatStatistikAendern(delta, datum)

        'Jahr korrigieren / hinzufügen
        JahrStatistikAendern(delta, datum)

    End Sub

#Region "PDFLaufzettel"
    Function ArbeitskopieAnlegen(ByVal DateiNameMitPfad As String) As String
        Kopie = DateiNameMitPfad
        If Kopie = "" Then
            MsgBox("Fehler: Kein Pfad angegeben")
            Return "Fehler: Kein Pfad angegeben"
        Else
            System.IO.File.Copy(OriginalPDF, Kopie)
        End If

        Return Kopie
    End Function

    Private Function MonatZurNummer(ByVal Nr As Integer) As String
        Dim m As String = ""
        If Nr = 1 Then m = "Jan"
        If Nr = 2 Then m = "Feb"
        If Nr = 3 Then m = "März"
        If Nr = 4 Then m = "April"
        If Nr = 5 Then m = "Mai"
        If Nr = 6 Then m = "Juni"
        If Nr = 7 Then m = "Juli"
        If Nr = 8 Then m = "Aug"
        If Nr = 9 Then m = "Sept"
        If Nr = 10 Then m = "Okt"
        If Nr = 11 Then m = "Nov"
        If Nr = 12 Then m = "Dez"
        Return m
    End Function

    Private Sub PDFVorgabenEintragen()
        Dim ausg As Integer = Now.Month + 2
        Dim jahr As Integer = Now.Year
        If ausg > 12 Then
            ausg = ausg - 12
            jahr = jahr + 1
        End If
        tbAusgabe.Text = ausg.ToString + "-" + jahr.ToString
        tbDatum.Text = Now.ToShortDateString
        tbSeite.Text = " (Plan: )"

        Dim j As String = jahr.ToString
        j = Mid(j, Len(j) - 1)

        tb_zieldatei.Text = "dnp_" + j + NullVorWert(ausg.ToString, 2) + "_Laufzettel_XXXX.PDF"

        PfadFuerDenLaufzettelSetzen(MonatZurNummer(Now.Month))

        tbRubrik.Focus()

        'Testdaten
        'tbRubrik.Text = "Testrubrik"
        'tbThema.Text = "Thema"
        'tbSeite.Text = "123"
        'tbA1.Text = "Eine Anmerkung und "
        'tbA2.Text = "die zweite Zeile der Anmerkung."

    End Sub

    Private Sub DatenInPDFeintragen(ByVal PDFVorlage As String, ByVal ZielPDF As String)
        ' To Do: Übergebene Pfade prüfen!
        Dim pdfReader As New pdfReader(PDFVorlage)
        Dim pdfStamper As New pdfStamper(pdfReader, New System.IO.FileStream(
            ZielPDF, System.IO.FileMode.Create), "\0", True)

        Dim pdfFelder As AcroFields = pdfStamper.AcroFields
        pdfFelder.SetField("Ausgabe", tbAusgabe.Text)
        pdfFelder.SetField("Rubrik", tbRubrik.Text)
        pdfFelder.SetField("Thema", tbThema.Text)
        pdfFelder.SetField("Seite", tbSeite.Text)
        pdfFelder.SetField("redakteur", tbRedakteur.Text)
        pdfFelder.SetField("Datum", tbDatum.Text)
        pdfFelder.SetField("Anmerkungen 1", tbA1.Text)
        pdfFelder.SetField("Anmerkungen 2", tbA2.Text)
        pdfFelder.SetField("Anmerkungen 3", tbA3.Text)
        pdfFelder.SetField("Anmerkungen 4", tbA4.Text)

        pdfStamper.FormFlattening = False
        pdfStamper.Close()
        pdfReader.Close()

        MsgBox("Die Daten wurden in " & ZielPDF & " eingetragen.")
    End Sub

    Private Sub DatenInsPDFeintragen()
        Dim pdfReader As New pdfReader(Kopie)
        Dim newFile As String = tb_zielpfad.Text + tb_zieldatei.Text
        If System.IO.File.Exists(newFile) Then
            MsgBox("Zieldatei " + newFile + " existiert bereits!")
            Exit Sub
        End If

        Dim pdfStamper As New pdfStamper(pdfReader, New System.IO.FileStream(
            newFile, System.IO.FileMode.Create), "\0", True)

        Dim pdfFormFields As AcroFields = pdfStamper.AcroFields
        pdfFormFields.SetField("Ausgabe", tbAusgabe.Text)
        pdfFormFields.SetField("Rubrik", tbRubrik.Text)
        pdfFormFields.SetField("Thema", tbThema.Text)
        pdfFormFields.SetField("Seite", tbSeite.Text)
        pdfFormFields.SetField("redakteur", tbRedakteur.Text)
        pdfFormFields.SetField("Datum", tbDatum.Text)
        pdfFormFields.SetField("Anmerkungen 1", tbA1.Text)
        pdfFormFields.SetField("Anmerkungen 2", tbA2.Text)
        pdfFormFields.SetField("Anmerkungen 3", tbA3.Text)
        pdfFormFields.SetField("Anmerkungen 4", tbA4.Text)

        pdfStamper.FormFlattening = False
        pdfStamper.Close()
        pdfReader.Close()

        MsgBox("Die Daten wurden in " & newFile & " eingetragen.")
    End Sub

    Private Sub HeftNrFuerDateiname(ByVal JJAA As String)
        ' dnp_JJAA_Laufzettel_XXXX.pdf
        Dim dn As String = tb_zieldatei.Text
        tb_zieldatei.Text = Replace(dn, Mid(dn, 5, 4), JJAA)
    End Sub

    Private Sub ThemaFuerDateinameUndPfad()
        If tbThema.Text <> "" Then
            tb_zieldatei.Text = Replace(tb_zieldatei.Text, "XXXX", tbThema.Text)
            If InStr(tb_zielpfad.Text, tbThema.Text) < 1 Then tb_zielpfad.Text += tbThema.Text
            prüfePfad(tb_zielpfad)
        End If
    End Sub

    Private Sub PfadFuerDenLaufzettelSetzen(ByVal Monat As String)
        Dim t As String = tbAusgabe.Text
        If t <> "" Then
            Dim ausg As String = Replace(Mid(t, 1, 2), "-", "")
            Dim jahr As String = Mid(t, Len(t) - 3)
            Dim pf As String = jahr + "-" + NullVorWert(ausg, 2) + "-" + Monat
            tb_zielpfad.Text = Arbeitspfad + pf + "\"
            prüfePfad(tb_zielpfad)
        End If
    End Sub

    Private Function prüfePfad(ByRef tb As TextBox) As Boolean
        Dim ok As Boolean = False
        If Mid(tb.Text, tb.Text.Length, 1) <> "\" Then tb.Text += "\"
        If System.IO.Directory.Exists(tb.Text) Then
            tb.Background = Brushes.LightGreen
            ok = True
        Else
            tb.Background = Brushes.MistyRose
        End If
        Return ok
    End Function

    Private Function FormularIstVollstaendig() As Boolean
        Dim ok As Boolean = True
        If tbAusgabe.Text = "" Then ok = False
        If tbRubrik.Text = "" Then ok = False
        If tbThema.Text = "" Then ok = False Else ThemaFuerDateinameUndPfad()
        If tbSeite.Text = "" Then ok = False
        If tbRedakteur.Text = "" Then ok = False
        If tbDatum.Text = "" Then ok = False

        Dim DateiName As String = tb_zielpfad.Text + tb_zieldatei.Text
        If System.IO.File.Exists(DateiName) Then ok = False

        Return ok
    End Function

#End Region

    ' Events der Benutzeroberfläche
    Private Sub lbl_wum_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles lbl_wum.MouseDown
        wumLabelHervorheben()
        AktHeft = Auftrag.Heft.wum
        HeftOderAusgabeGeaendert()
    End Sub

    Private Sub lbl_dnp_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles lbl_dnp.MouseDown
        dnpLabelHervorheben()
        AktHeft = Auftrag.Heft.dnp
        HeftOderAusgabeGeaendert()
    End Sub

    Private Sub lbl_com_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles lbl_com.MouseDown
        comLabelHervorheben()
        AktHeft = Auftrag.Heft.com
        HeftOderAusgabeGeaendert()
    End Sub

    Private Sub lbl_HeftNrminus1_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles lbl_HeftNrminus1.MouseDown
        lbl_HeftNrAktuell.Content = lbl_HeftNrminus1.Content
        lbl_HeftNrminus1.Content = HeftNrMinus1(CStr(lbl_HeftNrAktuell.Content))
        lbl_HeftNrplus1.Content = HeftNrPlus1(CStr(lbl_HeftNrAktuell.Content))
        HeftOderAusgabeGeaendert()
    End Sub

    Private Sub lbl_HeftNrplus1_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles lbl_HeftNrplus1.MouseDown
        lbl_HeftNrAktuell.Content = lbl_HeftNrplus1.Content
        lbl_HeftNrminus1.Content = HeftNrMinus1(CStr(lbl_HeftNrAktuell.Content))
        lbl_HeftNrplus1.Content = HeftNrPlus1(CStr(lbl_HeftNrAktuell.Content))
        HeftOderAusgabeGeaendert()
    End Sub

    Private Sub bt_NeuerAuftrag_Click(sender As Object, e As RoutedEventArgs) Handles bt_NeuerAuftrag.Click
        'Eingabe wieder abschalten
        If stp_NeuerAuftrag.Visibility = Windows.Visibility.Visible Then
            stp_Listbox_etc.Margin = New Thickness(0, 0, 30, 10)
            stp_NeuerAuftrag.Visibility = Windows.Visibility.Collapsed

            Exit Sub
        End If

        stp_Listbox_etc.Margin = New Thickness(0, -94, 0, 0)

        tb_heft.Text = AktHeft.ToString
        tb_ausgabe.Text = AktAusgabe
        tb_typ.Text = Auftrag.Typ.std.ToString
        tb_je.Text = "35"
        stp_NeuerAuftrag.Visibility = Windows.Visibility.Visible
        tb_text.Text = ""
        tb_text.Focus()
    End Sub

    Private Sub bt_NeuenAuftragSichern_Click(sender As Object, e As RoutedEventArgs) Handles bt_NeuenAuftragSichern.Click

        Dim typfehler As Boolean = False
        Dim t As String = tb_typ.Text
        If t <> "std" And t <> "st" And t <> "pauschal" Then typfehler = True
        If typfehler Or tb_text.Text = "" Or tb_je.Text = "" Then
            MsgBox("Felder dürfen nicht leer bleiben oder fehlerhafter Typ")
            Exit Sub
        End If

        ' Bücherfehler abfangen: Bücher müssen vom Typ st sein.
        If InStr(tb_text.Text.ToLower, "bücher") > 0 Or InStr(tb_text.Text.ToLower, "buecher") > 0 Then
            If tb_typ.Text <> "st" Then
                tb_typ.Text = "st"
                tb_je.Text = "180"
                MsgBox("Bücher müssen vom Typ >st< sein!")
                Exit Sub  ' Es muss eine Möglichkeit zum Ändern der Eingaben geben
            End If
        End If

        t = tb_je.Text
        If t <> "35" And t <> "15" And t <> "225" Then
            Dim yn As MsgBoxResult = MsgBox("Ungewöhnlicher Betrag (" + t + ") übernehmen?", vbYesNo)
            If yn = vbNo Then Exit Sub
        End If

        'Fehlerbehandlung steht in den Eingabeevents
        Dim h As Auftrag.Heft = CType([Enum].Parse(GetType(Auftrag.Heft), tb_heft.Text), Auftrag.Heft)
        Dim typ As Auftrag.Typ = CType([Enum].Parse(GetType(Auftrag.Typ), tb_typ.Text), Auftrag.Typ)

        Dim a As New Auftrag(h, tb_ausgabe.Text, tb_text.Text, typ, CDbl(tb_je.Text))
        If Not neuenAuftragSichern(a) Then MsgBox("Neuer Auftrag " + a.Code + " konnte nicht gesichert werden!")

        stp_Listbox_etc.Margin = New Thickness(0, 0, 30, 10)
        stp_NeuerAuftrag.Visibility = Windows.Visibility.Collapsed

        'Liste der Aufträge neu einlesen, schließlich ist einer hinzugekommen
        ListeDerAuftraegeInListBoxEinlesen()
    End Sub

    Private Sub tb_heft_LostMouseCapture(sender As Object, e As MouseEventArgs) Handles tb_heft.LostMouseCapture
        If tb_heft.Text <> "dnp" Then
            If tb_heft.Text <> "wum" Then
                MsgBox("Bitte dnp oder wum eingeben!")
                tb_heft.Text = "dnp"
                tb_heft.Focus()
            End If
        End If
    End Sub

    Private Sub tb_ausgabe_LostFocus(sender As Object, e As RoutedEventArgs) Handles tb_ausgabe.LostFocus
        Dim fehler As Boolean = False
        If tb_ausgabe.Text.Length <> 4 Then fehler = True
        Try
            Dim i As Integer = CInt(tb_ausgabe.Text)
        Catch ex As Exception
            fehler = True
        End Try
        If fehler Then
            MsgBox("Bitte die Ausgabe als JJMM eingeben, z.B. 1405")
            tb_ausgabe.Text = AktAusgabe
            tb_ausgabe.Focus()
        End If
    End Sub

    Private Sub tb_typ_LostFocus(sender As Object, e As RoutedEventArgs) Handles tb_typ.LostFocus
        Dim t As String = tb_typ.Text
        Try
            [Enum].Parse(GetType(Auftrag.Typ), t)
            If t = "std" Then tb_je.Text = "35"
            If t = "st" Then tb_je.Text = "15"
            If t = "pauschal" Then tb_je.Text = "225"

        Catch ex As Exception
            MsgBox("Gültige Eingaben sind: std, st, pauschal")
            tb_typ.Text = "std"
            tb_typ.Focus()
        End Try
    End Sub

    Private Sub tb_text_TextChanged(sender As Object, e As TextChangedEventArgs) Handles tb_text.TextChanged
        If bt_NeuenAuftragSichern Is Nothing Or tb_text.Text = "" Then Exit Sub
        bt_NeuenAuftragSichern.Content = "Auftrag " + tb_heft.Text + tb_ausgabe.Text + tb_text.Text + " sichern"
        If tb_text.Text.ToLower = "bücher" Or tb_text.Text.ToLower = "buecher" Then
            tb_typ.Text = "st"
            tb_je.Text = "180"
        ElseIf InStr(tb_text.Text.ToLower, "news") > 0 Then
            tb_typ.Text = "st"
            tb_je.Text = "15"
        End If
    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles Me.Closing
        ProgrammEnde()
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        ProgrammStart()
    End Sub

    Private Sub lb_auftragsliste_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lb_auftragsliste.SelectionChanged
        If lb_auftragsliste.SelectedIndex > -1 Then AuftragAnzeigen(lb_auftragsliste.SelectedItem.ToString)
    End Sub

    Private Sub lb_AEzeilenST_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles lb_AEzeilenST.KeyUp

        If e.Key = Key.Delete Then
            If lb_AEzeilenST.SelectedIndex < 0 Then Exit Sub
            Dim idx As Integer = lb_AEzeilenST.SelectedIndex
            Dim itm As String = lb_AEzeilenST.SelectedItem.ToString

            'Sicherheitsabfrage
            If MsgBox("Wollen Sie die Zeile" + vbCrLf + "[" + itm + "]" + vbCrLf + "wirklich löschen?", CType(vbYesNo + MsgBoxStyle.Question, MsgBoxStyle)) = vbNo Then Exit Sub

            'AE in den Papierkorb sichern
            AEPapierkorb = AktAuftrag.AEs(idx)

            'Alles Ok. Also weg damit.
            AktAuftrag.DeleteAE(idx)

            'Auftrag speichern
            geaendertenAuftragSichern(AktAuftrag)

            'Neuanzeige, damit die Listbox aktualisiert wird
            AuftragAnzeigen(AktAuftrag.Code)

            ' NewsZahl anpassen
            kwNewsZahlGeaendert(NewsZaehlenKW(AktHeft, aktKWMo, aktKWSo))
        End If

        If e.Key = Key.Insert Then
            If AEPapierkorb.Text <> "" Then
                'Sicherheitsabfrage
                If MsgBox("Wollen Sie die Zeile" + vbCrLf + "[" + AEPapierkorb.ToString + "]" + vbCrLf + "wirklich einfügen?", CType(vbYesNo + MsgBoxStyle.Question, MsgBoxStyle)) = vbNo Then Exit Sub

                'Alles Ok. Also wieder anhängen.
                AktAuftrag.AppendAE(AEPapierkorb.Text, AEPapierkorb.Datum, Today, Today, AEPapierkorb.Wert)

                'Auftrag speichern
                geaendertenAuftragSichern(AktAuftrag)

                'Neuanzeige, damit die Listbox aktualisiert wird
                AuftragAnzeigen(AktAuftrag.Code)
            End If
        End If
    End Sub

    Private Sub lb_AEzeilenST_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lb_AEzeilenST.SelectionChanged
        lbl_idxStueck.Content = lb_AEzeilenST.SelectedIndex.ToString
        StueckAEMaskeLeeren()
        If lb_AEzeilenST.SelectedIndex < 0 Then Exit Sub
        StueckAEInMaskeUebernehmen(lb_AEzeilenST.SelectedIndex)
    End Sub

    Private Sub lb_AEzeilenZeit_KeyUp(sender As Object, e As KeyEventArgs) Handles lb_AEzeilenZeit.KeyUp
        If e.Key = Key.Delete Then
            If lb_AEzeilenZeit.SelectedIndex < 0 Then Exit Sub
            Dim idx As Integer = lb_AEzeilenZeit.SelectedIndex
            Dim itm As String = lb_AEzeilenZeit.SelectedItem.ToString

            'Sicherheitsabfrage
            If MsgBox("Wollen Sie die Zeile" + vbCrLf + "[" + itm + "]" + vbCrLf + "wirklich löschen?", CType(vbYesNo + MsgBoxStyle.Question, MsgBoxStyle)) = vbNo Then Exit Sub

            'AE in den Papierkorb sichern
            AEPapierkorb = AktAuftrag.AEs(idx)

            'Alles Ok. Also weg damit.
            AktAuftrag.DeleteAE(idx)

            'Auftrag speichern
            geaendertenAuftragSichern(AktAuftrag)

            'Neuanzeige, damit die Listbox aktualisiert wird
            AuftragAnzeigen(AktAuftrag.Code)
        End If

        If e.Key = Key.Insert Then
            If AEPapierkorb.Text <> "" Then
                'Sicherheitsabfrage
                If MsgBox("Wollen Sie die Zeile" + vbCrLf + "[" + AEPapierkorb.ToString + "]" + vbCrLf + "wirklich einfügen?", CType(vbYesNo + MsgBoxStyle.Question, MsgBoxStyle)) = vbNo Then Exit Sub

                'Alles Ok. Also wieder anhängen.
                AktAuftrag.AppendAE(AEPapierkorb.Text, AEPapierkorb.Datum, AEPapierkorb.von, AEPapierkorb.bis, 0)

                'Auftrag speichern
                geaendertenAuftragSichern(AktAuftrag)

                'Neuanzeige, damit die Listbox aktualisiert wird
                AuftragAnzeigen(AktAuftrag.Code)
            End If
        End If
    End Sub

    Private Sub lb_AEzeilenZeit_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lb_AEzeilenZeit.SelectionChanged
        lbl_idxZeit.Content = lb_AEzeilenZeit.SelectedIndex.ToString
        ZeitAEMaskeLeeren()
        If lb_AEzeilenZeit.SelectedIndex < 0 Then Exit Sub
        ZeitAEinMaskeUebernehmen(lb_AEzeilenZeit.SelectedIndex)
    End Sub

    Private Sub bt_neueEingabezeileZeit_Click(sender As Object, e As RoutedEventArgs) Handles bt_neueEingabezeileZeit.Click
        ez_tbtext.IsEnabled = True : ez_tbdatum.IsEnabled = True
        ZeitAEMaskeLeeren()
    End Sub

    Private Sub bt_neueEingabezeileStueck_Click(sender As Object, e As RoutedEventArgs) Handles bt_neueEingabezeileStueck.Click
        st_tbtext.IsEnabled = True : st_tbdatum.IsEnabled = True
        StueckAEMaskeLeeren()
    End Sub

    Private Sub st_tbdatum_LostFocus(sender As Object, e As RoutedEventArgs) Handles st_tbdatum.LostFocus
        If st_tbdatum.Text = "" Then Exit Sub
        If testDatumOK(st_tbdatum.Text, st_tbdatum) = False Then MsgBox(st_tbdatum.Text + " ist kein gültiges Datum." + vbCrLf + "Bitte korrigieren, z.B. 10.12.2014")
    End Sub

    Private Sub st_tbdatum_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles st_tbdatum.MouseDoubleClick
        ' Datum von heute einfügen
        st_tbdatum.Text = Today.ToShortDateString
    End Sub

    Private Sub st_tbdatum_MouseEnter(sender As Object, e As MouseEventArgs) Handles st_tbdatum.MouseEnter
        If st_tbdatum.Text = "" Then st_tbdatum.Text = Today.ToShortDateString
    End Sub

    Private Sub st_tbwert_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles st_tbwert.MouseDoubleClick
        st_tbwert.Text = "15"
    End Sub

    Private Sub bt_stueck_OK_Click(sender As Object, e As RoutedEventArgs) Handles bt_stueck_OK.Click
        'Führende Leerzeichen abschneiden
        st_tbtext.Text = st_tbtext.Text.ohneFührendeLeerzeichen

        'Fehleingaben abweisen
        Dim Fehler As String = ""
        If st_tbwert.Text = "" Then
            Fehler = "Wert"
        Else
            Try
                Dim dbl As Double = CDbl(st_tbwert.Text)
            Catch ex As Exception
                Fehler = "gültigen Wert"
            End Try
        End If
        If Fehler <> "" Then
            MsgBox("Bitte einen " + Fehler + " eingeben!")
            st_tbwert.Focus()
            Exit Sub
        End If

        If st_tbtext.Text.Trim = "" Or st_tbdatum.Text.Trim = "" Then
            MsgBox("Bitte einen Text und ein Datum eingeben!")
            st_tbtext.Focus()
            Exit Sub
        End If

        Dim d As Date
        Try
            d = CDate(st_tbdatum.Text)
        Catch ex As Exception
            MsgBox("Bitte ein gültiges Datum eingeben!")
            st_tbdatum.Focus()
            Exit Sub
        End Try

        'Keine doppelten Einträge zulassen
        If AEzeileStueckStatus_NEU Then
            For Each x As AE In AktAuftrag.AEs
                If x.Text = st_tbtext.Text Then
                    MsgBox("Text schon in Liste vorhanden. Bitte eindeutige Texte eingeben.")
                    st_tbtext.Focus()
                    Exit Sub
                End If
            Next
        End If

        'Index der Liste merken (wird bei zusätzlichem Eintrag erhöht)
        Dim idx As Integer = lb_AEzeilenST.SelectedIndex
        'Fallunterscheidung: Neue AE oder Ändern einer AE
        If AEzeileStueckStatus_NEU Then
            'Eine neue AE wurde eingegeben, also anhängen
            AktAuftrag.AppendAE(st_tbtext.Text, CDate(st_tbdatum.Text), Today, Today, CDbl(st_tbwert.Text))
            'Die zuletzt eingetragene AE markieren
            idx = AktAuftrag.AEs.Count - 1

        Else
            'Ändern einer vorhandenen AE
            If idx < 0 Then
                MsgBox("Konnte AE nicht ändern, weil keine AE-Zeile angewählt ist!")
                Exit Sub
            End If
            AktAuftrag.stAEaendern(idx.ToString, st_tbtext.Text, CDate(st_tbdatum.Text), CDbl(st_tbwert.Text))
        End If

        'Auftrag wegschreiben (neu einlesen wäre quatsch, kommt ja nichts neues)
        geaendertenAuftragSichern(AktAuftrag)

        'Stückzeile leeren
        StueckAEMaskeLeeren()

        'Jetzt gibt es eine zusätzliche oder eine geänderte AE, also die Liste aktualisieren
        lb_AEzeilenST.Items.Clear()
        For Each x As AE In AktAuftrag.AEs
            lb_AEzeilenST.Items.Add(x.ToStringStueckFixLength)
        Next

        ' Keinen Eintrag der Liste aktivieren
        lb_AEzeilenST.SelectedIndex = -1


        ' Die Liste auf den richtigen Eintrag setzen
        Try
            'lb_AEzeilenST.SelectedIndex = idx
        Catch ex As Exception
        End Try

        'Auftragskopf aktualisieren
        lbl_AuftragsString.Content = AktAuftrag.toString

        'Jetzt wird eine bereits gesichert AEzeile angezeigt
        'Könnte doppelt sein, weil das Setzen des Indexes bereits dafür sorgt.
        'AEzeileStueckStatus_NEU = False

        ' jetzt noch die NewsZahl aktualisieren
        kwNewsZahlGeaendert(NewsZaehlenKW(AktHeft, aktKWMo, aktKWSo))
    End Sub

    Private Sub bt_zeit_OK_Click(sender As Object, e As RoutedEventArgs) Handles bt_zeit_OK.Click
        'Fehleingaben abweisen
        Dim Fehler As String = ""
        If ez_tbtext.Text.Trim = "" Or ez_tbdatum.Text.Trim = "" Then
            MsgBox("Bitte einen Text und ein Datum eingeben!")
            ez_tbtext.Focus()
            Exit Sub
        End If
        Dim d As Date
        Try
            d = CDate(ez_tbdatum.Text)
            d = CDate(ez_tbvon.Text)
            d = CDate(ez_tbbis.Text)
        Catch ex As Exception
            MsgBox("Bitte ein gültiges Datum sowie gültige Zeiten für von und bis eingeben!")
            ez_tbdatum.Focus()
            Exit Sub
        End Try

        'Keine doppelten Einträge zulassen
        If AEzeileZeitStatus_NEU Then
            For Each x As AE In AktAuftrag.AEs
                If x.Text = ez_tbtext.Text Then
                    MsgBox("Text schon in Liste vorhanden. Bitte eindeutige Texte eingeben.")
                    ez_tbtext.Focus()
                    Exit Sub
                End If
            Next
        End If

        'Index der Liste merken (wird bei zusätzlichem Eintrag erhöht)
        Dim idx As Integer = lb_AEzeilenZeit.SelectedIndex
        'Fallunterscheidung: Neue AE oder Ändern einer AE
        If AEzeileZeitStatus_NEU Then
            'Eine neue AE wurde eingegeben, also anhängen
            AktAuftrag.AppendAE(ez_tbtext.Text, CDate(ez_tbdatum.Text), CDate(ez_tbvon.Text), CDate(ez_tbbis.Text), -1)
            'Die zuletzt eingetragene AE markieren
            idx = AktAuftrag.AEs.Count - 1
        Else
            'Ändern einer vorhandenen AE
            If idx < 0 Then
                MsgBox("Konnte AE nicht ändern, weil keine AE-Zeile angewählt ist!")
                Exit Sub
            End If
            AktAuftrag.stdAEaendern(idx.ToString, ez_tbtext.Text, CDate(ez_tbdatum.Text), CDate(ez_tbvon.Text), CDate(ez_tbbis.Text))
        End If

        'Auftrag wegschreiben (neu einlesen wäre quatsch, kommt ja nichts neues)
        geaendertenAuftragSichern(AktAuftrag)

        'Zeiteingabezeile leeren
        ZeitAEMaskeLeeren()

        'Jetzt gibt es eine zusätzliche oder eine geänderte AE, also die Liste aktualisieren
        lb_AEzeilenZeit.Items.Clear()
        For Each x As AE In AktAuftrag.AEs
            lb_AEzeilenZeit.Items.Add(x.ToStringZeitFixLength)
        Next

        ' Keinen Eintrag der Liste markieren
        lb_AEzeilenZeit.SelectedIndex = -1

        ' Die Liste auf den richtigen Eintrag setzen
        Try
            'lb_AEzeilenZeit.SelectedIndex = idx
        Catch ex As Exception
        End Try

        'Auftragskopf aktualisieren
        lbl_AuftragsString.Content = AktAuftrag.toString

        'Jetzt wird eine bereits gesichert AEzeile angezeigt ??? Abgeschaltet
        'Könnte doppelt sein, weil das Setzen des Indexes bereits dafür sorgt.
        'AEzeileZeitStatus_NEU = False

    End Sub

    Private Sub ez_tbdatum_LostFocus(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles ez_tbdatum.LostFocus
        If ez_tbdatum.Text = "" Then Exit Sub
        If testDatumOK(ez_tbdatum.Text, ez_tbdatum) = False Then MsgBox(ez_tbdatum.Text + " ist kein gültiges Datum." + vbCrLf + "Bitte korrigieren, z.B. 10.12.2014")
    End Sub

    Private Sub ez_tbdatum_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles ez_tbdatum.MouseDoubleClick
        ' Datum von heute einfügen
        ez_tbdatum.Text = Today.ToShortDateString
    End Sub

    Private Sub ez_tbdatum_MouseEnter(sender As Object, e As MouseEventArgs) Handles ez_tbdatum.MouseEnter
        If ez_tbdatum.Text = "" Then ez_tbdatum.Text = Today.ToShortDateString
    End Sub

    Private Sub ez_tbtext_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ez_tbtext.MouseDoubleClick
        ez_tbtext.Text = standardTexte()
    End Sub

    Private Sub ez_tbvon_TextChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles ez_tbvon.TextChanged
        ' wenn sich schon ein Ergebnis berechnen lässt, dann mach's ...
        Try
            Dim von As Date = CDate(ez_tbvon.Text)
            Dim bis As Date = CDate(ez_tbbis.Text)
            ez_lbldauer.Content = Dauer(von, bis)
        Catch ex As Exception
            ' zu früh probiert ... 
        End Try

        Dim l As String = ""
        If ez_lbldauer.Content IsNot Nothing Then l = ez_lbldauer.Content.ToString

        'If l <> "" And l <> "00:00:00" And testDatumOK(ez_tbdatum.Text, ez_tbdatum) Then
        If l <> "" And l <> "00:00:00" Then
            bt_zeit_OK.Visibility = Visibility.Visible
        Else
            bt_zeit_OK.Visibility = Visibility.Hidden
        End If
    End Sub

    Private Sub ez_tbvon_LostFocus(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles ez_tbvon.LostFocus
        Dim t As String = ez_tbvon.Text
        'leer gelassene Felder nicht behelligen
        If t = "" Then Exit Sub

        If testZeitangabeOK(t, ez_tbvon) = False Then
            MsgBox("Fehler: " + t + " ist keine gültige Zeitangabe." + vbCrLf + "Bitte korrigieren, z.B. 10:15")
            ez_tbvon.Text = ""
            ez_tbvon.Focus()
        End If
    End Sub

    Private Sub ez_tbvon_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ez_tbvon.MouseDoubleClick
        'Doppelklick trägt aktuelle Zeit ein - 15-Minuten-Raster
        ez_tbvon.Text = fuenfzehMinRasterNow(Now, False, "ez_tbvon")
    End Sub

    Private Sub ez_tbbis_TextChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles ez_tbbis.TextChanged
        ' wenn sich schon ein Ergebnis berechnen lässt, dann mach's ...
        Try
            Dim von As Date = CDate(ez_tbvon.Text)
            Dim bis As Date = CDate(ez_tbbis.Text)
            ez_lbldauer.Content = Dauer(von, bis)
        Catch ex As Exception
            ' zu früh probiert ... 
        End Try

        Dim l As String = ""
        If ez_lbldauer.Content IsNot Nothing Then l = ez_lbldauer.Content.ToString

        'If l <> "" And l <> "00:00:00" And testDatumOK(ez_tbdatum.Text, ez_tbdatum) Then
        If l <> "" And l <> "00:00:00" Then
            bt_zeit_OK.Visibility = Visibility.Visible
        Else
            bt_zeit_OK.Visibility = Visibility.Hidden
        End If
    End Sub

    Private Sub ez_tbbis_LostFocus(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles ez_tbbis.LostFocus
        Dim t As String = ez_tbbis.Text
        'leer gelassene Felder nicht behelligen
        If t = "" Then Exit Sub

        If testZeitangabeOK(t, ez_tbbis) = False Then
            MsgBox("Fehler: " + t + " ist keine gültige Zeitangabe." + vbCrLf + "Bitte korrigieren, z.B. 10:15")
            ez_tbbis.Text = ""
            ez_tbbis.Focus()
        End If
    End Sub

    Private Sub ez_tbbis_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ez_tbbis.MouseDoubleClick
        'Doppelklick trägt aktuelle Zeit ein - 15-Minuten-Raster
        Try
            ez_tbbis.Text = fuenfzehMinRasterNow(CDate(ez_tbvon.Text), True, "ez_tbbis")
        Catch ex As Exception
            ez_tbbis.Text = fuenfzehMinRasterNow(Now, True, "ez_tbbis")
        End Try
    End Sub

    Private Sub bt_suche_Click(sender As Object, e As RoutedEventArgs) Handles bt_suche.Click
        If tb_sucheAuftraege.Text = "" Then Exit Sub
        SucheNachAuftraegen(tb_sucheAuftraege.Text)
    End Sub

    Private Sub bt_RG_vorbereiten_Click(sender As Object, e As RoutedEventArgs) Handles bt_RG_vorbereiten.Click
        RGVorbereiten()
    End Sub

    Private Sub bt_Clipboard_Click(sender As Object, e As RoutedEventArgs) Handles bt_Clipboard.Click
        Clipboard.SetText(tb_RGtexte.Text)
    End Sub

    Private Sub bt_groesser_Click(sender As Object, e As RoutedEventArgs) Handles bt_groesser.Click
        Dim fs As Integer = CInt(tb_RGtexte.FontSize)
        If fs < 48 Then tb_RGtexte.FontSize = fs + 1
        lbl_fontsize.Content = fs.ToString
    End Sub

    Private Sub bt_kleiner_Click(sender As Object, e As RoutedEventArgs) Handles bt_kleiner.Click
        Dim fs As Integer = CInt(tb_RGtexte.FontSize)
        If fs > 7 Then tb_RGtexte.FontSize = fs - 1
        lbl_fontsize.Content = fs.ToString
    End Sub

    Private Sub lbl_tag_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles lbl_tag.MouseUp
        AuswertungVonBis(Today, Today)
    End Sub

    Private Sub lbl_kw_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles lbl_kw.MouseUp
        AuswertungVonBis(aktKWMo, aktKWSo)
    End Sub

    Private Sub lbl_monat_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles lbl_monat.MouseUp
        Dim erster As Date = Today.AddDays(-(Today.Day - 1))
        Dim letzter As Date = erster.AddMonths(1).AddDays(-1)
        AuswertungVonBis(erster, letzter)
    End Sub

    Private Sub bt_gr_Click(sender As Object, e As RoutedEventArgs) Handles bt_gr.Click
        Dim fs As Integer = CInt(tb_Statistik.FontSize)
        If fs < 48 Then tb_Statistik.FontSize = fs + 1
        lbl_fontsze.Content = fs.ToString
    End Sub

    Private Sub bt_kl_Click(sender As Object, e As RoutedEventArgs) Handles bt_kl.Click
        Dim fs As Integer = CInt(tb_Statistik.FontSize)
        If fs > 7 Then tb_Statistik.FontSize = fs - 1
        lbl_fontsze.Content = fs.ToString
    End Sub

    Private Sub bt_statistik_Tag_Click(sender As Object, e As RoutedEventArgs) Handles bt_statistik_Tag.Click
        AuswertungVonBis(Today, Today)
    End Sub

    Private Sub bt_Gestern_Click(sender As Object, e As RoutedEventArgs) Handles bt_Gestern.Click
        AuswertungVonBis(Today.AddDays(-1), Today.AddDays(-1))
    End Sub

    Private Sub bt_statistik_KW_Click(sender As Object, e As RoutedEventArgs) Handles bt_statistik_KW.Click
        AuswertungVonBis(MoDieserWoche, SoDieserWoche)
    End Sub

    Private Sub bt_Vorwoche_Click(sender As Object, e As RoutedEventArgs) Handles bt_Vorwoche.Click
        AuswertungVonBis(MoDerVorwoche, SoDerVorwoche)
    End Sub

    Private Sub bt_statistik_Monat_Click(sender As Object, e As RoutedEventArgs) Handles bt_statistik_Monat.Click
        If CInt(tb_wievielMonate.Text) < 1 Then Exit Sub
        Dim anz As Integer = 1
        Try
            anz = CInt(tb_wievielMonate.Text)
        Catch ex As Exception
        End Try

        Dim kurzJN As Boolean = CBool(cbx_kurz.IsChecked)

        If anz = 1 Then
            ' Start und Ende des aktuellen Monats ermitteln
            Dim erster As Date = Today.AddDays(-(Today.Day - 1))
            Dim letzter As Date = erster.AddMonths(1).AddDays(-1)
            AuswertungVonBis(erster, letzter, kurzJN, False)
        Else
            Dim d As Date = Today
            For i = anz To 1 Step -1
                ' Beispiel: aktueller Monat = 11, anz= 4
                ' Zielmonate sind 8, 9, 10, 11
                d = Today.AddMonths(1 - i)
                If i = anz Then
                    AuswertungVonBis(MonatsErster(d), MonatsLetzter(d), kurzJN, False)
                Else
                    AuswertungVonBis(MonatsErster(d), MonatsLetzter(d), kurzJN, True)
                End If
            Next
        End If

    End Sub

    Private Sub bt_details_Click(sender As Object, e As RoutedEventArgs) Handles bt_details.Click
        AuswertungHeftAusgabe(AktHeft.ToString + AktAusgabe)
    End Sub

    Private Sub bt_RGfürListe_Click(sender As Object, e As RoutedEventArgs) Handles bt_RGfürListe.Click
        Dim clip As String = CStr(bt_RGfürListe.Content)
        If CBool(InStr(clip, "...")) Then Exit Sub
        If Mid(clip, 1, 3) = "dnp" Then
            tb_Liste1.Text = clip + vbCrLf + tb_Liste1.Text
        ElseIf Mid(clip, 1, 3) = "wum" Then
            tb_Liste2.Text = clip + vbCrLf + tb_Liste2.Text
        Else
            tb_Liste3.Text = clip + vbCrLf + tb_Liste3.Text
        End If
        ti_Listen.Focus()
    End Sub

    Private Sub bt_typ_Click(sender As Object, e As RoutedEventArgs) Handles bt_typ.Click
        Dim t As String = tb_typ.Text
        Select Case t
            Case Is = "std"
                t = "st"
                tb_je.Text = "15"
            Case Is = "st"
                t = "pauschal"
                tb_je.Text = "225"
            Case Else
                t = "std"
                tb_je.Text = "35"
        End Select
        tb_typ.Text = t
    End Sub

    Private Sub lbl_x_Tag_MouseEnter(sender As Object, e As MouseEventArgs) Handles lbl_x_Tag.MouseEnter
        Dim daten As New Dictionary(Of String, Double)
        daten = DatenFuerChartHolen("t")
        myGraph.Children.Clear()
        myGraph.Background = Brushes.DarkSeaGreen
        BalkenGrafik(myGraph, daten, "Gearbeitet: letzte 10 Tage", "Euro", "Datum")
    End Sub

    Private Sub lbl_x_Woche_MouseEnter(sender As Object, e As MouseEventArgs) Handles lbl_x_Woche.MouseEnter
        Dim daten As New Dictionary(Of String, Double)
        daten = DatenFuerChartHolen("w")
        myGraph.Children.Clear()
        myGraph.Background = Brushes.DarkSeaGreen
        BalkenGrafik(myGraph, daten, "Gearbeitet: letzte 10 Wochen", "Euro", "Woche")
    End Sub

    Private Sub lbl_x_Monat_MouseEnter(sender As Object, e As MouseEventArgs) Handles lbl_x_Monat.MouseEnter
        Dim daten As New Dictionary(Of String, Double)
        daten = DatenFuerChartHolen("m")
        myGraph.Children.Clear()
        myGraph.Background = Brushes.DarkSeaGreen
        BalkenGrafik(myGraph, daten, "Gearbeitet: letzte 10 Monate", "Euro", "Monat")
    End Sub

    Private Sub bt_RGChart_Click(sender As Object, e As RoutedEventArgs) Handles bt_RGChart.MouseEnter
        Dim daten As New Dictionary(Of String, Double)
        daten = DatenFuerChartHolen("r")
        myGraph.Children.Clear()
        myGraph.Background = Brushes.DarkSeaGreen
        BalkenGrafik(myGraph, daten, "Abgerechnet: letzte 10 Monate", "Euro", "Monat")
    End Sub

    Private Sub bt_comChart_Click(sender As Object, e As RoutedEventArgs) Handles bt_comChart.MouseEnter
        Dim daten As New Dictionary(Of String, Double)
        daten = DatenFuerChartHolen("com")
        myComGraph.Children.Clear()
        myComGraph.Background = Brushes.DarkSeaGreen
        BalkenGrafik(myComGraph, daten, tb_charttitel.Text, tb_xachse.Text, tb_yachse.Text)
    End Sub

    Private Sub bt_clrChart_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles bt_clrChart.MouseDown
        myComGraph.Children.Clear()
        myComGraph.Background = Brushes.Transparent
    End Sub

    Private Sub bt_comDaten_Mouseup(sender As Object, e As MouseButtonEventArgs) Handles bt_comDaten.MouseUp
        TabWechseln(ti_Daten)
    End Sub

    Private Sub tb_KeyPressed(ByVal sender As System.Object, e As System.Windows.Input.KeyEventArgs) Handles tb_Daten1.KeyDown

        ' Muss ganz vorne bleiben !!
        If e.Key = Key.F1 Then
            If Me.WindowState = Windows.WindowState.Minimized Then Me.WindowState = Windows.WindowState.Normal Else Me.Left = 20
        End If

        ' Text suchen
        If e.Key = Key.F2 Then
            If tb_Daten1.Text = "" Then Exit Sub

            If tb_suche.Text = "" Then
                tb_suche.Focus()
            Else
                ' Cursor hinter die Markierung setzen
                If tb_Daten1.SelectedText.Length > 1 Then tb_Daten1.CaretIndex = tb_Daten1.SelectionStart + tb_Daten1.SelectionLength
                sucheStarten()
            End If
        End If

        'Text ersetzen (einmal)
        If e.Key = Key.F3 Then
            If tb_Daten1.Text = "" Then Exit Sub

            If tb_suche.Text = "" Then
                tb_suche.Focus()
            ElseIf tb_ersetze.Text = "" Then
                tb_ersetze.Focus()
            Else
                If tb_Daten1.SelectedText = tb_suche.Text Then
                    tb_Daten1.SelectionBrush = Windows.Media.Brushes.LightPink
                    tb_Daten1.SelectedText = tb_ersetze.Text
                    ' Cursor hinter die Markierung setzen
                    If tb_Daten1.SelectedText.Length > 1 Then
                        tb_Daten1.CaretIndex = tb_Daten1.SelectionStart + tb_Daten1.SelectionLength
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub tb_suche_KeyPressed(ByVal sender As System.Object, e As System.Windows.Input.KeyEventArgs) Handles tb_suche.KeyDown
        If e.Key = Key.Enter Then sucheStarten()
        If e.Key = Key.F2 And tb_suche.Text.Length > 0 Then sucheStarten()
        If e.Key = Key.F3 Then tb_ersetze.Focus()
    End Sub

    Private Sub tb_ersetze_KeyPressed(ByVal sender As System.Object, e As System.Windows.Input.KeyEventArgs) Handles tb_ersetze.KeyDown
        If e.Key = Key.Enter Then sucheStarten()
        If e.Key = Key.F3 And tb_suche.Text.Length > 0 Then sucheStarten()
        If e.Key = Key.F2 Then tb_suche.Focus()
    End Sub

    Private Sub sucheStarten()
        Dim zuSuchen As String = tb_suche.Text
        If ci = 0 Then
            ci = tb_Daten1.CaretIndex
            If ci < 1 Then ci = 1
        End If
        WeiterSuchen(tb_Daten1, zuSuchen)
    End Sub

    Private Sub TextMarkieren(ByRef tb As TextBox, ByVal start As Integer, ByVal length As Integer)
        tb.Focus()
        tb.SelectionStart = start : tb.SelectionLength = length
    End Sub

    Private Sub WeiterSuchen(ByRef durchSuche As TextBox, ByVal zuSuchen As String)
        Dim l As Integer = zuSuchen.Length
        'Nächstes Vorkommen des Textes suchen
        Dim treffer As Integer = InStr(ci, durchSuche.Text, zuSuchen)
        durchSuche.SelectionBrush = Windows.Media.Brushes.MediumSeaGreen
        If treffer > 1 Then
            TextMarkieren(durchSuche, treffer - 1, zuSuchen.Length)
            ci = treffer + l
        Else
            durchSuche.SelectionStart = 1
            durchSuche.SelectionLength = 0
            ci = 0
        End If
    End Sub

    Private Sub bt_alleErsetzen_Click(sender As Object, e As RoutedEventArgs) Handles bt_alleErsetzen.Click
        If tb_Daten1.Text = "" Or tb_suche.Text = "" Or tb_ersetze.Text = "" Then Exit Sub
        Dim tmp As String = tb_Daten1.Text
        abfrageUndAnzeige(Replace(tb_Daten1.Text, tb_suche.Text, tb_ersetze.Text), tmp, tb_Daten1)
    End Sub

    Private Sub bt_Canvas_kl_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles bt_Canvas_kl.MouseDown
        myComGraph.Height = myComGraph.Height - 60
        myComGraph.Width = myComGraph.Width - 20
        bt_comChart_Click(bt_Canvas_kl, Nothing)
    End Sub

    Private Sub bt_Canvas_gr_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles bt_Canvas_gr.MouseDown
        myComGraph.Height = myComGraph.Height + 40
        myComGraph.Width = myComGraph.Width + 10
        bt_comChart_Click(bt_Canvas_kl, Nothing)
    End Sub

    Private Sub bt_cutAtCurrent_Click(sender As Object, e As RoutedEventArgs) Handles bt_cutAtCurrent.Click
        If tb_Daten1.Text = "" Then Exit Sub
        Dim sb As New System.Text.StringBuilder
        Dim tmp As String = tb_Daten1.Text
        Dim tmp1 As String = tmp
        Dim x As Integer = tb_Daten1.CaretIndex

        'ZeilenAnfang der Zeile suchen in der x steht
        Dim zAnf As Integer = 1
        zAnf = posErstesZeichenDerZeile(tmp, x)
        'MsgBox("x = " + x.ToString + vbCrLf + "zAnf = " + zAnf.ToString)

        'Zeilen vor der Markierung verschonen
        sb.Append(Mid(tmp, 1, zAnf - 1))

        ' zu bearbeitenden Text hernehmen
        tmp1 = Mid(tmp, zAnf)

        'Position von x in der aktuellen zeile bestimmen
        x = x + 1 - zAnf
        If x < 0 Then Exit Sub

        Dim zeilen() As String = Split(tmp1, vbCrLf)

        For Each s As String In zeilen
            'MsgBox(s.Length.ToString + vbCrLf + x.ToString + vbCrLf + ">" + s + "<" + vbCrLf + ">" + Mid(s, 1, x) + "<")
            sb.Append(Mid(s, 1, x) + vbCrLf)
        Next
        abfrageUndAnzeige(sb.ToString, tmp, tb_Daten1)
    End Sub

    Private Sub bt_cutFromBeginn_Click(sender As Object, e As RoutedEventArgs) Handles bt_cutFromBeginn.Click
        If tb_Daten1.Text = "" Then Exit Sub
        Dim sb As New System.Text.StringBuilder
        Dim tmp As String = tb_Daten1.Text
        Dim tmp1 As String = tmp
        Dim x As Integer = tb_Daten1.CaretIndex + 1

        'ZeilenAnfang der Zeile suchen in der x steht (zB. 156)
        Dim zAnf As Integer = 1
        zAnf = posErstesZeichenDerZeile(tmp, x)
        'MsgBox("x = " + x.ToString + vbCrLf + "zAnf = " + zAnf.ToString)

        'Zeilen vor der Markierung verschonen
        sb.Append(Mid(tmp, 1, zAnf - 1))

        ' zu bearbeitenden Text hernehmen
        tmp1 = Mid(tmp, zAnf)

        'Position von x in der aktuellen zeile bestimmen
        x = x + 1 - zAnf
        If x < 0 Then Exit Sub

        Dim zeilen() As String = Split(tmp1, vbCrLf)
        For Each s As String In zeilen
            sb.Append(Mid(s, x) + vbCrLf)
        Next
        abfrageUndAnzeige(sb.ToString, tmp, tb_Daten1)
    End Sub

    Private Sub bt_cutLeerzeilen_Click(sender As Object, e As RoutedEventArgs) Handles bt_cutLeerzeilen.Click
        If tb_Daten1.Text = "" Then Exit Sub
        Dim tmp As String = tb_Daten1.Text
        Dim zeilen() As String = Split(tmp, vbCrLf)
        Dim sb As New System.Text.StringBuilder
        Dim s1 As String
        For Each s As String In zeilen
            s1 = Replace(s, " ", "")
            s1 = Replace(s, Chr(9), "")
            If Len(s1) > 0 Then sb.Append(s + vbCrLf)
        Next
        abfrageUndAnzeige(sb.ToString, tmp, tb_Daten1)
    End Sub

    Private Sub bt_F2ZeilenLoeschen_Click(sender As Object, e As RoutedEventArgs) Handles bt_F2ZeilenLoeschen.Click
        Dim noText As String = tb_Daten1.Text
        If noText = "" Then Exit Sub
        Dim sb As New System.Text.StringBuilder
        Dim zeilen() As String = Split(tb_Daten1.Text, vbCrLf)
        For Each s As String In zeilen
            If InStr(s, tb_suche.Text) < 1 Then
                sb.Append(s + vbCrLf)
            Else
                ' Zeile fällt raus
            End If
        Next
        Dim yesText As String = sb.ToString
        abfrageUndAnzeige(yesText, noText, tb_Daten1)
    End Sub

    Private Sub abfrageUndAnzeige(ByVal yesText As String, ByVal noText As String, ByRef tb As TextBox)
        tb.Text = yesText
        Dim jn As MessageBoxResult
        jn = CType(MsgBox("Änderungen beibehalten?", MsgBoxStyle.YesNoCancel), MessageBoxResult)
        If jn <> MessageBoxResult.Yes Then
            tb.Text = noText
        End If
    End Sub

    Private Function posErstesZeichenDerZeile(ByVal text As String, ByVal x As Integer) As Integer
        Dim zA As Integer = 1
        If text.Length >= x Then
            'Zeilenanfang ist entweder vbcrlf+1 oder der Anfang des Strings
            'Ist es der Anfang des Strings? Dann gibt es vorher kein vbcrlf
            If InStr(text, vbCrLf) < 1 Then
                zA = 1
            Else
                Dim nextCRLF As Integer = 0
                Dim i As Integer = 0
                Do
                    nextCRLF = InStr(zA, text, vbCrLf)
                    If nextCRLF < x Then
                        zA = nextCRLF + 2 'vbcrlf gehört zur Zeile davor
                    Else
                        Exit Do
                    End If
                    i = i + 1 : If i > 10 Then Exit Do
                Loop
            End If
        Else
            ' ungültige Daten 
            If text = "" Then zA = 0 Else zA = 1
        End If

        Return zA
    End Function

    Private Function posInZeile(ByVal text As String, ByVal x As Integer) As Integer
        Dim pos As Integer = 1

        Dim zAnf As Integer = x
        Dim z As String = ""
        While z <> Chr(10)
            zAnf = zAnf - 1
            If zAnf > 1 Then z = Mid(text, zAnf, 1) Else z = Chr(10) : zAnf = zAnf - 1
        End While
        pos = x - zAnf

        Return pos
    End Function

    Private Sub bt_cutBlock_Click(sender As Object, e As RoutedEventArgs) Handles bt_cutBlock.Click
        If tb_Daten1.Text = "" Then Exit Sub
        Dim sb As New System.Text.StringBuilder
        Dim tmp As String = tb_Daten1.Text
        Dim tmp1 As String = tmp
        Dim x As Integer = tb_Daten1.SelectionStart
        Dim delta As Integer = tb_Daten1.SelectionLength
        If delta < 1 Then Exit Sub 'nichts markiert

        'ZeilenAnfang der Zeile suchen in der x steht (zB. 156)
        Dim zAnf As Integer = 1
        zAnf = posErstesZeichenDerZeile(tmp, x)
        'MsgBox("x = " + x.ToString + vbCrLf + "zAnf = " + zAnf.ToString)

        'Zeilen vor der Markierung verschonen
        sb.Append(Mid(tmp, 1, zAnf - 1))

        ' zu bearbeitenden Text hernehmen
        tmp1 = Mid(tmp, zAnf)

        'Position von x und y in der aktuellen zeile bestimmen
        x = x + 1 - zAnf
        Dim y As Integer = x + delta + 1

        If x < 0 Then Exit Sub

        Dim zeilen() As String = Split(tmp1, vbCrLf)
        For Each s As String In zeilen
            'MsgBox(s.Length.ToString + vbCrLf + x.ToString + vbCrLf + y.ToString + vbCrLf + ">" + s + "<" + vbCrLf + ">" + Mid(s, 1, x) + Mid(s, y) + "<")
            sb.Append(Mid(s, 1, x) + Mid(s, y) + vbCrLf)
        Next
        abfrageUndAnzeige(sb.ToString, tmp, tb_Daten1)
    End Sub

    Private Function welcheHefteInMatrixSchreiben(ByVal von As String, ByVal bis As String) As List(Of String)
        Dim Heftliste As New List(Of String)
        Heftliste.Add("dnp" + von) : Heftliste.Add("wum" + von) : Heftliste.Add("com" + von)

        If CInt(von) < CInt(bis) Then
            Dim nxt As String = von  ' Erstmal identisch, hochgezählt wird in der Schleife
            Do
                nxt = HeftNrPlus1(nxt)
                Heftliste.Add("dnp" + nxt) : Heftliste.Add("wum" + nxt) : Heftliste.Add("com" + nxt)
                If nxt = bis Then Exit Do
            Loop
        Else
            ' von und bis sind identisch, Liste ist fertig
        End If

        Return Heftliste
    End Function

    Private Function MatrixErstellen(ByVal Heftliste As List(Of String)) As Double(,)
        Dim Matrix(36, 12) As Double
        Dim ListeDerAuftraege As New List(Of String)
        Dim Heft As String = ""
        Dim Ausgabe As String = ""

        For Each hft In Heftliste
            Heft = Mid(hft, 1, 3)
            Ausgabe = Mid(hft, 4)
            ListeDerAuftraege = ListeDerAuftraegeInListeEinlesen(Heft, Ausgabe, ListeDerAuftraege)
        Next

        Dim mZeile As Integer = 0
        Dim mMonat As Integer = 0
        Dim HeftAusg As String = ""

        For Each kode As String In ListeDerAuftraege
            HeftAusg = Mid(kode, 1, 7)
            'Matrix-Zeile ermitteln
            mZeile = Heftliste.IndexOf(HeftAusg) + 1

            'Auftrag einlesen
            Dim a As New Auftrag
            a = AuftragEinLesen(kode) : If a Is Nothing Then Exit For
            For Each ae As AE In a.AEs
                'Monat ermitteln
                mMonat = ae.Datum.Month
                'ae.Wert zum Matrixwert addieren
                Matrix(mZeile, mMonat) += ae.Wert
                'ae.Wert zu den Matrixsummen addieren
                Matrix(0, mMonat) += ae.Wert
                Matrix(mZeile, 0) += ae.Wert
            Next
        Next

        Return Matrix
    End Function

    Private Function ZahlNachMonat(ByVal Zahl As Integer) As String
        Dim Monatstext As String = ""
        Select Case Zahl
            Case Is = 1
                Monatstext = "Jan"
            Case Is = 2
                Monatstext = "Feb"
            Case Is = 3
                Monatstext = "Mrz"
            Case Is = 4
                Monatstext = "Apr"
            Case Is = 5
                Monatstext = "Mai"
            Case Is = 6
                Monatstext = "Juni"
            Case Is = 7
                Monatstext = "Juli"
            Case Is = 8
                Monatstext = "Aug"
            Case Is = 9
                Monatstext = "Sept"
            Case Is = 10
                Monatstext = "Okt"
            Case Is = 11
                Monatstext = "Nov"
            Case Is = 12
                Monatstext = "Dez"
            Case Else
                Monatstext = "n.A."
        End Select
        Return Monatstext
    End Function

    Private Sub MatrixSchreiben(ByVal Matrix As Double(,), ByRef Hefte As List(Of String), ByRef tb As TextBox)
        tb.Text = ""
        Dim Trenner = " | "
        Dim zaehler As Integer = 0

        Dim s As New System.Text.StringBuilder

        s.Append("       " + Trenner)
        ' Kopfzeile schreiben
        For Monat = 1 To 12
            If Matrix(0, Monat) > 0 Then
                s.Append(StringFuellen(ZahlNachMonat(Monat), 5, vonVorne:=True) + Trenner)
                zaehler += 1
            End If
        Next
        s.Append(StringFuellen("Summe", 6, vonVorne:=True) + Trenner)

        Dim Linie = xMal("-", 18 + zaehler * 8)
        Dim doppelLinie = xMal("=", 18 + zaehler * 8)


        For zeile = 1 To Hefte.Count
            If Matrix(zeile, 0) > 0 Then
                s.Append(vbCrLf + Linie)
                s.Append(vbCrLf + Hefte(zeile - 1) + Trenner) ' Die Hefte-Liste ist Nullbasiert!
                For Monat = 1 To 12
                    If Matrix(0, Monat) > 0 Then
                        s.Append(StringFuellen(OhneNachkommaStellen(Matrix(zeile, Monat).ToString), 5, vonVorne:=True) + Trenner)
                    End If
                Next
                s.Append(StringFuellen(OhneNachkommaStellen(Matrix(zeile, 0).ToString), 6, vonVorne:=True) + Trenner)
            End If
        Next
        ' Summenzeile für die Monate
        s.Append(vbCrLf + Linie)
        s.Append(vbCrLf + "Gesamt " + Trenner)
        Dim Gesamt As Double = 0
        For Monat = 1 To 12
            If Matrix(0, Monat) > 0 Then
                Gesamt += Matrix(0, Monat)
                s.Append(StringFuellen(OhneNachkommaStellen(Matrix(0, Monat).ToString), 5, vonVorne:=True) + Trenner)
            End If
        Next
        s.Append(StringFuellen(OhneNachkommaStellen(Gesamt.ToString), 6, vonVorne:=True) + Trenner)
        s.Append(vbCrLf + doppelLinie + vbCrLf)

        tb.Text = s.ToString

    End Sub

    Private Function Matrix_Heftvorgaben_ok(ByVal von As String, ByVal bis As String) As Boolean
        Dim tf As Boolean = False
        Try
            If CInt(von) > 1404 And
               CInt(bis) < 3001 And
               CInt(von) <= CInt(bis) And
               CInt(bis) - CInt(von) < 101 Then
                tf = True
            End If
        Catch ex As Exception
        End Try

        Return tf
    End Function

    Private Sub bt_Matrix_Click(sender As Object, e As RoutedEventArgs) Handles bt_Matrix.Click
        If Matrix_Heftvorgaben_ok(tb_Matrix_von.Text, tb_Matrix_bis.Text) Then
            Dim Hefte As New List(Of String)
            Hefte = welcheHefteInMatrixSchreiben(tb_Matrix_von.Text, tb_Matrix_bis.Text)
            Dim Matrix(36, 12) As Double
            Matrix = MatrixErstellen(Hefte)
            MatrixSchreiben(Matrix, Hefte, tb_Statistik)
        Else
            MsgBox("Keine gültigen Angaben für Start/Ende der Matrix")
        End If
    End Sub

#Region "TabDaten"

    Private Function tbDatenUeberschreiben() As Boolean
        Dim antw As Integer
        antw = MsgBox("Den aktuellen Text überschreiben?", MsgBoxStyle.YesNo, "Bitte bestätigen")
        'MsgBox(antw.ToString)
        Dim ret As Boolean = True
        If antw <> 6 Then ret = False
        Return ret
    End Function

    Private Sub bt_ladeTag_Click(sender As Object, e As RoutedEventArgs) Handles bt_ladeTag.Click
        Dim noText As String = tb_Daten1.Text
        Dim yesText As String = Replace(qTextDateiLesen(blConst.TAGESDATEIPFAD, False), ":", " ")
        abfrageUndAnzeige(yesText, noText, tb_Daten1)
    End Sub

    Private Sub bt_ladeKW_Click(sender As Object, e As RoutedEventArgs) Handles bt_ladeKW.Click
        Dim noText As String = tb_Daten1.Text
        Dim yesText As String = Replace(qTextDateiLesen(blConst.KWDATEIPFAD, False), ":", " ")
        abfrageUndAnzeige(yesText, noText, tb_Daten1)
    End Sub

    Private Sub bt_ladeMonat_Click(sender As Object, e As RoutedEventArgs) Handles bt_ladeMonat.Click
        Dim noText As String = tb_Daten1.Text
        Dim yesText As String = Replace(qTextDateiLesen(blConst.MONATSDATEIPFAD, False), ":", " ")
        abfrageUndAnzeige(yesText, noText, tb_Daten1)
    End Sub

    Private Sub bt_ladeJahr_Click(sender As Object, e As RoutedEventArgs) Handles bt_ladeJahr.Click
        Dim noText As String = tb_Daten1.Text
        Dim yesText As String = Replace(qTextDateiLesen(blConst.JAHRESDATEIPFAD, False), ":", " ")
        abfrageUndAnzeige(yesText, noText, tb_Daten1)
    End Sub

    Private Sub tb_Daten1_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles tb_Daten1.MouseRightButtonUp
        Dim text As String = ""
        If tb_Daten1.SelectionLength > 1 Then text = tb_Daten1.SelectedText Else text = tb_Daten1.Text
        Dim ausw As String = textAuswerten(text)
        Clipboard.SetText(ausw)
        MsgBox(ausw)
    End Sub

    Private Function textAuswerten(ByVal text As String) As String
        If text = "" Then Return "" : Exit Function
        Dim erg As New System.Text.StringBuilder
        Dim zeichen As Integer = text.Length
        Dim zeilen As String() = Split(text, vbCrLf)
        Dim anzZeilen As Integer = zeilen.Count - zaehleZeilenOhneZahl(zeilen)
        Dim zahlenStrings() As String
        Dim werteMatrix(anzZeilen, 6) As Double
        Dim werteCount(6) As Integer
        Dim wert As Double = 0
        Dim AnzZeros(6) As Integer

        Dim aktZeile As Integer = 1
        For Each z As String In zeilen
            If z = "" Then
            Else
                zahlenStrings = TextAnalyseModul.selectNumbers(z)
                If zahlenStrings IsNot Nothing Then
                    Dim aktSpalte As Integer = 1
                    For Each x As String In zahlenStrings
                        If x IsNot Nothing Then
                            Try
                                wert = CDbl(x) : If wert = 0 Then AnzZeros(aktSpalte) += 1
                                werteMatrix(aktZeile, aktSpalte) = wert
                                werteCount(aktSpalte) += 1
                                aktSpalte += 1
                            Catch ex As Exception
                                werteMatrix(aktZeile, aktSpalte) = 0
                            End Try
                        End If
                    Next
                    aktZeile += 1
                End If
            End If
        Next

        'Spaltensummen ausrechnen, in Index 0 schreiben
        For i = 1 To 6  'max. Zahl Spalten
            For y = 1 To anzZeilen
                werteMatrix(0, i) += werteMatrix(y, i)
            Next
        Next

        'Min und Max suchen
        Dim min(6) As Double
        Dim max(6) As Double

        For y = 1 To anzZeilen
            For i = 1 To 6
                If y = 1 Then
                    min(i) = werteMatrix(y, i)
                    max(i) = werteMatrix(y, i)
                Else
                    If werteMatrix(y, i) < min(i) Then
                        min(i) = werteMatrix(y, i)
                    End If
                    If werteMatrix(y, i) > max(i) Then
                        max(i) = werteMatrix(y, i)
                    End If
                End If

            Next
        Next

        'Ergebnis zusammenbauen
        erg.Append(zeichen.ToString + " Zeichen | ")
        erg.Append(anzZeilen.ToString + " Zeilen" + vbCrLf)
        For j = 1 To 6
            If werteMatrix(0, j) > 0 Then
                erg.Append(vbCrLf + "Summe (" + j.ToString + ") " + Math.Round(werteMatrix(0, j), 2).ToString + vbCrLf + "     ")
                erg.Append(werteCount(j).ToString + " Werte" + " | Schnitt ")
                erg.Append(Math.Round(werteMatrix(0, j) / werteCount(j), 2).ToString + vbCrLf + "     ")
                erg.Append("Min " + Math.Round(min(j), 2).ToString)
                erg.Append(" | Max " + Math.Round(max(j), 2).ToString + vbCrLf + "     ")
                erg.Append("Anz. Nullwerte " + AnzZeros(j).ToString + " (" + Math.Round(AnzZeros(j) / werteCount(j) * 100, 1).ToString + " %)" + vbCrLf)
            End If
        Next

        Return erg.ToString
    End Function

    Private Sub tb_Daten1_TextChanged(sender As Object, e As TextChangedEventArgs) Handles tb_Daten1.TextChanged
        Dim zeichen As Integer = tb_Daten1.Text.Length
        Dim zeilen As Integer = Split(tb_Daten1.Text, vbCrLf).Count
        lbl_textInfo.Content = "Zeilen: " + zeilen.ToString + " / " + "Zeichen: " + zeichen.ToString
    End Sub

    Private Sub bt_goText_Click(sender As Object, e As RoutedEventArgs) Handles bt_goText.Click
        ' Text generieren für die Jobkralle-Abfrage

        Dim c13 As String = Chr(13)
        Dim c34 As String = Chr(34)
        Dim t As String = "|"

        Dim z1 As String = "Muster|http://www.jobkralle.de/jobs?title=qyWAS&location=qyWO&radius=0&type%5B%5D=0&type%5B%5D=1&type%5B%5D=2&type%5B%5D=3&type%5B%5D=4&type%5B%5D=5&age=31&sorting=0&display=1" + c13
        Dim z2 As String = "sp1|<meta name=" + c34 + "description" + c34 + " content=" + c34 + c13


        Dim ArrStaedte() As String = {"München", "Hamburg", "Berlin", "Stuttgart", "Frankfurt am Main", "Nürnberg",
                                      "Köln", "Düsseldorf", "Leipzig", "Dresden", "Hannover", "Dortmund",
                                      "Bremen", "Essen", "Duisburg"}

        Dim ArrLaender() As String = {"Bayern", "Baden-Württemberg", "Nordrhein-Westfalen", "Hessen", "Niedersachsen",
                                      "Bremen", "Hamburg", "Saarland", "Schleswig-Holstein", "Rheinland-Pfalz", "Berlin",
                                      "Sachsen", "Thüringen", "Sachsen-Anhalt", "Brandenburg", "Mecklenburg-Vorpommern"}


        tb_Daten1.Clear()
        Dim sb As New System.Text.StringBuilder
        sb.Append(z1 + z2)

        Dim arr() As String
        If rb_laender.IsChecked Then
            arr = ArrLaender
        Else
            arr = ArrStaedte
        End If

        For Each x In arr
            sb.Append(c34 + tbwort1.Text + c34 + t + x + t + tbwort2.Text + c13)
            If tbwort1alt.Text <> "" Then
                sb.Append(c34 + tbwort1alt.Text + c34 + t + x + t + tbwort2alt.Text + c13)
            End If
        Next

        tb_Daten1.Text = sb.ToString
        Clipboard.SetText(tb_Daten1.Text)

    End Sub

    Private Sub rb_staedte_Checked(sender As Object, e As RoutedEventArgs) Handles rb_staedte.Checked
        If rb_staedte IsNot Nothing And rb_laender IsNot Nothing Then
            If rb_staedte.IsChecked Then
                rb_laender.IsChecked = False
            Else
                rb_laender.IsChecked = True
            End If
        End If
    End Sub

    Private Sub rb_laender_Checked(sender As Object, e As RoutedEventArgs) Handles rb_laender.Checked
        If rb_staedte IsNot Nothing And rb_laender IsNot Nothing Then
            If rb_laender.IsChecked Then
                rb_staedte.IsChecked = False
            Else
                rb_staedte.IsChecked = True
            End If
        End If
    End Sub
#End Region

#Region "UI-Events von PDF-Laufzettel"
    Private Sub lfz_clear_Click(sender As Object, e As RoutedEventArgs) Handles lfz_clear.Click
        tbRubrik.Text = "" : tbThema.Text = "" : tbSeite.Text = ""
        tbDatum.Text = ""
        tbA1.Text = "" : tbA2.Text = "" : tbA3.Text = "" : tbA4.Text = ""
        tb_zielpfad.Text = "" : tb_zieldatei.Text = ""
        PDFVorgabenEintragen()
    End Sub

    Private Sub bt_PDFschreiben_Click(sender As Object, e As RoutedEventArgs) Handles bt_PDFschreiben.Click
        If FormularIstVollstaendig() Then
            If prüfePfad(tb_zielpfad) = True Then
                Kopie = tb_zielpfad.Text + "Laufzettel_leer.PDF"
                If Not System.IO.File.Exists(Kopie) Then Kopie = ArbeitskopieAnlegen(Kopie)
                'DatenInPDFeintragen(Kopie, tb_zielpfad.Text + tb_zieldatei.Text)
                DatenInsPDFeintragen()
                System.IO.File.Delete(Kopie)
            Else
                Dim melde As String = "Es gibt ein Problem mit Dateiname und -pfad:" + vbCrLf + tb_zielpfad.Text + tb_zieldatei.Text
                tb_zielpfad.Focus()
                If tb_zielpfad.Text <> "" Then tb_zielpfad.SelectionStart = tb_zielpfad.Text.Length
                MsgBox(melde)
            End If
        Else
            Dim melde As String = "Das Formular wurde nicht vollständig ausgefüllt oder die Datei existiert bereits!"
            MsgBox(melde)
        End If
    End Sub

    Private Sub tbSeite_GotFocus(sender As Object, e As RoutedEventArgs) Handles tbSeite.GotFocus, tbA1.GotFocus
        ThemaFuerDateinameUndPfad()
    End Sub

    Private Sub tbAusgabe_LostFocus(sender As Object, e As RoutedEventArgs) Handles tbAusgabe.LostFocus
        Dim m As Integer = CInt(Replace(Mid(tbAusgabe.Text, 1, 2), "-", "")) - 2
        If m <= 0 Then m += 12
        PfadFuerDenLaufzettelSetzen(MonatZurNummer(m))
    End Sub

    Private Sub tb_zielpfad_LostFocus(sender As Object, e As RoutedEventArgs) Handles tb_zielpfad.LostFocus
        prüfePfad(tb_zielpfad)
    End Sub

    Private Sub plus1_Click(sender As Object, e As RoutedEventArgs) Handles plus1.Click
        Dim m As Integer = CInt(Replace(Mid(tbAusgabe.Text, 1, 2), "-", ""))
        Dim jahr As Integer = CInt(Mid(tbAusgabe.Text, tbAusgabe.Text.Length - 3))

        If m = 12 Then
            m = 0
            jahr += 1
        End If
        m = m + 1

        Dim neu As String = m.ToString + "-" + jahr.ToString
        tbAusgabe.Text = neu
        HeftNrFuerDateiname(Mid(jahr, 3) + NullVorWert(m.ToString, 2))

        m = m - 2
        If m <= 0 Then m += 12
        PfadFuerDenLaufzettelSetzen(MonatZurNummer(m))

    End Sub

    Private Sub tbA1_KeyUp(sender As Object, e As KeyEventArgs) Handles tbA1.KeyUp
        If Asc(e.Key.ToString) < 32 Or Asc(e.Key.ToString) > 128 Then Exit Sub
        If tbA1.Text.Length >= maxA1 Then tbA2.Focus()
    End Sub

    Private Sub tbA2_KeyUp(sender As Object, e As KeyEventArgs) Handles tbA2.KeyUp
        If Asc(e.Key.ToString) < 32 Or Asc(e.Key.ToString) > 128 Then Exit Sub
        If tbA2.Text.Length >= maxA2bis4 Then tbA3.Focus()
    End Sub

    Private Sub tbA3_KeyUp(sender As Object, e As KeyEventArgs) Handles tbA3.KeyUp
        If Asc(e.Key.ToString) < 32 Or Asc(e.Key.ToString) > 128 Then Exit Sub
        If tbA3.Text.Length >= maxA2bis4 Then tbA4.Focus()
    End Sub

    Private Sub tbA4_KeyUp(sender As Object, e As KeyEventArgs) Handles tbA4.KeyUp
        If Asc(e.Key.ToString) < 32 Or Asc(e.Key.ToString) > 128 Then Exit Sub
        If tbA4.Text.Length >= maxA2bis4 Then
            MsgBox("Anmerkungszeilen sind voll!")
        End If
    End Sub

#End Region

    ' Kontextmenü der Auftragsliste
    Private Sub Loeschen_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        AuftragLoeschenInPapierkorb()
    End Sub

    Private Sub Verschieben_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        VerschiebeAuftragInNaechstesHeft()
    End Sub

    Private Sub Umbenennen_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        AuftragUmbenennen()
    End Sub




#Region "Programmstarter"

    Private Sub But1_Click(sender As Object, e As RoutedEventArgs) Handles But1.Click
        blHilfsModul2014.starteProg(ProgPfade(1))
    End Sub

    Private Sub But2_Click(sender As Object, e As RoutedEventArgs) Handles But2.Click
        blHilfsModul2014.starteProg(ProgPfade(2))
    End Sub

    Private Sub But3_Click(sender As Object, e As RoutedEventArgs) Handles But3.Click
        blHilfsModul2014.starteProg(ProgPfade(3))
    End Sub

    Private Sub But4_Click(sender As Object, e As RoutedEventArgs) Handles But4.Click
        blHilfsModul2014.starteProg(ProgPfade(4))
    End Sub

    Private Sub But5_Click(sender As Object, e As RoutedEventArgs) Handles But5.Click
        blHilfsModul2014.starteProg(ProgPfade(5))
    End Sub

    Private Sub But6_Click(sender As Object, e As RoutedEventArgs) Handles But6.Click
        blHilfsModul2014.starteProg(ProgPfade(6))
    End Sub

    Private Sub But7_Click(sender As Object, e As RoutedEventArgs) Handles But7.Click
        blHilfsModul2014.starteProg(ProgPfade(7))
    End Sub

    Private Sub But8_Click(sender As Object, e As RoutedEventArgs) Handles But8.Click
        blHilfsModul2014.starteProg(ProgPfade(8))
    End Sub

    Private Sub But9_Click(sender As Object, e As RoutedEventArgs) Handles But9.Click
        blHilfsModul2014.starteProg(ProgPfade(9))
    End Sub

    Private Sub But10_Click(sender As Object, e As RoutedEventArgs) Handles But10.Click
        blHilfsModul2014.starteProg(ProgPfade(10))
    End Sub

    Private Sub But11_Click(sender As Object, e As RoutedEventArgs) Handles But11.Click
        blHilfsModul2014.starteProg(ProgPfade(11))
    End Sub

    Private Sub But12_Click(sender As Object, e As RoutedEventArgs) Handles But12.Click
        blHilfsModul2014.starteProg(ProgPfade(12))
    End Sub

    Private Sub But13_Click(sender As Object, e As RoutedEventArgs) Handles But13.Click
        blHilfsModul2014.starteProg(ProgPfade(13))
    End Sub

    Private Sub bt_ladeProgPfade_Click(sender As Object, e As RoutedEventArgs) Handles bt_ladeProgPfade.Click
        Dim noText As String = tb_Daten1.Text
        Dim yesText As String = TextDateiLesen(blConst.PfadZuProgPfadeText, True)
        abfrageUndAnzeige(yesText, noText, tb_Daten1)
    End Sub

    Private Sub bt_speichernProgPfade_Click(sender As Object, e As RoutedEventArgs) Handles bt_speichernProgPfade.Click
        ' Erste Zeile muss stimmen
        If Mid(tb_Daten1.Text, 1, 20) <> "Pfade zu Anwendungen" Then Exit Sub
        ' Sicherungskopie anlegen
        FileCopy(blConst.PfadZuProgPfadeText, blConst.PfadZuProgPfadeText + ".sik")
        If MsgBox("Pfade zu Anwendungen wirklich speichern?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then
            TextDateiSchreiben(blConst.PfadZuProgPfadeText, tb_Daten1.Text, True)
            ProgPfade = Split(TextDateiLesen(blConst.PfadZuProgPfadeText, False), vbCrLf)
            'IconsEinfuegen() führt hier oft zum Absturz, besser Neustart abwarten
        End If
    End Sub

    Private Sub Btp1_Click(sender As Object, e As RoutedEventArgs) Handles Btp1.Click
        blHilfsModul2014.starteProg(ProgPfade(14))
    End Sub

    Private Sub Btp2_Click(sender As Object, e As RoutedEventArgs) Handles Btp2.Click
        blHilfsModul2014.starteProg(ProgPfade(15))
    End Sub
    Private Sub Btp3_Click(sender As Object, e As RoutedEventArgs) Handles Btp3.Click
        blHilfsModul2014.starteProg(ProgPfade(16))
    End Sub
    Private Sub Btp4_Click(sender As Object, e As RoutedEventArgs) Handles Btp4.Click
        blHilfsModul2014.starteProg(ProgPfade(17))
    End Sub
    Private Sub Btp5_Click(sender As Object, e As RoutedEventArgs) Handles Btp5.Click
        blHilfsModul2014.starteProg(ProgPfade(18))
    End Sub
    Private Sub Btp6_Click(sender As Object, e As RoutedEventArgs) Handles Btp6.Click
        blHilfsModul2014.starteProg(ProgPfade(19))
    End Sub
    Private Sub Btp7_Click(sender As Object, e As RoutedEventArgs) Handles Btp7.Click
        blHilfsModul2014.starteProg(ProgPfade(20))
    End Sub


#End Region

End Class
