Option Strict On

Imports System
Imports System.IO
Imports System.IO.Compression

Module blHilfsModul2014

    Function RegistryReadKeyValue(ByVal AppName As String, ByVal Section As String, ByVal KeyName As String, Optional DefaultValue As String = "") As String
        Dim erg As String = ""
        If Trim(Section) <> "" And
           Trim(KeyName) <> "" Then
            Try
                erg = GetSetting(AppName, Section, KeyName, DefaultValue)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation,
                  "Fehler beim Zugriff auf die Registry")
            End Try
        End If
        Return erg
    End Function

    Sub RegistryWriteKeyValue(ByVal AppName As String, ByVal Section As String, ByVal KeyName As String, ByVal Value As String)
        If Trim(Section) <> "" And
           Trim(KeyName) <> "" And
           Trim(Value) <> "" Then
            Try
                SaveSetting(AppName, Section, KeyName, Value)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Fehler beim Zugriff auf die Registry")
            End Try
        End If
    End Sub

    Function GetNumberInString(ByVal s As String) As Integer
        ' Funktioniert, wenn s = sldkfjlaskf12345 lkjlk"  --> 12345
        ' Liefert Unsinn bei s = sddlk 123 dlkfjld 44455" --> 12344455
        ' Überlaufgrenze wird nicht berücksichtigt!
        ' Negative Zahlen können nicht gelesen werden.
        Dim NewStr As String = ""
        For x = 1 To Len(s)
            If InStr("0123456789", Mid(s, x, 1)) > 0 Then
                NewStr += Mid(s, x, 1)
            End If
        Next
        Return CInt(NewStr)
    End Function

    Sub DirectoryCopy(ByVal sourceDirName As String, ByVal destDirName As String, ByVal copySubDirs As Boolean, ByVal overwriteFiles As Boolean)
        'Stammt aus dem MSDN - basiert auf .NET 4.5

        ' Get the subdirectories for the specified directory.
        Dim dir As DirectoryInfo = New DirectoryInfo(sourceDirName)
        Dim dirs As DirectoryInfo() = dir.GetDirectories()

        If Not dir.Exists Then
            Throw New DirectoryNotFoundException( _
                "Source directory does not exist or could not be found: " _
                + sourceDirName)
        End If

        ' If the destination directory doesn't exist, create it.
        If Not Directory.Exists(destDirName) Then
            Directory.CreateDirectory(destDirName)
        Else
        End If

        ' Get the files in the directory and copy them to the new location.
        Dim files As FileInfo() = dir.GetFiles()
        For Each file In files
            Dim temppath As String = Path.Combine(destDirName, file.Name)
            file.CopyTo(temppath, overwriteFiles)
        Next file

        ' If copying subdirectories, copy them and their contents to new location.
        If copySubDirs Then
            For Each subdir In dirs
                Dim temppath As String = Path.Combine(destDirName, subdir.Name)
                DirectoryCopy(subdir.FullName, temppath, copySubDirs, overwriteFiles)
            Next subdir
        End If
    End Sub

    Function ZippeOrdnerSchnell(ByVal Ordner As String, ByVal ZipDateiMitPfad As String, ByVal ueberschreiben As Boolean) As Boolean
        Dim tf As Boolean = False
        If Not Directory.Exists(Ordner) Then Return tf

        ' ZIP-Datei ggf. löschen
        If File.Exists(ZipDateiMitPfad) Then
            If ueberschreiben Then File.Delete(ZipDateiMitPfad) Else Return tf
        End If
        tf = True
        ZipFile.CreateFromDirectory(Ordner, ZipDateiMitPfad, CompressionLevel.Fastest, True)

        Return tf
    End Function

    Function ZippeOrdnerOptimal(ByVal Ordner As String, ByVal ZipDateiMitPfad As String, ByVal ueberschreiben As Boolean) As Boolean
        Dim tf As Boolean = False
        If Not Directory.Exists(Ordner) Then Return tf

        ' ZIP-Datei ggf. löschen
        If File.Exists(ZipDateiMitPfad) Then
            If ueberschreiben Then File.Delete(ZipDateiMitPfad) Else Return tf
        End If
        tf = True
        ZipFile.CreateFromDirectory(Ordner, ZipDateiMitPfad, CompressionLevel.Optimal, True)

        Return tf
    End Function

    Function xMal(ByVal s As String, ByVal x As Integer) As String
        Dim sb As New System.Text.StringBuilder
        If x < 1 Then Return ""
        For i = 1 To x
            sb.Append(s)
        Next
        Return sb.ToString
    End Function

    Function StringFuellen(ByVal s As String, ByVal anz As Integer, Optional ByVal kürzen As Boolean = False, Optional ByVal Füllzeichen As String = " ", Optional vonVorne As Boolean = False) As String
        Dim sb As New System.Text.StringBuilder
        Dim v As String = xMal(Füllzeichen, anz)

        Dim vorn As Integer = 0
        If Not kürzen Then If s.Length > anz Then anz = s.Length

        If Not vonVorne Then
            sb.Append(s + Mid(v, 1, anz))
        Else
            If s.Length > anz And kürzen Then s = Mid(s, 1, anz)
            vorn = anz - s.Length
            If vorn = 0 Then
                sb.Append(s)
            ElseIf vorn > 0 Then
                sb.Append(Mid(v, 1, vorn) + s)
            End If
        End If

        Return Mid(sb.ToString, 1, anz)
    End Function

    Function NumFormat(ByVal wert As Double, ByVal NachKommaStellen As Integer, ByVal PlatzVormKomma As Short, Optional ByVal Fuellzeichen As String = " ") As String
        Dim ret As String = ""
        ret = StellenHintermKomma(wert, NachKommaStellen)

        Dim dot As Integer = InStr(ret, ",")
        Dim anz As Integer = PlatzVormKomma
        If dot > 0 Then
            If dot = 1 Then
                ret = "0" + ret  ' Erstes Zeichen ist ein Komma
                dot = 2
            End If
            anz = 1 + PlatzVormKomma - dot
        Else
            ' Kein Komma gefunden
            anz = PlatzVormKomma - ret.Length
        End If

        Return xMal(Fuellzeichen, anz) + ret
    End Function

    Function StellenHintermKomma(ByVal wert As Double, ByVal stellen As Integer) As String
        Dim ret As String = Math.Round(wert, stellen).ToString
        Dim dot As Integer = InStr(ret, ",")
        If dot < 1 Then
            'Zahl hat gar keine Nachkommastellen
            ret = ret + "," + NullVorWert("", stellen)
        Else
            Dim vormKomma As Integer = dot - 1
            ret = ret + NullVorWert("", stellen) ' Maximale Zahl an Nullen anhängen
            ret = Mid(ret, 1, vormKomma + 1 + stellen)
        End If
        If stellen = 0 Then ret = Replace(ret, ",", "")
        Return ret
    End Function

    Function OhneNachkommaStellen(ByVal s As String) As String
        ' Verbesserte und getestete Version vom 30. 4. 2014
        Dim ret As String = "0"
        Try
            Dim zahl As Double = CDbl(s)
            ret = Math.Truncate(zahl).ToString
        Catch ex As Exception
            ' keine Zahl, es wird 0 zurückgegeben
        End Try
        Return ret

    End Function

    Function ZweiStellenHintermKomma(ByVal wert As String) As String
        Dim ret As String = wert
        Try
            ret = Math.Round(CDbl(wert), 2).ToString
        Catch ex As Exception
            
        End Try

        Dim dot As Integer = InStr(ret, ",")
        If dot < 1 Then
            ret = ret + ",00"
        Else
            Dim l As Integer = ret.Length
            If l - dot = 1 Then ret = ret + "0"
        End If
        Return ret
    End Function

    Function Dauer(ByVal von As Date, ByVal bis As Date) As TimeSpan
        Dim d As TimeSpan = Today - Today
        Try
            If bis > von Then d = bis - von
        Catch ex As Exception
            ' keine Meldung. Zurückgegeben wird "00:00:00"
        End Try
        Return d
    End Function

    Function testDatumOK(ByVal d As String, ByRef tb As TextBox) As Boolean
        Dim tf = False
        Try
            Dim dt As Date = CDate(d)
            tf = True
        Catch ex As Exception
            d = DatumKorrekturversuch(d)
            Try
                Dim dt As Date = CDate(d)
                tf = True
            Catch exx As Exception
                tf = False
            End Try
        End Try

        If tf = True Then
            'Datum formatieren
            tb.Text = CDate(d).ToShortDateString
        Else
            tb.Text = ""
            tb.Focus()
        End If
        Return tf
    End Function

    Function DatumKorrekturversuch(ByVal d As String) As String
        ' Leerstring und unsinnige Eingaben gar nicht erst durch diese Prüfung lassen, also ab 5.5.5 --> 05.05.2005
        If d.Length > 4 Then
            ' doppelte Punkte durch einfache ersetzen
            Do
                d = Replace(d, "..", ".")
            Loop Until InStr(d, "..") < 1

            ' nur für 310314 --> 31032014
            If d.Length = 6 And InStr(d, ".") = 0 Then
                Dim yfirst2 As String = Mid(Today.Year.ToString, 1, 2)
                Dim ylast2 As String = Mid(Today.Year.ToString, 3)
                If Mid(d, 3, 2) <> yfirst2 And Mid(d, 5) = ylast2 Then
                    d = Mid(d, 1, 4) + yfirst2 + ylast2
                End If
            End If

            'vergessene Punkte einfügen
            If InStr(d, ".") = 0 And d.Length > 5 Then
                d = Mid(d, 1, 2) + "." + Mid(d, 3, 2) + "." + Mid(d, 5)
            End If

            If d.Length = 9 Then
                If InStr(d, "." + Today.Year.ToString) > 0 Then
                    d = Mid(d, 1, 2) + "." + Mid(d, 3)
                Else
                    d = Mid(d, 1, 5) + "." + Mid(d, 6)
                End If
            End If
        End If
        Return d
    End Function

    Function ZeitangabeKorrekturversuch(ByVal t As String) As String
        ' Tippfehler Punkt, Strichpunkt oder Komma statt Doppelpunkt korrigieren
        If CBool(InStr(t, ".")) Then t = Replace(t, ".", ":")
        If CBool(InStr(t, ",")) Then t = Replace(t, ",", ":")
        If CBool(InStr(t, ";")) Then t = Replace(t, ";", ":")

        If t.Length = 4 And InStr(t, ":") < 1 Then t = Mid(t, 1, 2) + ":" + Mid(t, 3)
        If t.Length = 2 And InStr(t, ":") < 1 Then t = Mid(t, 1, 2) + ":00"

        If t.Length = 3 And t.IndexOf(":") < 1 Then t = "0" + Mid(t, 1, 1) + ":" + Mid(t, 2) ' 930 --> 09:30

        Return t
    End Function

    Function testZeitangabeOK(ByVal t As String, ByRef tb As TextBox) As Boolean
        Dim tf As Boolean = False
        Dim d As Date
        Try
            d = CDate(t)
            tf = True
        Catch ex As Exception
            t = ZeitangabeKorrekturversuch(t)
            Try
                d = CDate(t)
                tf = True
            Catch exx As Exception
                tf = False
            End Try
        End Try
        If tf = True Then
            tb.Text = d.ToShortTimeString
        Else
            tb.Text = ""
            tb.Focus()
        End If
        Return tf
    End Function

    Function fuenfzehMinRasterNow(ByVal start As Date, ByVal plus As Boolean, ByVal sender As String) As String
        ' plus = true heißt die viertelstunde später / false = früher

        Dim std As Integer = start.Hour
        Dim min As Integer = start.Minute

        If plus Then
            ' vorwärtsrechnen
            If min <= 15 Then min = 15
            If min < 59 And min >= 45 Then min = 0 : std = std + 1
            If min < 45 And min >= 30 Then min = 45
            If min < 30 And min >= 15 Then min = 30

        Else
            ' zurückrechnen
            If min >= 45 Then min = 45
            If min < 45 And min >= 30 Then min = 30
            If min < 30 And min >= 15 Then min = 15
            If min < 15 Then min = 0
        End If

        ' Gegebenenfalls eine 0 voranstellen
        Dim h As String = NullVorWert(std.ToString, 2)
        Dim m As String = NullVorWert(min.ToString, 2)

        Return h + ":" + m
    End Function

    Function NullVorWert(ByVal wert As String, ByVal Stellen As Integer, Optional ByVal Fuellzeichen As String = "0") As String
        ' Getestete Version vom 30. 4. 2014

        ' Werte mit mehr als den geforderten Stellen vor einem eventuellen Komma werden nicht gekürzt
        Dim komma As Integer = InStr(wert, ",")

        If komma > 0 Then
            ' Ein Komma ist enthalten
            ' 12,1234  -> Komma = 3, Nachkommastellen = 4, Length = 7

            Dim hintermKomma As Integer = wert.Length - komma
            Dim vormKomma As Integer = wert.Length - 1 - hintermKomma

            ' Hat schon genügend Stellen vor dem Komma (es wird nichts abgeschnitten!)
            If vormKomma >= Stellen Then Return wert
            Try
                ' Negative Werte werden nicht verändert
                Dim test As Double = CDbl(wert)
                If test < 0 Then Return wert

            Catch ex As Exception
                ' Es ist kein Wert, sondern ein Text, 
                ' das Komma wird hier nicht berücksichtigt
                vormKomma = 0
            End Try

            ' Füllzeichen davorsetzen
            wert = xMal(Fuellzeichen, Stellen - vormKomma) + wert
        Else
            ' Es ist kein Komma drin
            ' Werte mit mehr als den geforderten Stellen werden nicht gekürzt
            If wert.Length >= Stellen Then Return wert
            Try
                ' Negative Werte werden nicht verändert
                Dim test As Double = CDbl(wert)
                If test < 0 Then Return wert
            Catch ex As Exception
                ' Es ist kein Wert, sondern ein Text, darf trotzdem durchlaufen
            End Try

            ' Füllzeichen davorsetzen
            wert = xMal(Fuellzeichen, Stellen - wert.Length) + wert
            For i = 1 To Stellen - wert.Length
                wert = Fuellzeichen + wert
            Next
        End If
        Return wert
    End Function

    Function Kalenderwoche(ByVal d As Date) As Integer
        Dim kw As Integer = 0
        Dim y As Integer = CInt(d.Year)
        kw = DatePart(Microsoft.VisualBasic.DateInterval.WeekOfYear, d, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        If kw > 52 Then
            If y <> 2015 And y <> 2020 And y <> 2026 Then
                kw = 1 'Fehler im Allgorithmus: kw 53 gibts nur 2015, 2020 und 2026 (später interessierts mich nicht)
            End If
        End If
        Return kw
    End Function

    Function IstWE(ByVal d As Date) As Boolean
        Dim ret As Boolean = False
        Dim tg As String = WeekdayName(Weekday(d, FirstDayOfWeek.Monday))
        tg = Mid(tg, 1, 2)
        If tg = "Sa" Or tg = "So" Then ret = True
        Return ret
    End Function

    Function IsNotWE(ByVal d As Date) As Boolean
        Return Not IstWE(d)
    End Function

    Function istFixFeiertagInBayern(ByVal d As Date) As Boolean
        Dim fixFeiertageBy As String = "0101,0601,0105,2512,2612,3112"
        Return aufFeiertagTesten(d, fixFeiertageBy)
    End Function

    Function istFixFeiertagInDtl(ByVal d As Date) As Boolean
        Dim fixFeiertageDtl As String = "0101,0105,2512,2612,3112"
        Return aufFeiertagTesten(d, fixFeiertageDtl)
    End Function

    Private Function aufFeiertagTesten(ByVal d As Date, ByVal Feiertage As String) As Boolean
        Dim ttmm As String = NullVorWert(d.Day.ToString, 2) + NullVorWert(d.Month.ToString, 2)
        If InStr(Feiertage, ttmm) > 0 Then Return True Else Return False
    End Function

    Function MonatsErster(ByVal d As Date) As Date
        Return d.AddDays(-(d.Day - 1))
    End Function

    Function MonatsLetzter(ByVal d As Date) As Date
        Return MonatsErster(d).AddMonths(1).AddDays(-1)
    End Function

    Function MoDieserWoche() As Date
        Return Today.AddDays(-(Today.DayOfWeek - 1))
    End Function

    Function SoDieserWoche() As Date
        Return Today.AddDays(7 - Today.DayOfWeek)
    End Function

    Function MoDerVorwoche() As Date
        Dim DatumInVorwoche As Date = Today.AddDays(-7)
        Dim TagDerVorwoche As DayOfWeek = DatumInVorwoche.DayOfWeek
        Return DatumInVorwoche.AddDays(-(TagDerVorwoche - 1))
    End Function

    Function SoDerVorwoche() As Date
        Dim DatumInVorwoche As Date = Today.AddDays(-7)
        Dim TagDerVorwoche As DayOfWeek = DatumInVorwoche.DayOfWeek
        Return DatumInVorwoche.AddDays(7 - TagDerVorwoche)
    End Function

    Function zaehleZeilenOhneLeerzeilen(ByVal zeilen() As String) As Integer
        Dim z0 As Integer = 0
        For Each z In zeilen
            If z = "" Then z0 += 1
        Next
        Dim erg As Integer = zeilen.Count - z0
        If erg < 0 Then erg = 0
        Return erg
    End Function

    Sub starteProg(ByVal pfad As String)
        If pfad = "" Then Exit Sub
        Try
            Dim proc As New Process
            proc = Process.Start(pfad)

        Catch ex As Exception
            MsgBox(pfad + vbCrLf + ex.ToString)
        End Try
    End Sub

    Function erstesIconHolen(ByVal p As String) As Image
        Dim img As New Image

        Try
            Dim ico As System.Drawing.Icon = System.Drawing.Icon.ExtractAssociatedIcon(p)
            Dim strm As New MemoryStream
            Dim bmp = ico.ToBitmap
            bmp.Save(strm, System.Drawing.Imaging.ImageFormat.Png)
            strm.Seek(0, SeekOrigin.Begin)
            Dim pbd As New PngBitmapDecoder(strm, BitmapCreateOptions.None, BitmapCacheOption.Default)
            img.Source = pbd.Frames(0)
        Catch ex As Exception
            'MsgBox(p & vbCrLf & ex.ToString)
        End Try
        Return img
    End Function

    Function WochentagDeutsch(ByVal wotag As Integer) As String
        Dim ret As String = wotag.ToString
        If wotag = 1 Then
            ret = "Montag"
        ElseIf wotag = 2 Then
            ret = "Dienstag"
        ElseIf wotag = 3 Then
            ret = "Mittwoch"
        ElseIf wotag = 4 Then
            ret = "Donnerstag"
        ElseIf wotag = 5 Then
            ret = "Freitag"
        ElseIf wotag = 6 Then
            ret = "Samstag"
        ElseIf wotag = 7 Then
            ret = "Sonntag"
        End If
        Return ret
    End Function

End Module
