Imports SchwabenCode.QuickIO
Module QuickIOmodul
    Function qTextDateiLesen(ByVal pfad As String, ByVal fehlermeldung As Boolean) As String
        Dim ret As String = ""
        If QuickIOFile.Exists(pfad) Then
            Try
                ret = QuickIOFile.ReadAllText(pfad)
            Catch except As Exception
                If fehlermeldung Then MsgBox(except.Message & vbNewLine & "Fehler in TextDateiLesen()", MsgBoxStyle.Exclamation)
            End Try
        End If
        Return ret
    End Function

    Function qTextDateiSchreiben(ByVal pfad As String, ByVal Text As String, ByVal fehlermeldung As Boolean) As Boolean
        Dim ret As Boolean = True
        Try
            QuickIOFile.WriteAllText(pfad, Text)
        Catch except As Exception
            ret = False
            If fehlermeldung Then MsgBox(except.Message & vbNewLine & "Fehler in TextDateiSchreiben()", MsgBoxStyle.Exclamation)
        End Try

        Return ret
    End Function

    Function qTextZeilenLesen(ByVal pfad As String, ByVal fehlermeldung As Boolean) As List(Of String)
        Dim zeilen As New List(Of String)
        If QuickIOFile.Exists(pfad) Then
            Try
                Dim erg = QuickIOFile.ReadAllLines(pfad)
                For Each z In erg
                    zeilen.Add(z.ToString)
                Next
            Catch except As Exception
                If fehlermeldung Then MsgBox(except.Message & vbNewLine & "Fehler in TextZeilenLesen()", MsgBoxStyle.Exclamation)
            End Try
        End If
        Return zeilen
    End Function

    Function qListeAllerUnterordner(ByVal Pfad As String) As List(Of String)
        ' Nicht nur im obersten Directory sondern auch die in den Unterordnern
        Dim ordner As Object
        Dim ListeDerOrdner As New List(Of String)
        If QuickIODirectory.Exists(Pfad) Then
            ordner = QuickIODirectory.EnumerateDirectories(Pfad, System.IO.SearchOption.AllDirectories)
            For Each o In ordner
                ListeDerOrdner.Add(o.FullName)
            Next
        End If

        Return ListeDerOrdner
    End Function

    Function qListeDerUnterordner(ByVal Pfad As String) As List(Of String)
        'Nur die Unterordner im angegebenen Ordner
        Dim ordner As Object
        Dim ListeDerOrdner As New List(Of String)
        If QuickIODirectory.Exists(Pfad) Then
            ordner = QuickIODirectory.EnumerateDirectories(Pfad, System.IO.SearchOption.TopDirectoryOnly)
            For Each o In ordner
                ListeDerOrdner.Add(o.FullName)
            Next
        End If

        Return ListeDerOrdner
    End Function

    Function qListeDerDateiNamenImOrdner(ByVal Pfad As String, ByVal DateiEndung As String) As List(Of String)
        If DateiEndung <> "*" Then DateiEndung = DateiEndung.ToUpper
        Dim ListeDerDateinamen As New List(Of String)

        If QuickIODirectory.Exists(Pfad) Then
            Dim dateien As Object = QuickIODirectory.EnumerateFiles(Pfad, System.IO.SearchOption.TopDirectoryOnly)
            Dim f As String = ""
            For Each d In dateien
                f = d.FullName
                If DateiEndung = "*" Then
                    ListeDerDateinamen.Add(f)
                Else
                    If f.ToUpper.EndsWith(DateiEndung) Then ListeDerDateinamen.Add(f)
                End If
            Next
        End If
        Return ListeDerDateinamen
    End Function

    Function qListeAllerDateienImOrdnerUndDessenUnterordnern(ByVal Pfad As String, Optional ByVal DateiEndung As String = "*") As List(Of String)
        If DateiEndung <> "*" Then DateiEndung = DateiEndung.ToUpper
        Dim ListeDerDateinamen As New List(Of String)

        If QuickIODirectory.Exists(Pfad) Then
            Dim dateien As Object = QuickIODirectory.EnumerateFiles(Pfad, System.IO.SearchOption.AllDirectories)
            Dim f As String = ""
            For Each d In dateien
                f = d.FullName
                If DateiEndung = "*" Then
                    ListeDerDateinamen.Add(f)
                Else
                    If f.ToUpper.EndsWith(DateiEndung) Then ListeDerDateinamen.Add(f)
                End If
            Next
        End If
        Return ListeDerDateinamen
    End Function

End Module

