Option Strict On

Public Class blConst
    Public Const STANDARDPFAD = "D:\myProgData\blAuftrag2014\"
    Public Const AKTAUFTRAGDATEI = "AktAuftrag"
    Public Const DATEIENDUNG = ".dat"
    Public Const PAPIERKORB = "D:\myProgData\blAuftrag2014\papierkorb\"
    Public Const TAGESDATEIPFAD = "D:\myProgData\blAuftrag2014\StatistikTag.txt"
    Public Const KWDATEIPFAD = "D:\myProgData\blAuftrag2014\StatistikKW.txt"
    Public Const MONATSDATEIPFAD = "D:\myProgData\blAuftrag2014\StatistikMonat.txt"
    Public Const JAHRESDATEIPFAD = "D:\myProgData\blAuftrag2014\StatistikJahr.txt"
    Public Const ZIPpfad = "D:\Skydrive\blA14sik\"
    Public Const PfadZuProgPfadeText = STANDARDPFAD + "ProgPfade.txt"
    Public Const PfadZuWERruftAN = STANDARDPFAD & "blWerRuftAn.vb.exe"

    Public Const CSVTRENNER = "|"
    Public Const ZEILENENDE = vbCrLf
End Class

Class Auftrag
    Inherits blConst
    Public Enum Typ
        std = 1
        st = 2
        pauschal = 3
    End Enum
    Public Enum Heft
        dnp = 1
        wum = 2
        com = 3
    End Enum

    Property Kunde As Heft
    Property Ausgabe As String
    Property Text As String
    Property Art As Typ
    Property Satz As Double
    Property Bisher As Double

    ' Der aktuelle Wert darf von außen nicht geändert werden
    ' Setzen des Wertes über setAktWert nur von dieser Klasse aus möglich
    ' Das Setzen löst einen Event aus, damit die Statistik angepasst werden kann.
    ' Lesen des Wertes über AktWert auch von der GUI aus möglich.
    ' Das Datum brauchts für die Statistik um festzustellen, welcher Tag geändert wird
    Dim _DatumZumGeaendertenWert As Date
    Dim _aktWert As Double
    Private WriteOnly Property setAktWert As Double
        Set(value As Double)
            Dim Delta = value - _aktWert
            _aktWert = value
            RaiseEvent ValueChanged(Me, Delta, _DatumZumGeaendertenWert)
        End Set
    End Property
    ReadOnly Property AktWert As Double
        Get
            Return _aktWert
        End Get
    End Property

    ' Der Code dient quasi als ID für den Auftrag, 
    ' er darf nicht geändert werden
    Dim _code As String
    ReadOnly Property Code As String
        Get
            Return _code
        End Get
    End Property

    Property HeftPfad As String
    Property AusgabePfad As String
    Property DateiPfadUndName As String
    Property AEs As List(Of AE)

    Public Event ValueChanged(ByVal a As Auftrag, ByVal Delta As Double, ByVal Datum As Date)

    Public Sub New()
    End Sub

    Public Sub New(ByVal h As Auftrag.Heft, ByVal ausgabe As String, ByVal text As String, ByVal std_st As Auftrag.Typ, ByVal je As Double)
        Me.Kunde = h
        Me.Ausgabe = ausgabe
        Me.Text = text
        Me.Art = std_st
        Me.Satz = je
        Me.Bisher = 0

        Me._DatumZumGeaendertenWert = Today
        Me.setAktWert = 0

        Me._code = Kunde.ToString + ausgabe + text
        Me.HeftPfad = STANDARDPFAD + Kunde.ToString + "\"
        Me.AusgabePfad = Me.HeftPfad + Kunde.ToString + ausgabe + "\"
        Me.DateiPfadUndName = Me.AusgabePfad + Me.Code + DATEIENDUNG
        Me.AEs = New List(Of AE)
    End Sub

    Public Sub AppendAE(ByVal t As String, ByVal d As Date, ByVal v As Date, ByVal b As Date, ByVal EingabeWert As Double)
        Dim neueAE As New AE
        neueAE.Text = t : neueAE.Datum = d : neueAE.von = v : neueAE.bis = b
        neueAE.Dauer = neueAE.getDauer
        If Art = Typ.std Then
            Me.Bisher = Math.Round(Me.Bisher + CDbl((neueAE.Dauer.TotalMinutes / 60)), 2)
            neueAE.Wert = Math.Round(CDbl((neueAE.Dauer.TotalMinutes / 60) * Me.Satz), 2)
        ElseIf Art = Typ.st Then
            neueAE.Wert = EingabeWert
            Me.Bisher = Me.Bisher + 1
        ElseIf Art = Typ.pauschal Then
            neueAE.Wert = EingabeWert
            Me.Bisher = Me.Bisher + Math.Round(neueAE.Wert / Me.Satz, 2)
        Else
            neueAE.Wert = EingabeWert
        End If
        Me.AEs.Add(neueAE)
        Me.addWert(neueAE.Wert, neueAE.Datum)
    End Sub

    Public Sub stAEaendern(ByVal idx As String, ByVal t As String, ByVal d As Date, ByVal wert As Double)
        If CDbl(idx) < 0 Or CDbl(idx) > Me.AEs.Count - 1 Then Exit Sub
        Dim ae As AE = Me.AEs(CInt(idx))
        Dim alterWert As Double = ae.Wert
        Dim altesDatum As Date = ae.Datum

        ae.Text = t
        ae.Datum = d
        ae.Wert = wert

        StatistikPflegeBeiAenderung(alterWert, ae.Wert, altesDatum, ae.Datum)
        
        If alterWert <> ae.Wert Then Me.addWert(ae.Wert - alterWert, ae.Datum)
    End Sub

    Private Sub StatistikPflegeBeiAenderung(ByVal alterWert As Double, ByVal neuerWert As Double, altesDatum As Date, neuesDatum As Date)
        If altesDatum <> neuesDatum Then
            ' Es gibt einen Event zum zurücksetzen des Wertes für das alte Datum
            ' der zweite Event wird dann über addWert gefeuert
            RaiseEvent ValueChanged(Me, -alterWert, altesDatum)

            If alterWert = neuerWert Then
                ' Hier würde der zweite Event über addWert keine Änderung liefern
                ' Es muss der ursprüngliche Wert aber dem neuen Datum zugeordnet werden
                RaiseEvent ValueChanged(Me, alterWert, neuesDatum)
            End If
        End If
    End Sub

    Private Sub addWert(ByVal x As Double, ByVal DatumZumWert As Date)
        Me._DatumZumGeaendertenWert = DatumZumWert
        Me.setAktWert = Me.AktWert + x
    End Sub

    Private Sub resetWert()
        Me._DatumZumGeaendertenWert = Today
        Me.setAktWert = 0.0
    End Sub

    Public Sub stdAEaendern(ByVal idx As String, ByVal t As String, ByVal d As Date, ByVal v As Date, ByVal b As Date)
        If CDbl(idx) < 0 Or CDbl(idx) > Me.AEs.Count - 1 Then Exit Sub
        Dim ae As AE = Me.AEs(CInt(idx))
        Dim alterWert As Double = ae.Wert
        Dim altesBisher As Double = CDbl((ae.Dauer.TotalMinutes / 60))
        Dim altesDatum As Date = ae.Datum

        ae.Text = t
        ae.Datum = d
        ae.von = v
        ae.bis = b
        ae.Dauer = ae.getDauer()
        Me.Bisher = Math.Round(Me.Bisher + CDbl((ae.Dauer.TotalMinutes / 60)) - altesBisher, 2)
        ae.Wert = Math.Round(CDbl((ae.Dauer.TotalMinutes / 60) * Me.Satz), 2)

        If altesDatum <> d Then
            ' Getrennte Aufrufe von addWert feuern zwei getrennte Events, damit stimmt dann auch die Statistik
            Me.addWert(-alterWert, altesDatum)
            Me.addWert(ae.Wert, d)
        Else
            ' Die If-Abfrage verhindert, dass wegen einer Null-Änderung ein Event ausgelöst wird
            If alterWert <> ae.Wert Then Me.addWert(ae.Wert - alterWert, ae.Datum)
        End If
    End Sub

    Public Sub DeleteAE(ByVal idx As Integer)
        If idx < 0 Or idx > Me.AEs.Count - 1 Then Exit Sub

        Dim WertDerAE As Double = Me.AEs(idx).Wert
        Dim datum As Date = Me.AEs(idx).Datum

        'Sicherheitsabfrage erledigt der Aufrufer!
        'Weg damit.
        Me.AEs.RemoveAt(idx)

        If Me.Art = Typ.st Then Me.Bisher = Me.AEs.Count
        Me.addWert(-WertDerAE, datum)

    End Sub



    Public Function toDetailString() As String
        Dim t As String = " | "
        Dim s As New System.Text.StringBuilder
        s.Append(StringFuellen(Me.Code, 24, True, ".", False) + t)
        s.Append(StringFuellen(Me.Art.ToString, 6, True, ".", True) + t)
        s.Append(StringFuellen(Me.Bisher.ToString, 4, False, " ", False) + t)
        s.Append(StringFuellen(StellenHintermKomma(Me.AktWert, 2) + " Euro", 12, False, " ", True))

        Return s.ToString
    End Function



    Public Overrides Function toString() As String
        Dim t As String = " -- "
        Return Me.Code + t + Me.Bisher.ToString + t + Me.Art.ToString + t + " je: " + Me.Satz.ToString + t + Me.AktWert.ToString + t + Me.AEs.Count.ToString + " AEs" + ZEILENENDE
    End Function

    Public Function toCSV() As String
        Dim t As String = CSVTRENNER
        Dim CSVBuilder As New System.Text.StringBuilder
        CSVBuilder.Append(Me.Kunde.ToString + t + Ausgabe + t + Text + t + Me.Art.ToString + t + Me.Satz.ToString + t + Me.Bisher.ToString + t + Me.AktWert.ToString + ZEILENENDE)
        If AEs.Count > 0 Then
            For Each e As AE In AEs
                CSVBuilder.Append(e.toCSV)
            Next
        End If
        Return CSVBuilder.ToString
    End Function

    Public Function CSVtoObject(ByVal CSVString As List(Of String)) As Auftrag
        Dim ersteZeile As String = CSVString(0)
        Dim proptys() As String = Split(CSVString(0), CSVTRENNER)
        If proptys.Count < 7 Then
            MsgBox("zu wenige Eigenschaften gefunden!") : Stop
        End If

        Dim a As New Auftrag(
            CType([Enum].Parse(GetType(Heft), proptys(0)), Heft),
            proptys(1),
            proptys(2),
            CType([Enum].Parse(GetType(Typ), proptys(3)), Typ),
            CDbl(proptys(4))
            )
        a.Bisher = CDbl(proptys(5))
        a.setAktWert = CDbl(proptys(6))

        Dim wert As Double = 0
        Dim bisher As Double = 0
        For i = 1 To CSVString.Count - 1
            Dim ae As New AE()
            a.AEs.Add(ae.CSVtoObject(CSVString(i)))
            'Korrektur falsch berechneter Wertangaben
            wert = wert + a.AEs(i - 1).Wert
            bisher = Math.Round(bisher + CDbl((a.AEs(i - 1).getDauer.TotalMinutes / 60)), 2)
        Next
        a.setAktWert = wert
        If a.Art = Typ.std Then
            a.Bisher = bisher
        Else
            a.Bisher = a.AEs.Count
        End If
        Return a
    End Function

    ' Hilfsfunktionen

End Class
