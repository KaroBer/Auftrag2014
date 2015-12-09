Imports System.Runtime.CompilerServices
Module myExtensions

    <Extension()>
    Public Function f(ByVal s As String, ByVal c As myColumn) As String
        Return myTable.format(s, c)
    End Function

    <Extension()>
    Public Function f(ByVal d As Double, ByVal c As myColumn) As String
        Return myTable.format(d, c)
    End Function

    <Extension()>
    Public Function ohneFührendeLeerzeichen(ByVal t As String) As String
        If t.First = " " Then
            t = t.TrimStart(CChar(" ")).ohneFührendeLeerzeichen()
        End If

        Return t
    End Function


    <Extension()>
    Public Function Mal(ByVal s As String, ByVal x As Integer) As String
        Dim sb As New System.Text.StringBuilder
        For i = 1 To x
            sb.Append(s)
        Next
        Return sb.ToString
    End Function

    <Extension()>
    Public Function snText(ByVal s As String) As String
        Dim ret As String = s
        If s = "" Then
            ret = "none"
        ElseIf s = " " Then
            ret = "space"
        ElseIf s = Chr(9) Then
            ret = "tab"
        ElseIf s = vbCr Then
            ret = "CR"
        ElseIf s = vbLf Then
            ret = "LF"
        ElseIf s = vbCrLf Then
            ret = "CRLF"
        End If
        Return ret
    End Function

End Module
