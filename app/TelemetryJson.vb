Option Strict On
Option Explicit On

Imports System
Imports System.Text

Public NotInheritable Class TelemetryJson
    Private Sub New()
    End Sub

    Public Shared Function EscapeString(value As String) As String
        If value Is Nothing Then Return ""
        Dim sb As New StringBuilder(value.Length + 16)

        Dim i As Integer
        For i = 0 To value.Length - 1
            Dim ch As Char = value.Chars(i)

            Select Case ch
                Case """"c
                    sb.Append("\""")
                Case "\"c
                    sb.Append("\\")
                Case "/"c
                    sb.Append("\/")
                Case ChrW(8)   ' backspace
                    sb.Append("\b")
                Case ChrW(12)  ' formfeed
                    sb.Append("\f")
                Case ChrW(10)  ' newline
                    sb.Append("\n")
                Case ChrW(13)  ' carriage return
                    sb.Append("\r")
                Case ChrW(9)   ' tab
                    sb.Append("\t")
                Case Else
                    Dim code As Integer = AscW(ch)
                    If code < 32 OrElse code > 126 Then
                        sb.Append("\u")
                        sb.Append(code.ToString("x4"))
                    Else
                        sb.Append(ch)
                    End If
            End Select
        Next

        Return sb.ToString()
    End Function

    Public Shared Function Quote(value As String) As String
        Return """" & EscapeString(value) & """"
    End Function
End Class
