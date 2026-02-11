Option Strict On
Option Explicit On

Imports System
Imports System.IO

' ============================================================
' UISettings — Read/write theme preference to output/ui.settings
' Simple INI-style: THEME=Light or THEME=Dark
' VB.NET 4.0 compatible, no external dependencies.
' ============================================================
Public Module UISettings

    Private Const SETTINGS_FILENAME As String = "ui.settings"

    Private Function GetSettingsPath() As String
        Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, SETTINGS_FILENAME)
    End Function

    ''' <summary>
    ''' Load the saved theme name. Returns "Light" or "Dark".
    ''' Defaults to "Light" if file missing or unreadable.
    ''' </summary>
    Public Function LoadThemeName() As String
        Try
            Dim p As String = GetSettingsPath()
            If Not File.Exists(p) Then Return "Light"

            Dim lines() As String = File.ReadAllLines(p)
            For Each rawLine As String In lines
                Dim line As String = rawLine.Trim()
                If line = "" OrElse line.StartsWith("#") OrElse line.StartsWith(";") Then Continue For

                Dim eq As Integer = line.IndexOf("="c)
                If eq <= 0 Then Continue For

                Dim key As String = line.Substring(0, eq).Trim().ToUpperInvariant()
                Dim value As String = line.Substring(eq + 1).Trim()

                If key = "THEME" Then
                    If value.ToUpperInvariant() = "DARK" Then Return "Dark"
                    Return "Light"
                End If
            Next
        Catch
        End Try

        Return "Light"
    End Function

    ''' <summary>
    ''' Save the theme name to ui.settings.
    ''' </summary>
    Public Sub SaveThemeName(name As String)
        Try
            Dim p As String = GetSettingsPath()
            Dim safeName As String = If(name IsNot Nothing AndAlso name.Trim().ToUpperInvariant() = "DARK", "Dark", "Light")

            ' Read existing lines and update THEME key, or append it
            Dim lines As New System.Collections.Generic.List(Of String)()
            Dim found As Boolean = False

            If File.Exists(p) Then
                For Each rawLine As String In File.ReadAllLines(p)
                    Dim trimmed As String = rawLine.Trim()
                    Dim eq As Integer = trimmed.IndexOf("="c)
                    If eq > 0 Then
                        Dim key As String = trimmed.Substring(0, eq).Trim().ToUpperInvariant()
                        If key = "THEME" Then
                            lines.Add("THEME=" & safeName)
                            found = True
                            Continue For
                        End If
                    End If
                    lines.Add(rawLine)
                Next
            End If

            If Not found Then
                lines.Add("THEME=" & safeName)
            End If

            File.WriteAllLines(p, lines.ToArray())
        Catch
            ' Never crash for settings save failure
        End Try
    End Sub

End Module
