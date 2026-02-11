Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Text.RegularExpressions

Public Class MacroAutoDetect

    '============================================================
    ' Auto-detect Module + Procedure from .swp
    '============================================================
    Public Shared Function Detect(
        macroPath As String,
        ByRef moduleName As String,
        ByRef procName As String) As Boolean

        moduleName = ""
        procName = ""

        If Not File.Exists(macroPath) Then
            Return False
        End If

        Dim txt As String = ""

        Try
            txt = File.ReadAllText(macroPath)
        Catch
            Return False
        End Try

        ' --- VB Module name ---
        Dim mMatch As Match =
            Regex.Match(txt, "VB_Name\s*=\s*""([^""]+)""", RegexOptions.IgnoreCase)

        If mMatch.Success Then
            moduleName = mMatch.Groups(1).Value.Trim()
        End If

        ' --- First Sub procedure ---
        Dim pMatch As Match =
            Regex.Match(txt,
                        "^[\t ]*(Public|Private)?[\t ]*Sub[\t ]+([A-Za-z_][A-Za-z0-9_]*)[\t ]*\(",
                        RegexOptions.IgnoreCase Or RegexOptions.Multiline)

        If pMatch.Success Then
            procName = pMatch.Groups(2).Value.Trim()
        End If

        Return (moduleName <> "" AndAlso procName <> "")

    End Function

End Class
