Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Windows.Forms

Public Class MacroEngine

    '============================================================
    ' Run Design Tool by Slot ID
    '============================================================
    Public Shared Sub RunDesignTool(
        slotId As Integer,
        swVersion As String,
        selectedAssembly As String,
        logBox As TextBox)

        If selectedAssembly = "" Then
            MessageBox.Show("Please select an assembly first.")
            Return
        End If

        Dim macroPath As String =
            MacroRegistry.GetMacroPath(slotId)

        If macroPath = "" OrElse Not File.Exists(macroPath) Then
            Log(logBox, "Macro not found for slot " & slotId)
            Return
        End If

        Dim swApp As Object = Nothing

        Try
            swApp = CreateObject("SldWorks.Application." & GetSWProgID(swVersion))
            swApp.Visible = True
        Catch ex As Exception
            Log(logBox, "SolidWorks connection failed: " & ex.Message)
            Return
        End Try

        Dim errs As Integer = 0
        Dim warns As Integer = 0

        Try
            swApp.OpenDoc6(selectedAssembly, 2, 1 Or 64, "", errs, warns)
        Catch ex As Exception
            Log(logBox, "Failed to open assembly: " & ex.Message)
            Return
        End Try

        Log(logBox, "Running macro: " & Path.GetFileName(macroPath))

        Dim modName As String = ""
        Dim procName As String = ""

        If MacroAutoDetect.Detect(macroPath, modName, procName) Then
            RunMacro(swApp, macroPath, modName, procName, logBox)
            Exit Sub
        End If

        ' Fallbacks
        If RunMacro(swApp, macroPath, Nothing, Nothing, logBox) Then Exit Sub
        If RunMacro(swApp, macroPath, Nothing, "main", logBox) Then Exit Sub

        Log(logBox, "Macro failed.")

    End Sub

    '============================================================
    ' Safe Macro Run
    '============================================================
    Private Shared Function RunMacro(
        swApp As Object,
        path As String,
        moduleName As Object,
        procName As Object,
        logBox As TextBox) As Boolean

        Try
            Dim longErr As Integer = 0
            Dim ok As Boolean =
                swApp.RunMacro2(path, moduleName, procName, 2, longErr)

            If ok Then
                Log(logBox, "Macro executed successfully.")
            Else
                Log(logBox, "Macro failed (error " & longErr & ")")
            End If

            Return ok

        Catch ex As Exception
            Log(logBox, "Macro exception: " & ex.Message)
            Return False
        End Try

    End Function

    '============================================================
    ' SolidWorks Version → ProgID
    '============================================================
    Private Shared Function GetSWProgID(v As String) As String
        Select Case v
            Case "2022" : Return "30"
            Case "2023" : Return "31"
            Case "2024" : Return "32"
            Case "2025" : Return "33"
        End Select
        Return "30"
    End Function

    Private Shared Sub Log(tb As TextBox, msg As String)
        tb.AppendText(DateTime.Now.ToString("HH:mm:ss") & "  " & msg & vbCrLf)
    End Sub

End Class
