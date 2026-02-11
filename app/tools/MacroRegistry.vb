Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.IO

' ============================================================
' MacroRegistry
' ------------------------------------------------------------
' Acts as the SINGLE source of truth for macro slots
' Populated from MainForm.LoadMacrosFromConfig()
' Used by MacroEngine only
' ============================================================
Public Module MacroRegistry

    Private ReadOnly _slots As New Dictionary(Of Integer, MacroSlot)

    ' ------------------------------------------------------------
    ' Register macro slot from existing config parsing
    ' execRaw format:
    '   path.swp
    '   path.swp | Module
    '   path.swp | Module | Proc
    ' ------------------------------------------------------------
    Public Sub Register(slotId As Integer, execRaw As String)

        If execRaw Is Nothing Then Exit Sub

        Dim raw As String = execRaw.Trim()
        If raw = "" Then Exit Sub

        Dim path As String = ""
        Dim moduleName As String = ""
        Dim procName As String = ""

        Dim parts() As String = raw.Split("|"c)

        If parts.Length >= 1 Then path = parts(0).Trim()
        If parts.Length >= 2 Then moduleName = parts(1).Trim()
        If parts.Length >= 3 Then procName = parts(2).Trim()

        If path = "" Then Exit Sub

        _slots(slotId) = New MacroSlot With {
            .SlotId = slotId,
            .MacroPath = path,
            .ModuleName = moduleName,
            .ProcName = procName
        }

    End Sub

    ' ------------------------------------------------------------
    ' Retrieve macro slot (used by MacroEngine)
    ' ------------------------------------------------------------
    Public Function GetMacro(slotId As Integer) As MacroSlot
        If _slots.ContainsKey(slotId) Then
            Return _slots(slotId)
        End If
        Return Nothing
    End Function

    ' ------------------------------------------------------------
    ' Optional: Clear registry (future hot-reload support)
    ' ------------------------------------------------------------
    Public Sub Clear()
        _slots.Clear()
    End Sub

End Module
