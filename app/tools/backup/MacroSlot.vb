' ==============================
' MacroSlot.vb
' ==============================
Public Class MacroSlot
    Public Path As String
    Public ModuleName As String
    Public ProcName As String

    Public Sub New(p As String, m As String, r As String)
        Path = If(p, "").Trim()
        ModuleName = If(m, "").Trim()
        ProcName = If(r, "").Trim()
    End Sub
End Class
