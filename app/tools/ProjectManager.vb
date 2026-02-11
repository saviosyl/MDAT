Imports System.Windows.Forms
Imports System.IO

Public Module ProjectManager

    Public CurrentProject As String = "Default"

    Public Function GetProjectFolder() As String
        Dim baseDir As String = Path.Combine(Application.StartupPath, "Projects")
        If Not Directory.Exists(baseDir) Then Directory.CreateDirectory(baseDir)

        Dim p As String = Path.Combine(baseDir, CurrentProject)
        If Not Directory.Exists(p) Then Directory.CreateDirectory(p)

        Return p
    End Function

End Module
