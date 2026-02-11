Option Strict On
Option Explicit On

Imports System
Imports System.Windows.Forms

Module LicenseGeneratorProgram

    <STAThread()>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New LicenseGeneratorForm())
    End Sub

End Module
