Option Strict On
Option Explicit On

Imports System
Imports System.Windows.Forms

Module Program

    <STAThread()>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        Dim splash As New SplashForm()
        splash.Show()
        Application.DoEvents()

        Dim mainForm As New MainForm()
        splash.Close()
        splash.Dispose()

        Application.Run(mainForm)
    End Sub

End Module
