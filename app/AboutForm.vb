Option Strict On

Imports System.Drawing
Imports System.Windows.Forms

Public Class AboutForm
    Inherits Form

    Public Sub New()
        Me.Text = "About " & UITheme.AppName
        Me.Size = New Size(420, 260)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = UITheme.BG_MAIN

        Dim lblTitle As New Label()
        lblTitle.Text = UITheme.AppName
        lblTitle.Font = New Font("Segoe UI", 16, FontStyle.Bold)
        lblTitle.ForeColor = UITheme.TEXT_PRIMARY
        lblTitle.Location = New Point(20, 20)
        lblTitle.AutoSize = True

        Dim lblSub As New Label()
        lblSub.Text = UITheme.AppTagline
        lblSub.ForeColor = UITheme.TEXT_MUTED
        lblSub.Location = New Point(22, 55)
        lblSub.AutoSize = True

        Dim lblInfo As New Label()
        lblInfo.Text =
            "MDAT is a Mechanical Design Automation Tool." & Environment.NewLine &
            "This UI layer does not affect macro execution," & Environment.NewLine &
            "SolidWorks integration, or license enforcement." & Environment.NewLine &
            Environment.NewLine &
            "Powered by MetaMech Solutions"
        lblInfo.ForeColor = UITheme.TEXT_MUTED
        lblInfo.Location = New Point(22, 90)
        lblInfo.Size = New Size(360, 100)

        Dim btnOK As New Button()
        btnOK.Text = "OK"
        btnOK.Size = New Size(80, 28)
        btnOK.Location = New Point(300, 190)
        btnOK.BackColor = UITheme.BTN_BG
        btnOK.ForeColor = UITheme.TEXT_PRIMARY
        btnOK.FlatStyle = FlatStyle.Flat
        btnOK.FlatAppearance.BorderSize = 0
        AddHandler btnOK.Click, Sub() Me.Close()

        Me.Controls.Add(lblTitle)
        Me.Controls.Add(lblSub)
        Me.Controls.Add(lblInfo)
        Me.Controls.Add(btnOK)
    End Sub

End Class
