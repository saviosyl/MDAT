Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO

Public Class AboutForm
    Inherits Form

    Public Sub New()
        Me.Text = "About " & UITheme.AppName
        Me.Size = New Size(520, 400)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(15, 23, 42)

        ' Logo
        Dim picLogo As New PictureBox()
        picLogo.Size = New Size(240, 100)
        picLogo.SizeMode = PictureBoxSizeMode.Zoom
        picLogo.BackColor = Color.Transparent
        picLogo.Location = New Point((Me.ClientSize.Width - 240) \ 2, 20)
        Try
            Dim logoPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets\logo\logo.png")
            If File.Exists(logoPath) Then
                Using fs As New FileStream(logoPath, FileMode.Open, FileAccess.Read)
                    Using bmp As New Bitmap(fs)
                        picLogo.Image = New Bitmap(bmp)
                    End Using
                End Using
            End If
        Catch
        End Try
        Me.Controls.Add(picLogo)

        Dim lblTitle As New Label()
        lblTitle.Text = "MDAT"
        lblTitle.Font = New Font("Segoe UI", 20, FontStyle.Bold)
        lblTitle.ForeColor = Color.White
        lblTitle.BackColor = Color.Transparent
        lblTitle.AutoSize = True
        lblTitle.Location = New Point((Me.ClientSize.Width - 80) \ 2, 130)
        Me.Controls.Add(lblTitle)

        Dim lblSub As New Label()
        lblSub.Text = "Mechanical Design Automation Tool"
        lblSub.Font = New Font("Segoe UI", 10, FontStyle.Regular)
        lblSub.ForeColor = Color.FromArgb(169, 199, 232)
        lblSub.BackColor = Color.Transparent
        lblSub.AutoSize = True
        lblSub.Location = New Point((Me.ClientSize.Width - 260) \ 2, 165)
        Me.Controls.Add(lblSub)

        Dim lblVersion As New Label()
        lblVersion.Text = "Version 1.0"
        lblVersion.Font = New Font("Segoe UI", 9, FontStyle.Regular)
        lblVersion.ForeColor = Color.FromArgb(120, 140, 170)
        lblVersion.BackColor = Color.Transparent
        lblVersion.AutoSize = True
        lblVersion.Location = New Point((Me.ClientSize.Width - 80) \ 2, 195)
        Me.Controls.Add(lblVersion)

        Dim lblInfo As New Label()
        lblInfo.Text =
            "© 2026 MetaMech Solutions" & Environment.NewLine &
            "www.metamechsolutions.com" & Environment.NewLine &
            Environment.NewLine &
            "Designed by MetaMech Solutions"
        lblInfo.ForeColor = Color.FromArgb(169, 199, 232)
        lblInfo.BackColor = Color.Transparent
        lblInfo.Font = New Font("Segoe UI", 9, FontStyle.Regular)
        lblInfo.TextAlign = ContentAlignment.MiddleCenter
        lblInfo.Location = New Point(40, 225)
        lblInfo.Size = New Size(Me.ClientSize.Width - 80, 80)
        Me.Controls.Add(lblInfo)

        Dim btnOK As New Button()
        btnOK.Text = "OK"
        btnOK.Size = New Size(100, 34)
        btnOK.Location = New Point((Me.ClientSize.Width - 100) \ 2, 320)
        btnOK.BackColor = Color.FromArgb(0, 180, 255)
        btnOK.ForeColor = Color.White
        btnOK.FlatStyle = FlatStyle.Flat
        btnOK.FlatAppearance.BorderSize = 0
        btnOK.Font = New Font("Segoe UI Semibold", 10, FontStyle.Bold)
        btnOK.Cursor = Cursors.Hand
        AddHandler btnOK.Click, Sub() Me.Close()
        Me.Controls.Add(btnOK)
    End Sub

End Class
