Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports System.Diagnostics

Public Class AboutForm
    Inherits Form

    Public Sub New()
        Me.Text = "About " & UITheme.AppName
        Me.Size = New Size(520, 550)
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
            Environment.NewLine &
            "Designed by MetaMech Solutions"
        lblInfo.ForeColor = Color.FromArgb(169, 199, 232)
        lblInfo.BackColor = Color.Transparent
        lblInfo.Font = New Font("Segoe UI", 9, FontStyle.Regular)
        lblInfo.TextAlign = ContentAlignment.MiddleCenter
        lblInfo.Location = New Point(40, 225)
        lblInfo.Size = New Size(Me.ClientSize.Width - 80, 60)
        Me.Controls.Add(lblInfo)

        Dim accentCol As Color = Color.FromArgb(0, 180, 255)

        Dim lnkSite As New LinkLabel()
        lnkSite.Text = "metamechsolutions.com"
        lnkSite.Font = New Font("Segoe UI", 9, FontStyle.Regular)
        lnkSite.LinkColor = accentCol
        lnkSite.ActiveLinkColor = accentCol
        lnkSite.VisitedLinkColor = accentCol
        lnkSite.BackColor = Color.Transparent
        lnkSite.AutoSize = True
        lnkSite.Location = New Point((Me.ClientSize.Width - 160) \ 2, 285)
        AddHandler lnkSite.LinkClicked, Sub(ss As Object, ee As LinkLabelLinkClickedEventArgs)
                                             Try
                                                 Process.Start("https://metamechsolutions.com/")
                                             Catch
                                             End Try
                                         End Sub
        Me.Controls.Add(lnkSite)

        Dim lnkBlog As New LinkLabel()
        lnkBlog.Text = "Help & Blog"
        lnkBlog.Font = New Font("Segoe UI", 9, FontStyle.Regular)
        lnkBlog.LinkColor = accentCol
        lnkBlog.ActiveLinkColor = accentCol
        lnkBlog.VisitedLinkColor = accentCol
        lnkBlog.BackColor = Color.Transparent
        lnkBlog.AutoSize = True
        lnkBlog.Location = New Point((Me.ClientSize.Width - 80) \ 2, 305)
        AddHandler lnkBlog.LinkClicked, Sub(ss As Object, ee As LinkLabelLinkClickedEventArgs)
                                             Try
                                                 Process.Start("https://metamechsolutions.com/blog/")
                                             Catch
                                             End Try
                                         End Sub
        Me.Controls.Add(lnkBlog)

        Dim btnOK As New Button()
        btnOK.Text = "OK"
        btnOK.Size = New Size(100, 34)
        btnOK.Location = New Point((Me.ClientSize.Width - 100) \ 2, 340)
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
