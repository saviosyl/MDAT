Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Public Class SplashForm
    Inherits Form

    Public Sub New()
        Me.FormBorderStyle = FormBorderStyle.None
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Size = New Size(500, 300)
        Me.BackColor = Color.FromArgb(15, 23, 42)
        Me.ShowInTaskbar = False
        Me.TopMost = True

        ' Logo — centered, big
        Dim picLogo As New PictureBox() With {
            .Size = New Size(280, 120),
            .Location = New Point((500 - 280) \ 2, 30),
            .SizeMode = PictureBoxSizeMode.Zoom,
            .BackColor = Color.Transparent
        }
        Dim logoPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "logo", "logo.png")
        Try
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

        ' App name
        Dim lblName As New Label() With {
            .Text = "MDAT",
            .Font = New Font("Segoe UI", 24, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.Transparent,
            .AutoSize = True
        }
        lblName.Location = New Point((500 - 80) \ 2, 160)
        Me.Controls.Add(lblName)

        ' Tagline
        Dim lblTag As New Label() With {
            .Text = "Mechanical Design Automation Tool",
            .Font = New Font("Segoe UI", 10, FontStyle.Regular),
            .ForeColor = Color.FromArgb(169, 199, 232),
            .BackColor = Color.Transparent,
            .AutoSize = True
        }
        lblTag.Location = New Point((500 - 260) \ 2, 200)
        Me.Controls.Add(lblTag)

        ' Loading text
        Dim lblLoading As New Label() With {
            .Text = "Loading...",
            .Font = New Font("Segoe UI", 9, FontStyle.Italic),
            .ForeColor = Color.FromArgb(0, 180, 255),
            .BackColor = Color.Transparent,
            .AutoSize = True
        }
        lblLoading.Location = New Point((500 - 70) \ 2, 240)
        Me.Controls.Add(lblLoading)

        ' Version
        Dim lblVer As New Label() With {
            .Text = "v1.0",
            .Font = New Font("Segoe UI", 8, FontStyle.Regular),
            .ForeColor = Color.FromArgb(100, 120, 150),
            .BackColor = Color.Transparent,
            .AutoSize = True,
            .Location = New Point((500 - 30) \ 2, 270)
        }
        Me.Controls.Add(lblVer)
    End Sub
End Class
