Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Public Module FormHeader

    Public Sub AddPremiumHeader(frm As Form, formTitle As String, formSubtitle As String)
        Dim pnlHeader As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.FromArgb(15, 23, 42)
        }
        frm.Controls.Add(pnlHeader)

        ' Logo
        Dim picLogo As New PictureBox() With {
            .Size = New Size(140, 60),
            .Location = New Point(15, 10),
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
        pnlHeader.Controls.Add(picLogo)

        ' Title
        Dim lblTitle As New Label() With {
            .Text = formTitle,
            .Font = New Font("Segoe UI", 14.0F, FontStyle.Bold),
            .ForeColor = Color.White,
            .BackColor = Color.Transparent,
            .AutoSize = True,
            .Location = New Point(picLogo.Right + 12, 15)
        }
        pnlHeader.Controls.Add(lblTitle)

        ' Subtitle
        Dim lblSub As New Label() With {
            .Text = formSubtitle,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular),
            .ForeColor = Color.FromArgb(169, 199, 232),
            .BackColor = Color.Transparent,
            .AutoSize = True,
            .Location = New Point(picLogo.Right + 12, 45)
        }
        pnlHeader.Controls.Add(lblSub)

        ' Accent line
        Dim pnlLine As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 3,
            .BackColor = Color.FromArgb(0, 180, 255)
        }
        frm.Controls.Add(pnlLine)

        ' Make sure header is on top (dock order)
        pnlLine.BringToFront()
        pnlHeader.BringToFront()
    End Sub

End Module
