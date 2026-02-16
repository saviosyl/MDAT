Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Module FormHeader

    Public Sub AddPremiumHeader(frm As Form, formTitle As String, formSubtitle As String, Optional addFooter As Boolean = True)
        ' 1. Collect all existing controls from the form
        Dim existingControls As New List(Of Control)()
        For Each c As Control In frm.Controls
            existingControls.Add(c)
        Next

        ' 2. Remove them all
        frm.Controls.Clear()

        ' 3. Create content panel with AutoScroll, dock Fill
        Dim pnlContent As New Panel() With {
            .Dock = DockStyle.Fill,
            .AutoScroll = True,
            .Padding = New Padding(0, 5, 0, 0)
        }

        ' 4. Move all existing controls into content panel
        For Each c As Control In existingControls
            pnlContent.Controls.Add(c)
        Next

        ' 5. Add content panel to form FIRST (dock fill must be added before dock top/bottom)
        frm.Controls.Add(pnlContent)

        ' 6. Add footer (dock bottom) — before header so dock order is correct
        If addFooter Then
            Dim isDark As Boolean = UITheme.IsDark

            Dim pnlFooter As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 28
            }
            frm.Controls.Add(pnlFooter)

            Dim pnlFBorder As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 1,
                .BackColor = Color.FromArgb(60, 0, 180, 255)
            }
            pnlFooter.Controls.Add(pnlFBorder)

            pnlFooter.BackColor = If(isDark, Color.FromArgb(10, 25, 41), Color.FromArgb(240, 242, 245))
            Dim textCol As Color = If(isDark, Color.FromArgb(169, 199, 232), Color.FromArgb(80, 90, 100))
            Dim accentCol As Color = Color.FromArgb(0, 180, 255)

            Dim lblCopy As New Label() With {
                .Text = "© 2026 MetaMech Solutions",
                .Font = New Font("Segoe UI", 7.5F, FontStyle.Regular),
                .ForeColor = textCol,
                .BackColor = Color.Transparent,
                .AutoSize = True,
                .Location = New Point(10, 7)
            }
            pnlFooter.Controls.Add(lblCopy)

            Dim lnkSite As New LinkLabel() With {
                .Text = "metamechsolutions.com",
                .Font = New Font("Segoe UI", 7.5F, FontStyle.Regular),
                .LinkColor = accentCol,
                .ActiveLinkColor = accentCol,
                .VisitedLinkColor = accentCol,
                .BackColor = Color.Transparent,
                .AutoSize = True
            }
            AddHandler lnkSite.LinkClicked, Sub(s As Object, ev As LinkLabelLinkClickedEventArgs)
                                                 Try
                                                     System.Diagnostics.Process.Start("https://metamechsolutions.com/")
                                                 Catch
                                                 End Try
                                             End Sub
            pnlFooter.Controls.Add(lnkSite)

            Dim lnkBlog As New LinkLabel() With {
                .Text = "Help & Blog",
                .Font = New Font("Segoe UI", 7.5F, FontStyle.Regular),
                .LinkColor = accentCol,
                .ActiveLinkColor = accentCol,
                .VisitedLinkColor = accentCol,
                .BackColor = Color.Transparent,
                .AutoSize = True,
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right
            }
            AddHandler lnkBlog.LinkClicked, Sub(s As Object, ev As LinkLabelLinkClickedEventArgs)
                                                 Try
                                                     System.Diagnostics.Process.Start("https://metamechsolutions.com/blog/")
                                                 Catch
                                                 End Try
                                             End Sub
            pnlFooter.Controls.Add(lnkBlog)

            AddHandler pnlFooter.Resize, Sub(s As Object, ev As EventArgs)
                                              Try
                                                  lnkSite.Location = New Point((pnlFooter.Width - lnkSite.Width) \ 2, 7)
                                                  lnkBlog.Location = New Point(pnlFooter.Width - lnkBlog.Width - 10, 7)
                                              Catch
                                              End Try
                                          End Sub
            Try
                lnkSite.Location = New Point((pnlFooter.Width - lnkSite.Width) \ 2, 7)
                lnkBlog.Location = New Point(pnlFooter.Width - lnkBlog.Width - 10, 7)
            Catch
            End Try
        End If

        ' 7. Add accent line (dock top)
        Dim pnlLine As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 3,
            .BackColor = Color.FromArgb(0, 180, 255)
        }
        frm.Controls.Add(pnlLine)

        ' 8. Add header panel (dock top) — added last so it appears at the very top
        Dim pnlHeader As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.FromArgb(15, 23, 42)
        }
        frm.Controls.Add(pnlHeader)

        ' Set form icon
        Try
            Dim icoPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "mdat.ico")
            If File.Exists(icoPath) Then
                frm.Icon = New Icon(icoPath)
            End If
        Catch
        End Try

        ' Logo
        Dim picLogo As New PictureBox() With {
            .Size = New Size(240, 65),
            .Location = New Point(10, 8),
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
            .Location = New Point(picLogo.Right + 10, 12)
        }
        pnlHeader.Controls.Add(lblTitle)

        ' Subtitle
        Dim lblSub As New Label() With {
            .Text = formSubtitle,
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Regular),
            .ForeColor = Color.FromArgb(169, 199, 232),
            .BackColor = Color.Transparent,
            .AutoSize = True,
            .Location = New Point(picLogo.Right + 10, 42)
        }
        pnlHeader.Controls.Add(lblSub)
    End Sub

End Module
