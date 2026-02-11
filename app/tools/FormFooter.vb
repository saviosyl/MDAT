Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Diagnostics

Public Module FormFooter

    Public Sub AddPremiumFooter(frm As Form)
        Dim isDark As Boolean = UITheme.IsDark

        Dim pnlFooter As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 28
        }
        frm.Controls.Add(pnlFooter)

        ' 1px accent top border
        Dim pnlBorder As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 1,
            .BackColor = Color.FromArgb(60, 0, 180, 255)
        }
        pnlFooter.Controls.Add(pnlBorder)

        ' Background
        pnlFooter.BackColor = If(isDark, Color.FromArgb(10, 25, 41), Color.FromArgb(240, 242, 245))

        Dim textCol As Color = If(isDark, Color.FromArgb(169, 199, 232), Color.FromArgb(80, 90, 100))
        Dim accentCol As Color = Color.FromArgb(0, 180, 255)

        ' Left: copyright
        Dim lblCopy As New Label() With {
            .Text = "© 2026 MetaMech Solutions",
            .Font = New Font("Segoe UI", 7.5F, FontStyle.Regular),
            .ForeColor = textCol,
            .BackColor = Color.Transparent,
            .AutoSize = True,
            .Location = New Point(10, 7)
        }
        pnlFooter.Controls.Add(lblCopy)

        ' Center: website link
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
                                                 Process.Start("https://metamechsolutions.com/")
                                             Catch
                                             End Try
                                         End Sub
        pnlFooter.Controls.Add(lnkSite)

        ' Right: blog link
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
                                                 Process.Start("https://metamechsolutions.com/blog/")
                                             Catch
                                             End Try
                                         End Sub
        pnlFooter.Controls.Add(lnkBlog)

        ' Position center and right labels on resize
        AddHandler pnlFooter.Resize, Sub(s As Object, ev As EventArgs)
                                          Try
                                              lnkSite.Location = New Point((pnlFooter.Width - lnkSite.Width) \ 2, 7)
                                              lnkBlog.Location = New Point(pnlFooter.Width - lnkBlog.Width - 10, 7)
                                          Catch
                                          End Try
                                      End Sub

        ' Initial position
        Try
            lnkSite.Location = New Point((pnlFooter.Width - lnkSite.Width) \ 2, 7)
            lnkBlog.Location = New Point(pnlFooter.Width - lnkBlog.Width - 10, 7)
        Catch
        End Try

        ' Footer should be at bottom, bring to front so dock order is correct
        pnlFooter.BringToFront()
    End Sub

End Module
