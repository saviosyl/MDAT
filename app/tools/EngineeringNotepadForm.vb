Option Strict On

Imports System.Drawing
Imports System.Windows.Forms

Public Class EngineeringNotepadForm
    Inherits Form

    Private txtEditor As TextBox
    Private lblStatus As Label

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()

        Me.Text = "Engineering Notepad"
        Me.Size = New Size(900, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = UITheme.BG_MAIN

        '================ HEADER =================
        Dim lblTitle As New Label()
        lblTitle.Text = "Engineering Notepad"
        lblTitle.Font = New Font("Segoe UI", 16, FontStyle.Bold)
        lblTitle.ForeColor = UITheme.TEXT_PRIMARY
        lblTitle.Location = New Point(20, 15)
        lblTitle.AutoSize = True
        Me.Controls.Add(lblTitle)

        Dim lblSub As New Label()
        lblSub.Text = "Notes • Calculations • Ideas"
        lblSub.Font = New Font("Segoe UI", 9)
        lblSub.ForeColor = UITheme.TEXT_MUTED
        lblSub.Location = New Point(22, 48)
        lblSub.AutoSize = True
        Me.Controls.Add(lblSub)

        Dim accent As New Panel()
        accent.BackColor = UITheme.ACCENT_PRIMARY
        accent.Location = New Point(0, 80)
        accent.Size = New Size(Me.Width, 2)
        Me.Controls.Add(accent)

        '================ TOOLBAR =================
        Dim pnlToolbar As New Panel()
        pnlToolbar.Location = New Point(0, 82)
        pnlToolbar.Size = New Size(Me.Width, 40)
        pnlToolbar.BackColor = UITheme.BG_PANEL
        Me.Controls.Add(pnlToolbar)

        AddToolbarButton(pnlToolbar, "Ø", 20)
        AddToolbarButton(pnlToolbar, "±", 70)
        AddToolbarButton(pnlToolbar, "≤", 120)
        AddToolbarButton(pnlToolbar, "≥", 170)
        AddToolbarButton(pnlToolbar, "=", 240)

        '================ EDITOR =================
        txtEditor = New TextBox()
        txtEditor.Multiline = True
        txtEditor.ScrollBars = ScrollBars.Both
        txtEditor.WordWrap = False
        txtEditor.Font = New Font("Consolas", 11)
        txtEditor.ForeColor = UITheme.TEXT_PRIMARY
        txtEditor.BackColor = UITheme.BG_CARD
        txtEditor.BorderStyle = BorderStyle.FixedSingle
        txtEditor.Location = New Point(20, 140)
        txtEditor.Size = New Size(Me.ClientSize.Width - 40, Me.ClientSize.Height - 200)
        txtEditor.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        txtEditor.Text = "Type engineering notes or quick calculations here…" & Environment.NewLine
        Me.Controls.Add(txtEditor)

        '================ STATUS =================
        lblStatus = New Label()
        lblStatus.Text = "Ready"
        lblStatus.Dock = DockStyle.Bottom
        lblStatus.Height = 22
        lblStatus.TextAlign = ContentAlignment.MiddleLeft
        lblStatus.ForeColor = UITheme.TEXT_MUTED
        lblStatus.BackColor = UITheme.BG_PANEL
        Me.Controls.Add(lblStatus)

    End Sub

    Private Sub AddToolbarButton(parent As Panel, text As String, x As Integer)
        Dim b As New Button()
        b.Text = text
        b.Size = New Size(40, 26)
        b.Location = New Point(x, 7)
        b.BackColor = UITheme.BTN_BG
        b.ForeColor = UITheme.TEXT_PRIMARY
        b.FlatStyle = FlatStyle.Flat
        b.FlatAppearance.BorderSize = 0
        parent.Controls.Add(b)
    End Sub

End Class
