Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Globalization
Imports System.IO

Public Class LicenseGeneratorForm
    Inherits Form

    ' ================= THEME =================
    Private ReadOnly BG As Color = Color.FromArgb(18, 22, 30)
    Private ReadOnly PANEL As Color = Color.FromArgb(28, 34, 44)
    Private ReadOnly PURPLE As Color = Color.FromArgb(150, 90, 190)
    Private ReadOnly GREEN As Color = Color.LightGreen
    Private ReadOnly YELLOW As Color = Color.Gold
    Private ReadOnly TXT As Color = Color.Gainsboro
    Private ReadOnly MUTED As Color = Color.FromArgb(170, 170, 170)
    Private ReadOnly ERR As Color = Color.OrangeRed

    ' ================= UI =================
    Private pnl As Panel

    Private lblTitle As Label
    Private lblSub As Label

    Private lblTier As Label
    Private cmbTier As ComboBox

    Private lblSeats As Label
    Private nudSeats As NumericUpDown

    Private lblExpiry As Label
    Private dtExpiry As DateTimePicker
    Private lblTrialNote As Label

    Private lblId As Label
    Private txtId As TextBox
    Private btnNewId As Button

    Private btnGenerate As Button
    Private btnCopy As Button
    Private btnSave As Button

    Private txtOut As TextBox
    Private lblStatus As Label

    Public Sub New()
        Me.Text = "MetaMech License Generator (Internal)"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Size = New Size(760, 560)
        Me.BackColor = BG

        BuildUI()
        WireEvents()

        ' Defaults
        cmbTier.SelectedIndex = 0 ' Trial
        nudSeats.Value = 3
        SetNewLicenseId()
        RefreshTierUI()
    End Sub

    Private Sub BuildUI()

        pnl = New Panel()
        pnl.Location = New Point(12, 12)
        pnl.Size = New Size(720, 500)
        pnl.BackColor = PANEL
        Me.Controls.Add(pnl)

        lblTitle = New Label()
        lblTitle.Text = "MetaMech License Generator"
        lblTitle.Font = New Font("Segoe UI", 14.0F, FontStyle.Bold)
        lblTitle.ForeColor = Color.White
        lblTitle.Location = New Point(18, 14)
        lblTitle.AutoSize = True
        pnl.Controls.Add(lblTitle)

        lblSub = New Label()
        lblSub.Text = "Generates RSA SHA-256 license: LICENSEID|TIER|EXPIRY_UTC|SEATS"
        lblSub.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        lblSub.ForeColor = MUTED
        lblSub.Location = New Point(20, 44)
        lblSub.AutoSize = True
        pnl.Controls.Add(lblSub)

        Dim y As Integer = 80

        ' Tier
        lblTier = MakeLabel("Tier", 20, y)
        pnl.Controls.Add(lblTier)

        cmbTier = New ComboBox()
        cmbTier.Location = New Point(170, y - 2)
        cmbTier.Size = New Size(260, 24)
        cmbTier.DropDownStyle = ComboBoxStyle.DropDownList
        cmbTier.BackColor = BG
        cmbTier.ForeColor = TXT
        cmbTier.Items.Add("0 - Trial (3 days)")
        cmbTier.Items.Add("1 - Standard")
        cmbTier.Items.Add("2 - Premium")
        cmbTier.Items.Add("3 - Premium Plus")
        pnl.Controls.Add(cmbTier)

        y += 34

        ' Seats
        lblSeats = MakeLabel("Seats", 20, y)
        pnl.Controls.Add(lblSeats)

        nudSeats = New NumericUpDown()
        nudSeats.Location = New Point(170, y - 2)
        nudSeats.Size = New Size(100, 24)
        nudSeats.Minimum = 1D
        nudSeats.Maximum = 999D
        nudSeats.BackColor = BG
        nudSeats.ForeColor = TXT
        nudSeats.Value = 3D
        pnl.Controls.Add(nudSeats)

        Dim lblSeatsHint As Label = New Label()
        lblSeatsHint.Text = "Number of allowed seats (stored in payload + signed)."
        lblSeatsHint.ForeColor = MUTED
        lblSeatsHint.Location = New Point(280, y)
        lblSeatsHint.AutoSize = True
        pnl.Controls.Add(lblSeatsHint)

        y += 34

        ' Expiry
        lblExpiry = MakeLabel("Expiry (UTC)", 20, y)
        pnl.Controls.Add(lblExpiry)

        dtExpiry = New DateTimePicker()
        dtExpiry.Location = New Point(170, y - 2)
        dtExpiry.Size = New Size(260, 24)
        dtExpiry.Format = DateTimePickerFormat.Custom
        dtExpiry.CustomFormat = "yyyy-MM-dd  HH:mm"
        dtExpiry.CalendarMonthBackground = BG
        dtExpiry.CalendarForeColor = TXT
        pnl.Controls.Add(dtExpiry)

        lblTrialNote = New Label()
        lblTrialNote.Text = "Trial auto-sets expiry = now + 3 days (UTC)"
        lblTrialNote.ForeColor = MUTED
        lblTrialNote.Location = New Point(170, y + 26)
        lblTrialNote.AutoSize = True
        pnl.Controls.Add(lblTrialNote)

        y += 60

        ' License ID
        lblId = MakeLabel("License ID", 20, y)
        pnl.Controls.Add(lblId)

        txtId = New TextBox()
        txtId.Location = New Point(170, y - 2)
        txtId.Size = New Size(380, 24)
        txtId.BackColor = BG
        txtId.ForeColor = TXT
        pnl.Controls.Add(txtId)

        btnNewId = New Button()
        btnNewId.Text = "New"
        btnNewId.Location = New Point(560, y - 3)
        btnNewId.Size = New Size(60, 26)
        btnNewId.BackColor = YELLOW
        btnNewId.ForeColor = Color.Black
        btnNewId.FlatStyle = FlatStyle.Flat
        btnNewId.FlatAppearance.BorderSize = 0
        pnl.Controls.Add(btnNewId)

        y += 44

        ' Buttons
        btnGenerate = New Button()
        btnGenerate.Text = "Generate License"
        btnGenerate.Location = New Point(170, y)
        btnGenerate.Size = New Size(160, 34)
        btnGenerate.BackColor = PURPLE
        btnGenerate.ForeColor = Color.White
        btnGenerate.FlatStyle = FlatStyle.Flat
        btnGenerate.FlatAppearance.BorderSize = 0
        pnl.Controls.Add(btnGenerate)

        btnCopy = New Button()
        btnCopy.Text = "Copy"
        btnCopy.Location = New Point(340, y)
        btnCopy.Size = New Size(90, 34)
        btnCopy.BackColor = GREEN
        btnCopy.ForeColor = Color.Black
        btnCopy.FlatStyle = FlatStyle.Flat
        btnCopy.FlatAppearance.BorderSize = 0
        pnl.Controls.Add(btnCopy)

        btnSave = New Button()
        btnSave.Text = "Save license.key"
        btnSave.Location = New Point(440, y)
        btnSave.Size = New Size(150, 34)
        btnSave.BackColor = Color.FromArgb(0, 170, 210)
        btnSave.ForeColor = Color.Black
        btnSave.FlatStyle = FlatStyle.Flat
        btnSave.FlatAppearance.BorderSize = 0
        pnl.Controls.Add(btnSave)

        y += 48

        ' Output
        txtOut = New TextBox()
        txtOut.Location = New Point(20, y)
        txtOut.Size = New Size(680, 160)
        txtOut.Multiline = True
        txtOut.ScrollBars = ScrollBars.Vertical
        txtOut.ReadOnly = True
        txtOut.BackColor = BG
        txtOut.ForeColor = TXT
        pnl.Controls.Add(txtOut)

        y += 170

        lblStatus = New Label()
        lblStatus.Text = "Status: Ready"
        lblStatus.ForeColor = MUTED
        lblStatus.Location = New Point(20, y)
        lblStatus.AutoSize = True
        pnl.Controls.Add(lblStatus)

    End Sub

    Private Function MakeLabel(ByVal text As String, ByVal x As Integer, ByVal y As Integer) As Label
        Dim l As New Label()
        l.Text = text
        l.ForeColor = TXT
        l.Location = New Point(x, y)
        l.AutoSize = True
        Return l
    End Function

    Private Sub WireEvents()
        AddHandler cmbTier.SelectedIndexChanged, AddressOf TierChanged
        AddHandler btnNewId.Click, AddressOf NewIdClicked
        AddHandler btnGenerate.Click, AddressOf GenerateClicked
        AddHandler btnCopy.Click, AddressOf CopyClicked
        AddHandler btnSave.Click, AddressOf SaveClicked
    End Sub

    Private Sub TierChanged(ByVal sender As Object, ByVal e As EventArgs)
        RefreshTierUI()
    End Sub

    Private Sub RefreshTierUI()
        Dim tier As Integer = GetSelectedTier()

        If tier = 0 Then
            ' Trial: expiry auto now+3 days UTC
            dtExpiry.Enabled = False
            lblTrialNote.Visible = True
        Else
            dtExpiry.Enabled = True
            lblTrialNote.Visible = False
        End If

        ' Auto set a sensible default expiry for paid tiers (1 year)
        If tier <> 0 Then
            Dim nowUtc As DateTime = DateTime.UtcNow
            dtExpiry.Value = nowUtc.AddYears(1).ToLocalTime()
        Else
            dtExpiry.Value = DateTime.UtcNow.AddDays(LicenseGeneratorCore.TRIAL_DAYS).ToLocalTime()
        End If
    End Sub

    Private Sub NewIdClicked(ByVal sender As Object, ByVal e As EventArgs)
        SetNewLicenseId()
        lblStatus.Text = "Status: New License ID generated"
        lblStatus.ForeColor = MUTED
    End Sub

    Private Sub SetNewLicenseId()
        txtId.Text = Guid.NewGuid().ToString("N")
    End Sub

    Private Sub GenerateClicked(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim tier As Integer = GetSelectedTier()
            Dim seats As Integer = CInt(nudSeats.Value)

            Dim licenseId As String = txtId.Text.Trim()
            If licenseId = "" Then
                Throw New ApplicationException("License ID cannot be blank.")
            End If

            Dim expiryUtc As DateTime

            If tier = 0 Then
                expiryUtc = DateTime.UtcNow.AddDays(LicenseGeneratorCore.TRIAL_DAYS)
            Else
                ' User picks local date/time; convert to UTC
                Dim localVal As DateTime = dtExpiry.Value
                expiryUtc = localVal.ToUniversalTime()
            End If

            Dim lic As String = LicenseGeneratorCore.GenerateLicense(licenseId, tier, expiryUtc, seats)

            txtOut.Text = lic
            lblStatus.Text = "Status: License generated (Tier " & tier.ToString(CultureInfo.InvariantCulture) &
                             ", Seats " & seats.ToString(CultureInfo.InvariantCulture) & ")"
            lblStatus.ForeColor = GREEN

        Catch ex As Exception
            lblStatus.Text = "Status: ERROR - " & ex.Message
            lblStatus.ForeColor = ERR
        End Try
    End Sub

    Private Sub CopyClicked(ByVal sender As Object, ByVal e As EventArgs)
        Try
            If txtOut.Text.Trim() = "" Then Exit Sub
            Clipboard.SetText(txtOut.Text.Trim())
            lblStatus.Text = "Status: Copied to clipboard"
            lblStatus.ForeColor = GREEN
        Catch ex As Exception
            lblStatus.Text = "Status: ERROR - " & ex.Message
            lblStatus.ForeColor = ERR
        End Try
    End Sub

    Private Sub SaveClicked(ByVal sender As Object, ByVal e As EventArgs)
        Try
            If txtOut.Text.Trim() = "" Then
                Throw New ApplicationException("Nothing to save. Generate license first.")
            End If

            Dim sfd As New SaveFileDialog()
            sfd.Filter = "License Key|license.key|All Files|*.*"
            sfd.FileName = "license.key"

            If sfd.ShowDialog() <> DialogResult.OK Then Exit Sub

            File.WriteAllText(sfd.FileName, txtOut.Text.Trim())

            lblStatus.Text = "Status: Saved to " & sfd.FileName
            lblStatus.ForeColor = GREEN

        Catch ex As Exception
            lblStatus.Text = "Status: ERROR - " & ex.Message
            lblStatus.ForeColor = ERR
        End Try
    End Sub

    Private Function GetSelectedTier() As Integer
        If cmbTier.SelectedIndex < 0 Then Return 0
        Select Case cmbTier.SelectedIndex
            Case 0 : Return 0
            Case 1 : Return 1
            Case 2 : Return 2
            Case 3 : Return 3
        End Select
        Return 0
    End Function

End Class
