Imports System
Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.Xml
Imports System.Reflection

Public Class MainForm
    Inherits Form

    '================ APP VERSION =================
    Private Const APP_VERSION As String = "1.0.0"
    Private Const MIN_SUPPORTED_VERSION As String = "1.0.0"

    '================ THEME =================
    Private ReadOnly MM_BG As Color = Color.FromArgb(14, 18, 26)
    Private ReadOnly MM_PANEL As Color = Color.FromArgb(24, 30, 40)
    Private ReadOnly MM_ACCENT As Color = Color.FromArgb(0, 170, 210)
    Private ReadOnly MM_TEXT As Color = Color.White
    Private ReadOnly MM_SUB As Color = Color.Gainsboro
    Private ReadOnly MM_GUIDE As Color = Color.Gold

    '================ UI =================
    Private pnlLeft As Panel
    Private pnlEng As Panel
    Private txtLog As TextBox
    Private cmbSW As ComboBox
    Private btnSelect As Button
    Private btnAdmin As Button
    Private btnSupport As Button
    Private picLogo As PictureBox

    ' Header labels
    Private lblVersion As Label
    Private lblLicense As Label

    '================ DATA =================
    Private MacroNames(10) As String
    Private MacroFiles(10) As String
    Private MacroModules(10) As String
    Private MacroMethods(10) As String
    Private MacroButtons(10) As Button
    Private GuideButtons(10) As Button

    Private selectedAssembly As String = ""
    Private isRunning As Boolean = False

    Private Const CONFIG_FILE As String = "admin_macros.xml"
    Private Const ENG_WIDTH As Integer = 220

    '====================================================
    ' FORM INIT (UI LOCKED)
    '====================================================
    Public Sub New()

        Me.Text = "MetaMech Mechanical Design Automation"
        Me.Size = New Size(1350, 720)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = MM_BG
        Me.ForeColor = MM_TEXT
        Me.Font = New Font("Segoe UI", 10)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        ' -------- LOGO (TRUE LEFT EDGE) --------
        picLogo = New PictureBox()
        picLogo.Size = New Size(320, 130)
        picLogo.Location = New Point(-70, 10)
        picLogo.SizeMode = PictureBoxSizeMode.Zoom
        picLogo.BackColor = Color.Transparent
        LoadLogo()
        Me.Controls.Add(picLogo)
        picLogo.BringToFront()

        ' -------- TITLE --------
        Dim lblTitle As New Label()
        lblTitle.Text = "Mechanical Design Automation"
        lblTitle.Font = New Font("Segoe UI", 20, FontStyle.Bold)
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(170, 25)
        Me.Controls.Add(lblTitle)
        lblTitle.BringToFront()

        Dim lblBrand As New Label()
        lblBrand.Text = "@ Designed by MetaMech Solutions"
        lblBrand.Font = New Font("Segoe UI", 10, FontStyle.Italic)
        lblBrand.ForeColor = MM_SUB
        lblBrand.AutoSize = True
        lblBrand.Location = New Point(170, 65)
        Me.Controls.Add(lblBrand)
        lblBrand.BringToFront()

        ' -------- VERSION / LICENSE --------
        lblVersion = New Label()
        lblVersion.Text = "Version: " & APP_VERSION
        lblVersion.Font = New Font("Segoe UI", 9)
        lblVersion.ForeColor = MM_SUB
        lblVersion.AutoSize = True
        lblVersion.Location = New Point(Me.Width - 170, 70)
        lblVersion.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Me.Controls.Add(lblVersion)

        lblLicense = New Label()
        lblLicense.Text = "License: Checking..."
        lblLicense.Font = New Font("Segoe UI", 9)
        lblLicense.ForeColor = MM_SUB
        lblLicense.AutoSize = True
        lblLicense.Location = New Point(Me.Width - 230, 90)
        lblLicense.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Me.Controls.Add(lblLicense)

        ' -------- ADMIN / SUPPORT --------
        btnAdmin = CreateHeaderButton("ADMIN", Me.Width - 220)
        AddHandler btnAdmin.Click, AddressOf AdminLogin
        btnSupport = CreateHeaderButton("SUPPORT", Me.Width - 120)
        Me.Controls.Add(btnAdmin)
        Me.Controls.Add(btnSupport)

        ' -------- LEFT MACROS --------
        pnlLeft = New Panel()
        pnlLeft.Size = New Size(280, 520)
        pnlLeft.Location = New Point(20, 160)
        pnlLeft.BackColor = MM_PANEL
        pnlLeft.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(pnlLeft)

        ' -------- ENGINEERING TOOLS --------
        pnlEng = New Panel()
        pnlEng.Size = New Size(ENG_WIDTH, 520)
        pnlEng.Location = New Point(Me.Width - ENG_WIDTH - 40, 160)
        pnlEng.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        pnlEng.BackColor = MM_PANEL
        pnlEng.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(pnlEng)
        BuildEngineeringTools()

        ' -------- SOLIDWORKS --------
        cmbSW = New ComboBox()
        cmbSW.Items.AddRange(New String() {"2022", "2023", "2024", "2025"})
        cmbSW.SelectedIndex = 3
        cmbSW.DropDownStyle = ComboBoxStyle.DropDownList
        cmbSW.Location = New Point(320, 160)
        Me.Controls.Add(cmbSW)

        btnSelect = New Button()
        btnSelect.Text = "Select Assembly (.SLDASM)"
        btnSelect.Size = New Size(260, 36)
        btnSelect.Location = New Point(480, 155)
        btnSelect.ForeColor = MM_TEXT
        AddHandler btnSelect.Click, AddressOf SelectAssembly
        Me.Controls.Add(btnSelect)

        ' -------- LOG PANEL --------
        txtLog = New TextBox()
        txtLog.Multiline = True
        txtLog.ScrollBars = ScrollBars.Vertical
        txtLog.BackColor = Color.Black
        txtLog.ForeColor = Color.LightGray
        txtLog.ReadOnly = True
        txtLog.Location = New Point(320, 210)
        txtLog.Size = New Size(Me.Width - 320 - ENG_WIDTH - 80, 470)
        Me.Controls.Add(txtLog)

        ' -------- INIT --------
        CheckVersionSupport()
        ShowLicenseStatus()
        InitialiseMacrosSafe()
        BuildMacroButtons()
        SetMacroButtonsEnabled(False)

        LogMsg("Please select a SolidWorks assembly to enable macros.")

    End Sub

    '====================================================
    ' LOGGING
    '====================================================
    Private Sub LogMsg(msg As String)
        txtLog.AppendText(DateTime.Now.ToString("HH:mm:ss") & "  " & msg & vbCrLf)
    End Sub

    '====================================================
    ' LOGO
    '====================================================
    Private Sub LoadLogo()
        Dim baseDir As String = Application.StartupPath
        For Each f In {
            Path.Combine(baseDir, "metamech_logo.png"),
            Path.Combine(baseDir, "metamech-logo.png"),
            Path.Combine(baseDir, "logo.png")
        }
            If File.Exists(f) Then
                Using fs As New FileStream(f, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    picLogo.Image = Image.FromStream(fs)
                End Using
                Exit Sub
            End If
        Next
    End Sub

    '====================================================
    ' HEADER BUTTON
    '====================================================
    Private Function CreateHeaderButton(text As String, x As Integer) As Button
        Dim b As New Button()
        b.Text = text
        b.Size = New Size(90, 30)
        b.Location = New Point(x, 25)
        b.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        b.BackColor = MM_PANEL
        b.ForeColor = MM_TEXT
        b.FlatStyle = FlatStyle.Flat
        b.FlatAppearance.BorderColor = MM_ACCENT
        b.FlatAppearance.BorderSize = 1
        b.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        Return b
    End Function

    '====================================================
    ' ENGINEERING TOOLS
    '====================================================
    Private Sub BuildEngineeringTools()
        Dim lbl As New Label()
        lbl.Text = "ENGINEERING TOOLS"
        lbl.Dock = DockStyle.Top
        lbl.Height = 42
        lbl.TextAlign = ContentAlignment.MiddleCenter
        lbl.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lbl.ForeColor = MM_ACCENT
        pnlEng.Controls.Add(lbl)

        ' >>> ONLY CHANGE <<<
        AddEngButton("Conveyor Application Tool", 60, Sub()
            Dim f As New ConveyorCalculatorForm()
            f.Show()
        End Sub)

        AddEngButton("Notepad", 104, Sub() Process.Start("notepad.exe"))
    End Sub

    Private Sub AddEngButton(text As String, y As Integer, act As Action)
        Dim b As New Button()
        b.Text = text
        b.Size = New Size(170, 36)
        b.Location = New Point(25, y)
        b.ForeColor = MM_TEXT
        AddHandler b.Click, Sub() act()
        pnlEng.Controls.Add(b)
    End Sub

    '====================================================
    ' ADMIN
    '====================================================
    Private Sub AdminLogin(sender As Object, e As EventArgs)
        Dim dlg As New AdminLoginForm()
        dlg.ShowDialog(Me)
        If dlg.LoginSuccessful Then
            Dim ed As New AdminMacroEditorForm()
            ed.ShowDialog(Me)
            InitialiseMacrosSafe()
            BuildMacroButtons()
            ApplyLicenseTierLocks()
            SetMacroButtonsEnabled(selectedAssembly <> "")
        End If
    End Sub

    '====================================================
    ' MACROS
    '====================================================
    Private Sub BuildMacroButtons()

        pnlLeft.Controls.Clear()
        Dim y As Integer = 12

        For i As Integer = 1 To 10
            If MacroNames(i) = "" OrElse MacroFiles(i) = "" Then Continue For

            Dim btn As New Button()
            btn.Text = MacroNames(i)
            btn.Size = New Size(170, 36)
            btn.Location = New Point(12, y)
            btn.Tag = i
            btn.ForeColor = MM_TEXT
            AddHandler btn.Click, AddressOf RunMacro
            pnlLeft.Controls.Add(btn)
            MacroButtons(i) = btn

            Dim gbtn As New Button()
            gbtn.Text = "Guide"
            gbtn.Size = New Size(70, 36)
            gbtn.Location = New Point(190, y)
            gbtn.Tag = i
            gbtn.ForeColor = MM_GUIDE
            AddHandler gbtn.Click, AddressOf OpenGuide
            pnlLeft.Controls.Add(gbtn)
            GuideButtons(i) = gbtn

            y += 44
        Next

        ApplyLicenseTierLocks()

    End Sub

    Private Sub InitialiseMacrosSafe()

        For i As Integer = 1 To 10
            MacroNames(i) = ""
            MacroFiles(i) = ""
            MacroModules(i) = "Module1"
            MacroMethods(i) = "main"
        Next

        If Not File.Exists(CONFIG_FILE) Then Exit Sub

        Dim xd As New XmlDocument()
        xd.Load(CONFIG_FILE)

        For Each n As XmlNode In xd.SelectNodes("//Macro")
            Dim idx As Integer = CInt(n.Attributes("slot").Value)
            MacroNames(idx) = n("Name").InnerText
            MacroFiles(idx) = n("File").InnerText
            MacroModules(idx) = n("Module").InnerText
            MacroMethods(idx) = n("Method").InnerText
        Next

    End Sub

    '====================================================
    ' LICENSE / VERSION
    '====================================================
    Private Sub CheckVersionSupport()
        If New Version(APP_VERSION) < New Version(MIN_SUPPORTED_VERSION) Then
            MessageBox.Show(
                "This version is no longer supported." & vbCrLf &
                "Download the latest version from:" & vbCrLf &
                "https://www.metamechsolutions.ie",
                "Update Required",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
        End If
    End Sub

    Private Sub ShowLicenseStatus()

        Dim tier As Integer = GetLicenseTier()
        Dim exp As DateTime = GetLicenseExpiry()

        If tier = 0 Then
            lblLicense.Text = "License: TRIAL (All Features)"
        ElseIf tier > 0 Then
            If exp <> DateTime.MinValue Then
                lblLicense.Text = "License: Tier " & tier & " (Expires: " & exp.ToString("dd-MMM-yyyy") & ")"
            Else
                lblLicense.Text = "License: Tier " & tier
            End If
        Else
            lblLicense.Text = "License: INVALID"
        End If

        If exp <> DateTime.MinValue Then
            Dim daysLeft As Integer = CInt((exp - DateTime.Now).TotalDays)
            If daysLeft <= 7 AndAlso daysLeft > 0 Then
                LogMsg("WARNING: License expires in " & daysLeft & " days.")
            ElseIf daysLeft <= 0 Then
                MessageBox.Show(
                    "Your license has expired. Please contact MetaMech Solutions.",
                    "License Expired",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
            End If
        Else
            LogMsg("License provided via email. Place license.key beside the EXE.")
        End If

    End Sub


    Private Function GetLicenseTier() As Integer
        Return ReadLicenseInteger({"CurrentTier", "Tier", "LicenseTier"}, 0)
    End Function

    Private Function GetLicenseExpiry() As DateTime
        Try
            Dim t As Type = Type.GetType("LicenseInfo")
            If t Is Nothing Then Return DateTime.MinValue
            Dim p = t.GetProperty("ExpiryDate", BindingFlags.Public Or BindingFlags.Static)
            If p IsNot Nothing Then Return CType(p.GetValue(Nothing, Nothing), DateTime)
        Catch
        End Try
        Return DateTime.MinValue
    End Function

    Private Function ReadLicenseInteger(names() As String, fallback As Integer) As Integer
        Try
            Dim t As Type = Type.GetType("LicenseInfo")
            If t Is Nothing Then Return fallback
            For Each n In names
                Dim p = t.GetProperty(n, BindingFlags.Public Or BindingFlags.Static)
                If p IsNot Nothing Then Return CInt(p.GetValue(Nothing, Nothing))
            Next
        Catch
        End Try
        Return fallback
    End Function

    Private Sub ApplyLicenseTierLocks()
        Dim tier As Integer = GetLicenseTier()
        Dim maxSlot As Integer = If(tier = 1, 3, If(tier = 2, 6, 10))

        For i As Integer = 1 To 10
            If MacroButtons(i) IsNot Nothing Then
                MacroButtons(i).Enabled = (i <= maxSlot AndAlso selectedAssembly <> "")
            End If
        Next
    End Sub

    '====================================================
    ' EXECUTION
    '====================================================
    Private Sub RunMacro(sender As Object, e As EventArgs)

        If selectedAssembly = "" Then
            MessageBox.Show("Please select a SolidWorks assembly first.", "Assembly Required")
            LogMsg("Macro blocked: no assembly selected.")
            Exit Sub
        End If

        If isRunning Then Exit Sub

        Dim idx As Integer = CInt(CType(sender, Button).Tag)

        Try
            isRunning = True
            Dim swApp As Object = CreateObject(GetSWProgId())
            CallByName(swApp, "Visible", CallType.Set, True)
            CallByName(swApp, "OpenDoc6", CallType.Method, selectedAssembly, 2, 0, "", 0, 0)
            CallByName(swApp, "RunMacro", CallType.Method, MacroFiles(idx), MacroModules(idx), MacroMethods(idx))
            LogMsg("Executed: " & MacroNames(idx))
        Catch ex As Exception
            LogMsg("ERROR: " & ex.Message)
        Finally
            isRunning = False
        End Try

    End Sub

    Private Sub SelectAssembly(sender As Object, e As EventArgs)
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "SolidWorks Assembly (*.sldasm)|*.sldasm"
        If ofd.ShowDialog() = DialogResult.OK Then
            selectedAssembly = ofd.FileName
            LogMsg("Selected: " & selectedAssembly)
            SetMacroButtonsEnabled(True)
            ApplyLicenseTierLocks()
        End If
    End Sub

    Private Sub SetMacroButtonsEnabled(en As Boolean)
        For i As Integer = 1 To 10
            If MacroButtons(i) IsNot Nothing Then
                MacroButtons(i).Enabled = en
                GuideButtons(i).Enabled = en
            End If
        Next
    End Sub

    Private Sub OpenGuide(sender As Object, e As EventArgs)
        Dim idx As Integer = CInt(CType(sender, Button).Tag)
        Dim p As String = Path.Combine(Application.StartupPath, "Guides", "Macro" & idx & ".docx")
        If File.Exists(p) Then Process.Start(p)
    End Sub

    Private Function GetSWProgId() As String
        Select Case cmbSW.SelectedItem.ToString()
            Case "2022" : Return "SldWorks.Application.30"
            Case "2023" : Return "SldWorks.Application.31"
            Case "2024" : Return "SldWorks.Application.32"
            Case Else : Return "SldWorks.Application.33"
        End Select
    End Function

End Class
