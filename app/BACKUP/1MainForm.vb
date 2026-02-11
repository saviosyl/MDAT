Option Strict On
Option Explicit On

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

        ' -------- LOGO --------
        picLogo = New PictureBox()
        picLogo.Size = New Size(320, 130)
        picLogo.Location = New Point(-70, 10)
        picLogo.SizeMode = PictureBoxSizeMode.Zoom
        picLogo.BackColor = Color.Transparent
        LoadLogo()
        Me.Controls.Add(picLogo)

        ' -------- TITLE --------
        Dim lblTitle As New Label()
        lblTitle.Text = "Mechanical Design Automation"
        lblTitle.Font = New Font("Segoe UI", 20, FontStyle.Bold)
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(170, 25)
        Me.Controls.Add(lblTitle)

        Dim lblBrand As New Label()
        lblBrand.Text = "@ Designed by MetaMech Solutions"
        lblBrand.Font = New Font("Segoe UI", 10, FontStyle.Italic)
        lblBrand.ForeColor = MM_SUB
        lblBrand.AutoSize = True
        lblBrand.Location = New Point(170, 65)
        Me.Controls.Add(lblBrand)

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
        AddHandler btnSelect.Click, AddressOf SelectAssembly
        Me.Controls.Add(btnSelect)

        ' -------- LOG --------
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
        LoadLicense()
        InitialiseMacrosSafe()
        BuildMacroButtons()
        ApplyLicenseTierLocks()
        SetMacroButtonsEnabled(False)

        LogMsg("Please select a SolidWorks assembly to enable macros.")

    End Sub

    '====================================================
    ' LICENSE
    '====================================================
    Private Sub LoadLicense()
        Dim lic As LicenseInfo = Licensing.LoadAndVerifyLicense()

        If lic IsNot Nothing Then
            lblLicense.Text =
                If(lic.IsValid,
                   "License: Tier " & lic.Tier.ToString(),
                   "License: INVALID")

            If lic.IsExpired Then
                MessageBox.Show("License expired.", "License", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Else
            lblLicense.Text = "License: NOT FOUND"
        End If
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
        End If
    End Sub

    '====================================================
    ' MACROS
    '====================================================
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

    Private Sub BuildMacroButtons()
        pnlLeft.Controls.Clear()
        Dim y As Integer = 12

        For i As Integer = 1 To 10
            If MacroNames(i) = "" Then Continue For

            Dim btn As New Button()
            btn.Text = MacroNames(i)
            btn.Size = New Size(170, 36)
            btn.Location = New Point(12, y)
            btn.Tag = i
            AddHandler btn.Click, AddressOf RunMacro
            pnlLeft.Controls.Add(btn)
            MacroButtons(i) = btn

            Dim g As New Button()
            g.Text = "Guide"
            g.Size = New Size(70, 36)
            g.Location = New Point(190, y)
            g.Tag = i
            AddHandler g.Click, AddressOf OpenGuide
            pnlLeft.Controls.Add(g)
            GuideButtons(i) = g

            y += 44
        Next
    End Sub

    Private Sub ApplyLicenseTierLocks()
        Dim tier As Integer = LicenseInfo.Tier

        For i As Integer = 1 To 10
            If MacroButtons(i) IsNot Nothing Then
                MacroButtons(i).Enabled =
                    TierLocks.IsUnlocked(tier, i) AndAlso selectedAssembly <> ""
            End If
        Next
    End Sub

    '====================================================
    ' EXECUTION
    '====================================================
    Private Sub RunMacro(sender As Object, e As EventArgs)

        If selectedAssembly = "" Then Exit Sub
        If isRunning Then Exit Sub

        Dim idx As Integer = CInt(CType(sender, Button).Tag)

        Try
            isRunning = True

            Dim swApp As Object = CreateObject(GetSWProgId())
            CallByName(swApp, "Visible", CallType.Set, True)
            CallByName(swApp, "OpenDoc6", CallType.Method, selectedAssembly, 2, 0, "", 0, 0)
            CallByName(swApp, "RunMacro", CallType.Method,
                       MacroFiles(idx), MacroModules(idx), MacroMethods(idx))

            LogMsg("Executed macro: " & MacroNames(idx))

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
            End If
            If GuideButtons(i) IsNot Nothing Then
                GuideButtons(i).Enabled = en
            End If
        Next
    End Sub

    Private Sub OpenGuide(sender As Object, e As EventArgs)
        Dim idx As Integer = CInt(CType(sender, Button).Tag)
        Dim p As String = Path.Combine(Application.StartupPath, "Guides", "Macro" & idx & ".docx")
        If File.Exists(p) Then Process.Start(p)
    End Sub

    '====================================================
    ' ENGINEERING TOOLS
    '====================================================
    Private Sub BuildEngineeringTools()

        pnlEng.Controls.Clear()

        AddEngButton("Conveyor Application Tool", 60, 0, Sub()
            Dim f As New ConveyorCalculatorForm()
            f.Show()
        End Sub)

        AddEngButton("FlexLink Calculator", 104, 2, Sub()
            Dim f As New FlexLinkCalculatorForm()
            f.Show()
        End Sub)

        AddEngButton("Unit Converter", 148, 0, Sub()
            Dim f As New UnitConverterForm()
            f.Show()
        End Sub)

        AddEngButton("Notepad", 192, 0, Sub()
            Process.Start("notepad.exe")
        End Sub)

    End Sub

    Private Sub AddEngButton(text As String, y As Integer, minTier As Integer, act As Action)
        Dim b As New Button()
        b.Text = text
        b.Size = New Size(170, 36)
        b.Location = New Point(25, y)
        b.Enabled = (LicenseInfo.Tier >= minTier)
        AddHandler b.Click, Sub()
                                If LicenseInfo.Tier < minTier Then
                                    MessageBox.Show("This tool requires Tier " & minTier & " license.")
                                    Exit Sub
                                End If
                                act()
                            End Sub
        pnlEng.Controls.Add(b)
    End Sub

    '====================================================
    ' HELPERS
    '====================================================
    Private Sub CheckVersionSupport()
        If New Version(APP_VERSION) < New Version(MIN_SUPPORTED_VERSION) Then
            MessageBox.Show("Unsupported version.", "Update Required")
        End If
    End Sub

    Private Function CreateHeaderButton(text As String, x As Integer) As Button
        Dim b As New Button()
        b.Text = text
        b.Size = New Size(90, 30)
        b.Location = New Point(x, 25)
        b.FlatStyle = FlatStyle.Flat
        Return b
    End Function

    Private Sub LoadLogo()
        Dim p As String = Path.Combine(Application.StartupPath, "metamech-logo.png")
        If File.Exists(p) Then picLogo.Image = Image.FromFile(p)
    End Sub

    Private Function GetSWProgId() As String
        Select Case cmbSW.SelectedItem.ToString()
            Case "2022" : Return "SldWorks.Application.30"
            Case "2023" : Return "SldWorks.Application.31"
            Case "2024" : Return "SldWorks.Application.32"
            Case Else : Return "SldWorks.Application.33"
        End Select
    End Function

    Private Sub LogMsg(msg As String)
        txtLog.AppendText(DateTime.Now.ToString("HH:mm:ss") & "  " & msg & vbCrLf)
    End Sub

End Class
