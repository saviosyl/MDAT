Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing.Printing
Imports System.IO
Imports System.Collections.Generic
Imports System.Globalization

Public Class PneumaticCylinderCalculatorForm
    Inherits Form

    ' ============================================================
    ' THEME (Defaults = MetaMech dark, but MDAT can override via ApplyMDATTheme)
    ' IMPORTANT: Reset must NOT force these defaults back on.
    ' ============================================================
    Private themeBG As Color = Color.FromArgb(18, 22, 30)
    Private themePanel As Color = Color.FromArgb(28, 34, 44)
    Private themeAccent As Color = Color.FromArgb(150, 90, 190)
    Private themeText As Color = Color.White

    Private themeWarn As Color = Color.Gold
    Private themeErr As Color = Color.OrangeRed
    Private themeOk As Color = Color.LightGreen

    Private btnGreen As Color = Color.LightGreen
    Private btnYellow As Color = Color.Gold

    ' ============================================================
    ' DEFAULTS (used ONLY if user leaves fields blank)
    ' ============================================================
    Private Const DEFAULT_MASS_KG As Double = 1.0
    Private Const DEFAULT_MU As Double = 0.1
    Private Const DEFAULT_ACCEL As Double = 0.0

    Private Const DEFAULT_PRESSURE_BAR As Double = 6.0
    Private Const DEFAULT_PRESSURE_FACTOR_PCT As Double = 90.0
    Private Const DEFAULT_CYL_EFF_PCT As Double = 95.0

    Private Const DEFAULT_BORE_MM As Double = 32.0
    Private Const DEFAULT_ROD_MM As Double = 12.0

    Private Const DEFAULT_PEAK_FACTOR As Double = 1.3
    Private Const DEFAULT_SF As Double = 1.2

    ' Cycle time defaults (optional)
    Private Const DEFAULT_STROKE_MM As Double = 100.0
    Private Const DEFAULT_EXT_SPEED_MM_S As Double = 200.0
    Private Const DEFAULT_RET_SPEED_MM_S As Double = 250.0
    Private Const DEFAULT_DWELL_EXT_S As Double = 0.2
    Private Const DEFAULT_DWELL_RET_S As Double = 0.2

    ' Typical ISO cylinder bores (shop-floor common)
    Private Shared ReadOnly STANDARD_BORES_MM() As Double = New Double() { _
        8, 10, 12, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 320 _
    }

    ' Common rod sizes by bore (rule-of-thumb)
    Private Shared ReadOnly ROD_SUGGEST_TABLE As Dictionary(Of Integer, Integer) = New Dictionary(Of Integer, Integer)() From {
        {8, 3},
        {10, 4},
        {12, 4},
        {16, 6},
        {20, 8},
        {25, 10},
        {32, 12},
        {40, 16},
        {50, 20},
        {63, 20},
        {80, 25},
        {100, 32},
        {125, 36},
        {160, 40},
        {200, 50},
        {250, 63},
        {320, 80}
    }

    ' ============================================================
    ' LAYOUT (SplitContainer for resize + full screen)
    ' ============================================================
    Private split As SplitContainer
    Private pnlInput As Panel
    Private pnlResults As Panel

    Private desiredLeftWidth As Integer = 360
    Private desiredMinLeft As Integer = 320
    Private desiredMinRight As Integer = 520

    ' ============================================================
    ' OUTPUT
    ' ============================================================
    Private txtFinal As TextBox
    Private txtCalc As TextBox
    Private rtbWarn As RichTextBox
    Private lblStatus As Label

    Private lblOutFinal As Label
    Private lblOutCalc As Label
    Private lblOutWarn As Label

    ' ============================================================
    ' INPUTS
    ' ============================================================
    ' LOAD / MOTION
    Private txtMassKg As TextBox
    Private txtMu As TextBox
    Private txtAccel As TextBox

    Private cmbAnglePreset As ComboBox
    Private txtAngleCustom As TextBox

    ' CYLINDER
    Private cmbAction As ComboBox
    Private txtPressureBar As TextBox
    Private txtPressureFactor As TextBox
    Private txtCylEff As TextBox

    Private txtBoreMm As TextBox
    Private txtRodMm As TextBox

    ' FACTORS
    Private txtPeakFactor As TextBox
    Private txtSF As TextBox

    ' CYCLE TIME
    Private chkCycle As CheckBox
    Private txtStrokeMm As TextBox
    Private txtExtSpeed As TextBox
    Private txtRetSpeed As TextBox
    Private txtDwellExt As TextBox
    Private txtDwellRet As TextBox

    ' Buttons
    Private btnCalc As Button
    Private btnPdf As Button
    Private btnReset As Button

    ' Tooltips
    Private tip As ToolTip
    Private tipMap As Dictionary(Of Control, String)
    Private tipsArmed As Boolean = False

    ' ============================================================
    ' INIT
    ' ============================================================
    Public Sub New()
        Me.Text = "Pneumatic Cylinder Calculator"
        Me.Size = New Size(1180, 800)
        Me.MinimumSize = New Size(980, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MaximizeBox = True
        Me.BackColor = themeBG

        FormFooter.AddPremiumFooter(Me)
        BuildUI()
        FormHeader.AddPremiumHeader(Me, "Pneumatic Cylinder Calculator", "MetaMech Engineering Tools")

        AddHandler Me.MouseDown, AddressOf ArmTipsOnFirstClick

        AddHandler Me.Shown, AddressOf OnFormShownOrResized
        AddHandler Me.SizeChanged, AddressOf OnFormShownOrResized
        AddHandler Me.Resize, AddressOf OnFormShownOrResized

        SyncThemeFromCurrentUI()
        ApplyThemeToThisForm()
    End Sub

    Private Sub ArmTipsOnFirstClick(sender As Object, e As MouseEventArgs)
        tipsArmed = True
    End Sub

    Private Sub OnFormShownOrResized(sender As Object, e As EventArgs)
        ApplySplitLayoutSafe()
    End Sub

    ' ============================================================
    ' SAFE Split Layout (prevents SplitterDistance crash in small host windows)
    ' ============================================================
    Private Sub ApplySplitLayoutSafe()
        If split Is Nothing Then Exit Sub

        Try
            Dim w As Integer = split.Width
            Dim sw As Integer = split.SplitterWidth
            If w <= 0 Then Exit Sub

            If w < (desiredMinLeft + desiredMinRight + sw) Then
                split.Panel1MinSize = 0
                split.Panel2MinSize = 0
                SafeSetSplitterDistance(Math.Min(desiredLeftWidth, Math.Max(0, w - sw)))
                Exit Sub
            End If

            split.Panel1MinSize = 0
            split.Panel2MinSize = 0

            Dim maxDistNew As Integer = w - desiredMinRight - sw
            Dim dist As Integer = desiredLeftWidth
            If dist < desiredMinLeft Then dist = desiredMinLeft
            If dist > maxDistNew Then dist = maxDistNew

            SafeSetSplitterDistance(dist)

            split.Panel1MinSize = desiredMinLeft
            split.Panel2MinSize = desiredMinRight

            SafeSetSplitterDistance(dist)

        Catch
            Try
                split.Panel1MinSize = 0
                split.Panel2MinSize = 0
            Catch
            End Try
        End Try
    End Sub

    Private Sub SafeSetSplitterDistance(desired As Integer)
        If split Is Nothing Then Exit Sub

        Try
            Dim w As Integer = split.Width
            Dim sw As Integer = split.SplitterWidth
            If w <= 0 Then Exit Sub

            Dim min1 As Integer = split.Panel1MinSize
            Dim min2 As Integer = split.Panel2MinSize

            Dim maxDist As Integer = w - min2 - sw
            If maxDist < min1 Then
                split.Panel1MinSize = 0
                split.Panel2MinSize = 0
                min1 = 0
                min2 = 0
                maxDist = w - sw
            End If

            Dim dist As Integer = desired
            If dist < min1 Then dist = min1
            If dist > maxDist Then dist = maxDist
            If dist < 0 Then dist = 0

            split.SplitterDistance = dist
        Catch
        End Try
    End Sub

    ' ============================================================
    ' MDAT THEME HOOK
    ' ============================================================
    Public Sub ApplyMDATTheme(bg As Color, panel As Color, accent As Color, isDark As Boolean)
        themeBG = bg
        themePanel = panel
        themeAccent = accent

        If isDark Then
            themeText = Color.Gainsboro
        Else
            themeText = Color.FromArgb(35, 35, 35)
        End If

        themeWarn = Color.Gold
        themeErr = Color.OrangeRed
        themeOk = Color.LightGreen

        btnGreen = Color.LightGreen
        btnYellow = Color.Gold

        ApplyThemeToThisForm()
    End Sub

    Private Sub SyncThemeFromCurrentUI()
        Try
            themeBG = Me.BackColor
        Catch
        End Try

        Try
            If pnlInput IsNot Nothing Then themePanel = pnlInput.BackColor
        Catch
        End Try

        Try
            If btnCalc IsNot Nothing Then themeAccent = btnCalc.BackColor
        Catch
        End Try
    End Sub

    ' ============================================================
    ' BUILD UI
    ' ============================================================
    Private Sub BuildUI()

        tip = New ToolTip()
        tip.IsBalloon = True
        tip.ShowAlways = True
        tip.AutoPopDelay = 18000
        tip.InitialDelay = 200
        tip.ReshowDelay = 50

        tipMap = New Dictionary(Of Control, String)()

        split = New SplitContainer()
        split.Dock = DockStyle.Fill
        split.Orientation = Orientation.Vertical
        split.SplitterWidth = 6

        ' CRITICAL: keep mins at 0 during construction
        split.Panel1MinSize = 0
        split.Panel2MinSize = 0

        Me.Controls.Add(split)

        pnlInput = New Panel()
        pnlInput.Dock = DockStyle.Fill
        pnlInput.BackColor = themePanel
        pnlInput.AutoScroll = True
        split.Panel1.Controls.Add(pnlInput)

        pnlResults = New Panel()
        pnlResults.Dock = DockStyle.Fill
        pnlResults.BackColor = themeBG
        split.Panel2.Controls.Add(pnlResults)

        lblStatus = New Label()
        lblStatus.Text = "Status: Ready"
        lblStatus.ForeColor = themeOk
        lblStatus.AutoSize = False
        lblStatus.Height = 26
        lblStatus.Dock = DockStyle.Bottom
        lblStatus.TextAlign = ContentAlignment.MiddleLeft
        Me.Controls.Add(lblStatus)

        Dim y As Integer = 15

        AddSection("LOAD / MOTION", y) : y += 28

        txtMassKg = AddInput("Load Mass (kg)", y,
            "What: Total mass moved by cylinder (kg)." & vbCrLf &
            "Where: scale / drawing / BOM." & vbCrLf &
            "Include: payload + fixture + gripper." & vbCrLf &
            "Example: 2.0 + 1.2 + 0.8 = 4.0 kg") : y += 26

        txtMu = AddInput("Guide Friction (mu)", y,
            "What: friction coefficient of guides (unitless)." & vbCrLf &
            "Where: guide datasheet or estimate." & vbCrLf &
            "Typical: linear 0.01-0.03, slide rails 0.05-0.15." & vbCrLf &
            "Example: dry slide rail -> 0.10") : y += 26

        txtAccel = AddInput("Acceleration (m/s^2)", y,
            "What: acceleration during motion (adds inertia force)." & vbCrLf &
            "Where: motion profile (a = v/t)." & vbCrLf &
            "Example: 0.3 m/s in 0.2 s -> 1.5 m/s^2") : y += 30

        AddLabel("Angle Preset", y)
        cmbAnglePreset = AddCombo(y, New String() {"0 deg (Vertical Up)", "45 deg", "90 deg (Horizontal)", "180 deg (Vertical Down)", "Custom"},
            "Angle measured from vertical-up direction." & vbCrLf &
            "0 = lift up (worst for gravity), 90 = horizontal, 180 = down (gravity helps).")
        AddHandler cmbAnglePreset.SelectedIndexChanged, AddressOf AnglePresetChanged
        y += 26

        txtAngleCustom = AddInput("Angle Custom (deg)", y,
            "Custom angle from vertical-up (0..180)." & vbCrLf &
            "Example: 30") : y += 34

        AddSection("CYLINDER", y) : y += 28

        AddLabel("Action", y)
        cmbAction = AddCombo(y, New String() {"Extend (Push)", "Retract (Pull)"},
            "Extend uses bore area." & vbCrLf &
            "Retract uses annulus area (bore minus rod).")
        y += 26

        txtPressureBar = AddInput("Supply Pressure (bar)", y,
            "Pressure at regulator near cylinder." & vbCrLf &
            "Typical: 4..7 bar." & vbCrLf &
            "Example: 6") : y += 26

        txtPressureFactor = AddInput("Pressure Factor (%)", y,
            "Allows losses in valve/tube (effective pressure)." & vbCrLf &
            "Typical: 80..95." & vbCrLf &
            "Example: 90 means you assume only 90% of pressure reaches cylinder.") : y += 26

        txtCylEff = AddInput("Cylinder Efficiency (%)", y,
            "Seal/friction efficiency." & vbCrLf &
            "Typical: 90..97." & vbCrLf &
            "Example: 95") : y += 30

        txtBoreMm = AddInput("Bore Diameter (mm)", y,
            "Cylinder bore (piston diameter)." & vbCrLf &
            "Where: cylinder datasheet / model code." & vbCrLf &
            "Example: ISO D32 -> 32") : y += 26

        txtRodMm = AddInput("Rod Diameter (mm)", y,
            "Rod diameter affects retract force & stiffness." & vbCrLf &
            "Where: datasheet." & vbCrLf &
            "Tip: leave blank to auto-suggest typical rod for the selected bore.") : y += 34

        AddSection("FACTORS", y) : y += 28

        txtPeakFactor = AddInput("Peak / Breakaway Factor", y,
            "Extra factor for breakaway/shock." & vbCrLf &
            "Typical: 1.1..1.5. Example: 1.3") : y += 26

        txtSF = AddInput("Safety Factor", y,
            "Engineering safety margin." & vbCrLf &
            "Typical: 1.1..1.5. Example: 1.2") : y += 34

        AddSection("CYCLE TIME (OPTIONAL)", y) : y += 28

        chkCycle = AddCheck("Enable Cycle Time Calculation", y,
            "Adds simple cycle time estimate:" & vbCrLf &
            "Cycle = extend time + dwell(ext) + retract time + dwell(ret)." & vbCrLf &
            "Use typical speeds from valve settings or supplier data.")
        AddHandler chkCycle.CheckedChanged, AddressOf ToggleCycle
        y += 24

        txtStrokeMm = AddInput("Stroke (mm)", y,
            "What: cylinder stroke length (mm)." & vbCrLf &
            "Where: cylinder datasheet / drawing." & vbCrLf &
            "Example: 100 mm stroke.") : y += 24

        txtExtSpeed = AddInput("Extend Speed (mm/s)", y,
            "What: average extend speed (mm/s)." & vbCrLf &
            "Where: speed control setting, supplier chart, or measured test." & vbCrLf &
            "Typical: 50..500 mm/s (depends on load/valve)." & vbCrLf &
            "Example: 200 mm/s -> 100 mm stroke takes 0.5 s.") : y += 24

        txtRetSpeed = AddInput("Retract Speed (mm/s)", y,
            "What: average retract speed (mm/s)." & vbCrLf &
            "Where: speed control setting, supplier chart, or measured test." & vbCrLf &
            "Tip: retract can be faster/slower than extend." & vbCrLf &
            "Example: 250 mm/s -> 100 mm stroke takes 0.4 s.") : y += 24

        txtDwellExt = AddInput("Dwell @ Extended (s)", y,
            "What: wait time when fully extended (seconds)." & vbCrLf &
            "Where: PLC timer / sequence chart." & vbCrLf &
            "Examples:" & vbCrLf &
            " - 0.20 s for clamp settle" & vbCrLf &
            " - 0.50 s for process / sensor confirm") : y += 24

        txtDwellRet = AddInput("Dwell @ Retracted (s)", y,
            "What: wait time when fully retracted (seconds)." & vbCrLf &
            "Where: PLC timer / sequence chart." & vbCrLf &
            "Example: 0.20 s for part clear / next step ready.") : y += 30

        btnCalc = CreateButton("Calculate", themeAccent, New Point(20, y), True, AddressOf Calculate)
        btnPdf = CreateButton("Export PDF", btnGreen, New Point(130, y), False, AddressOf ExportPdf)
        btnReset = CreateButton("Reset", btnYellow, New Point(240, y), False, AddressOf ResetAll)
        pnlInput.Controls.Add(btnCalc)
        pnlInput.Controls.Add(btnPdf)
        pnlInput.Controls.Add(btnReset)

        ' OUTPUTS - Docked
        rtbWarn = New RichTextBox()
        rtbWarn.ReadOnly = True
        rtbWarn.BorderStyle = BorderStyle.FixedSingle
        rtbWarn.Multiline = True
        rtbWarn.ScrollBars = RichTextBoxScrollBars.Vertical
        rtbWarn.Dock = DockStyle.Fill

        txtCalc = New TextBox()
        txtCalc.Multiline = True
        txtCalc.ReadOnly = True
        txtCalc.ScrollBars = ScrollBars.Vertical
        txtCalc.Dock = DockStyle.Top
        txtCalc.Height = 220

        txtFinal = New TextBox()
        txtFinal.Multiline = True
        txtFinal.ReadOnly = True
        txtFinal.ScrollBars = ScrollBars.Vertical
        txtFinal.Dock = DockStyle.Top
        txtFinal.Height = 230

        lblOutWarn = New Label()
        lblOutWarn.Text = "WARNINGS / ERRORS"
        lblOutWarn.ForeColor = themeAccent
        lblOutWarn.AutoSize = False
        lblOutWarn.Height = 20
        lblOutWarn.Dock = DockStyle.Top

        lblOutCalc = New Label()
        lblOutCalc.Text = "CALCULATION METHOD (STEP-BY-STEP)"
        lblOutCalc.ForeColor = themeAccent
        lblOutCalc.AutoSize = False
        lblOutCalc.Height = 20
        lblOutCalc.Dock = DockStyle.Top

        lblOutFinal = New Label()
        lblOutFinal.Text = "FINAL RESULTS"
        lblOutFinal.ForeColor = themeAccent
        lblOutFinal.AutoSize = False
        lblOutFinal.Height = 20
        lblOutFinal.Dock = DockStyle.Top

        pnlResults.Controls.Add(rtbWarn)
        pnlResults.Controls.Add(lblOutWarn)
        pnlResults.Controls.Add(txtCalc)
        pnlResults.Controls.Add(lblOutCalc)
        pnlResults.Controls.Add(txtFinal)
        pnlResults.Controls.Add(lblOutFinal)

        If cmbAnglePreset IsNot Nothing AndAlso cmbAnglePreset.Items.Count > 0 Then cmbAnglePreset.SelectedIndex = 2 ' 90 deg default
        If cmbAction IsNot Nothing AndAlso cmbAction.Items.Count > 0 Then cmbAction.SelectedIndex = 0

        chkCycle.Checked = False
        ToggleCycle(Nothing, EventArgs.Empty)
        AnglePresetChanged(Nothing, EventArgs.Empty)
    End Sub

    ' ============================================================
    ' UI HELPERS
    ' ============================================================
    Private Sub AddSection(t As String, y As Integer)
        Dim lbl As New Label()
        lbl.Text = t
        lbl.ForeColor = themeAccent
        lbl.Location = New Point(10, y)
        lbl.AutoSize = True
        Try
            lbl.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        Catch
        End Try
        pnlInput.Controls.Add(lbl)
    End Sub

    Private Sub AddLabel(t As String, y As Integer)
        Dim lbl As New Label()
        lbl.Text = t
        lbl.ForeColor = themeText
        lbl.Location = New Point(14, y)
        lbl.AutoSize = True
        pnlInput.Controls.Add(lbl)
    End Sub

    Private Function AddInput(t As String, y As Integer, tipText As String) As TextBox
        AddLabel(t, y)

        Dim tb As New TextBox()
        tb.Location = New Point(170, y - 2)
        tb.Width = 160
        tb.Text = ""
        pnlInput.Controls.Add(tb)

        RegisterTip(tb, tipText)
        AddHandler tb.MouseDown, AddressOf ArmTipsOnFirstClick

        Return tb
    End Function

    Private Function AddCombo(y As Integer, items() As String, tipText As String) As ComboBox
        Dim cb As New ComboBox()
        cb.Location = New Point(170, y)
        cb.Width = 160
        cb.DropDownStyle = ComboBoxStyle.DropDownList
        cb.Items.AddRange(items)
        If cb.Items.Count > 0 Then cb.SelectedIndex = 0
        pnlInput.Controls.Add(cb)

        RegisterTip(cb, tipText)
        AddHandler cb.MouseDown, AddressOf ArmTipsOnFirstClick

        Return cb
    End Function

    Private Function AddCheck(t As String, y As Integer, tipText As String) As CheckBox
        Dim cb As New CheckBox()
        cb.Text = t
        cb.ForeColor = themeText
        cb.Location = New Point(14, y)
        pnlInput.Controls.Add(cb)

        RegisterTip(cb, tipText)
        AddHandler cb.MouseDown, AddressOf ArmTipsOnFirstClick

        Return cb
    End Function

    Private Function CreateButton(t As String, c As Color, p As Point, isPrimary As Boolean, h As EventHandler) As Button
        Dim b As New Button()
        b.Text = t
        b.BackColor = c
        b.ForeColor = If(isPrimary, Color.White, Color.Black)
        b.Location = p
        b.Size = New Size(90, 32)
        b.FlatStyle = FlatStyle.Flat
        b.FlatAppearance.BorderSize = 1
        AddHandler b.Click, h
        AddHandler b.MouseDown, AddressOf ArmTipsOnFirstClick
        Return b
    End Function

    Private Sub ToggleCycle(sender As Object, e As EventArgs)
        Dim en As Boolean = (chkCycle IsNot Nothing AndAlso chkCycle.Checked)
        SetEnabled(txtStrokeMm, en)
        SetEnabled(txtExtSpeed, en)
        SetEnabled(txtRetSpeed, en)
        SetEnabled(txtDwellExt, en)
        SetEnabled(txtDwellRet, en)
    End Sub

    Private Sub AnglePresetChanged(sender As Object, e As EventArgs)
        Dim isCustom As Boolean = (cmbAnglePreset IsNot Nothing AndAlso String.Equals(cmbAnglePreset.Text, "Custom", StringComparison.OrdinalIgnoreCase))
        SetEnabled(txtAngleCustom, isCustom)
        If Not isCustom AndAlso txtAngleCustom IsNot Nothing Then txtAngleCustom.Text = ""
    End Sub

    Private Sub SetEnabled(c As Control, en As Boolean)
        If c Is Nothing Then Exit Sub
        Try
            c.Enabled = en
        Catch
        End Try
    End Sub

    ' ============================================================
    ' APPLY THEME
    ' ============================================================
    Private Sub ApplyThemeToThisForm()
        Try
            Me.BackColor = themeBG
        Catch
        End Try

        Try
            If pnlInput IsNot Nothing Then pnlInput.BackColor = themePanel
            If pnlResults IsNot Nothing Then pnlResults.BackColor = themeBG
            If split IsNot Nothing Then
                split.BackColor = themeBG
                split.Panel1.BackColor = themePanel
                split.Panel2.BackColor = themeBG
            End If
        Catch
        End Try

        Try
            ApplyThemeRecursive(Me)
        Catch
        End Try

        Try
            If btnCalc IsNot Nothing Then
                btnCalc.BackColor = themeAccent
                btnCalc.ForeColor = Color.White
            End If
            If btnPdf IsNot Nothing Then
                btnPdf.BackColor = btnGreen
                btnPdf.ForeColor = Color.Black
            End If
            If btnReset IsNot Nothing Then
                btnReset.BackColor = btnYellow
                btnReset.ForeColor = Color.Black
            End If

            If lblOutFinal IsNot Nothing Then lblOutFinal.ForeColor = themeAccent
            If lblOutCalc IsNot Nothing Then lblOutCalc.ForeColor = themeAccent
            If lblOutWarn IsNot Nothing Then lblOutWarn.ForeColor = themeAccent
        Catch
        End Try
    End Sub

    Private Sub ApplyThemeRecursive(root As Control)
        If root Is Nothing Then Exit Sub

        For Each c As Control In root.Controls

            If TypeOf c Is Panel Then
                Dim p As Panel = DirectCast(c, Panel)
                If p Is pnlInput Then
                    p.BackColor = themePanel
                ElseIf p Is pnlResults Then
                    p.BackColor = themeBG
                Else
                    p.BackColor = themePanel
                End If
                p.ForeColor = themeText

            ElseIf TypeOf c Is SplitContainer Then
                Dim sc As SplitContainer = DirectCast(c, SplitContainer)
                sc.BackColor = themeBG
                sc.Panel1.BackColor = themePanel
                sc.Panel2.BackColor = themeBG

            ElseIf TypeOf c Is Label Then
                Dim lbl As Label = DirectCast(c, Label)
                lbl.BackColor = Color.Transparent
                If lbl.ForeColor <> themeAccent Then lbl.ForeColor = themeText

            ElseIf TypeOf c Is TextBox Then
                Dim tb As TextBox = DirectCast(c, TextBox)
                If tb.ReadOnly Then
                    tb.BackColor = themePanel
                Else
                    tb.BackColor = themeBG
                End If
                tb.ForeColor = themeText

            ElseIf TypeOf c Is RichTextBox Then
                Dim r As RichTextBox = DirectCast(c, RichTextBox)
                r.BackColor = themePanel
                r.ForeColor = themeText

            ElseIf TypeOf c Is ComboBox Then
                Dim cb As ComboBox = DirectCast(c, ComboBox)
                cb.BackColor = themeBG
                cb.ForeColor = themeText

            ElseIf TypeOf c Is CheckBox Then
                Dim chk As CheckBox = DirectCast(c, CheckBox)
                chk.ForeColor = themeText
                chk.BackColor = Color.Transparent

            ElseIf TypeOf c Is Button Then
                Dim b As Button = DirectCast(c, Button)
                b.FlatStyle = FlatStyle.Flat
                b.FlatAppearance.BorderSize = 1
                b.FlatAppearance.BorderColor = themeAccent
            End If

            If c.HasChildren Then ApplyThemeRecursive(c)
        Next
    End Sub

    ' ============================================================
    ' CALCULATE (REAL OUTPUTS)
    ' ============================================================
    Private Sub Calculate(sender As Object, e As EventArgs)
        SyncThemeFromCurrentUI()
        ClearOutputs()

        Dim mKg As Double = ReadD(txtMassKg, DEFAULT_MASS_KG)
        Dim mu As Double = ReadD(txtMu, DEFAULT_MU)
        Dim accel As Double = ReadD(txtAccel, DEFAULT_ACCEL)

        Dim angleDeg As Double = GetAngleDeg()
        Dim actionName As String = If(cmbAction Is Nothing, "Extend (Push)", cmbAction.Text)

        Dim pBar As Double = ReadD(txtPressureBar, DEFAULT_PRESSURE_BAR)
        Dim pFactorPct As Double = ReadD(txtPressureFactor, DEFAULT_PRESSURE_FACTOR_PCT)
        Dim cylEffPct As Double = ReadD(txtCylEff, DEFAULT_CYL_EFF_PCT)

        Dim boreMm As Double = ReadD(txtBoreMm, DEFAULT_BORE_MM)

        Dim rodProvided As Boolean = False
        Dim rodMm As Double = 0.0
        If txtRodMm IsNot Nothing Then
            Dim sRod As String = If(txtRodMm.Text, "").Trim()
            If sRod <> "" Then
                rodMm = ReadD(txtRodMm, DEFAULT_ROD_MM)
                rodProvided = True
            End If
        End If

        Dim peakFactor As Double = ReadD(txtPeakFactor, DEFAULT_PEAK_FACTOR)
        Dim sf As Double = ReadD(txtSF, DEFAULT_SF)

        ' Basic validation
        If mKg <= 0 Then AddError("Load Mass must be > 0 kg.") : SetStatusError("Status: Error") : Return
        If pBar <= 0 Then AddError("Supply Pressure must be > 0 bar.") : SetStatusError("Status: Error") : Return
        If boreMm <= 0 Then AddError("Bore Diameter must be > 0 mm.") : SetStatusError("Status: Error") : Return

        If mu < 0 Then AddWarn("Friction mu < 0 not valid. Using 0.") : mu = 0
        If accel < 0 Then AddWarn("Acceleration < 0 treated as 0.") : accel = 0
        If pFactorPct <= 0 OrElse pFactorPct > 100 Then AddWarn("Pressure Factor out of range. Using 90%.") : pFactorPct = DEFAULT_PRESSURE_FACTOR_PCT
        If cylEffPct <= 0 OrElse cylEffPct > 100 Then AddWarn("Cylinder Efficiency out of range. Using 95%.") : cylEffPct = DEFAULT_CYL_EFF_PCT
        If peakFactor <= 0 Then AddWarn("Peak Factor must be > 0. Using 1.3.") : peakFactor = DEFAULT_PEAK_FACTOR
        If sf <= 0 Then AddWarn("Safety Factor must be > 0. Using 1.2.") : sf = DEFAULT_SF

        ' Auto rod suggestion if blank
        Dim rodAutoSuggested As Boolean = False
        If Not rodProvided Then
            rodMm = SuggestRodForBore(boreMm)
            rodAutoSuggested = True
            AddWarn("Rod Diameter was blank -> auto-suggested rod = " & rodMm.ToString("0") & " mm (typical).")
        End If
        If rodMm < 0 Then rodMm = 0

        ' Physics (angle measured from vertical-up)
        Dim g As Double = 9.81
        Dim thetaRad As Double = angleDeg * Math.PI / 180.0

        Dim Fg As Double = mKg * g

        ' Component along motion axis: F_gravity = m*g*cos(theta)
        ' - theta=0 => full weight against upward motion
        ' - theta=90 => horizontal => 0
        ' - theta=180 => negative => gravity helps
        Dim FgravityAlong As Double = Fg * Math.Cos(thetaRad)

        ' Normal component: N = |m*g*sin(theta)|
        Dim N As Double = Fg * Math.Sin(thetaRad)
        If N < 0 Then N = -N
        Dim Ffric As Double = mu * N

        ' Inertia: F = m*a
        Dim Facc As Double = mKg * accel

        ' Base required force along axis (clamped)
        Dim Fbase As Double = FgravityAlong + Ffric + Facc
        Dim FbaseClamped As Double = Fbase
        If FbaseClamped < 0 Then FbaseClamped = 0

        Dim Fpeak As Double = FbaseClamped * peakFactor
        Dim Frequired As Double = Fpeak * sf

        ' Pneumatic force available
        Dim pPa As Double = pBar * 100000.0
        Dim pf As Double = pFactorPct / 100.0
        Dim eta As Double = cylEffPct / 100.0
        Dim pEff As Double = pPa * pf

        Dim boreM As Double = boreMm / 1000.0
        Dim rodM As Double = rodMm / 1000.0

        Dim Abore As Double = Math.PI * boreM * boreM / 4.0
        Dim Arod As Double = Math.PI * rodM * rodM / 4.0
        Dim Aann As Double = Abore - Arod
        If Aann < 0 Then Aann = 0

        Dim isRetract As Boolean = (actionName.IndexOf("Retract", StringComparison.OrdinalIgnoreCase) >= 0)
        Dim Aused As Double = Abore
        If isRetract Then
            Aused = Aann
            If Aused <= 0 Then
                AddError("Retract area is 0/negative (rod too large). Check Bore/Rod.")
                SetStatusError("Status: Error") : Return
            End If
        End If

        Dim Favailable As Double = pEff * Aused * eta
        If Favailable <= 0 Then
            AddError("Available cylinder force computed as 0. Check pressure/area/efficiency.")
            SetStatusError("Status: Error") : Return
        End If

        Dim usagePct As Double = (Frequired / Favailable) * 100.0

        ' Required pressure for current bore/action
        Dim pReqBar As Double = 0.0
        Dim pMarginBar As Double = 0.0
        Dim pReqOk As Boolean = False
        If Aused > 0 AndAlso pf > 0 AndAlso eta > 0 Then
            pReqBar = Frequired / (100000.0 * pf * Aused * eta)
            pMarginBar = pBar - pReqBar
            pReqOk = True
        End If

        ' Verdict
        Dim verdict As String = "PASS"
        If usagePct > 100.0 Then
            verdict = "FAIL"
            AddError("VERDICT: FAIL (usage > 100%). Increase bore / pressure or reduce load.")
        ElseIf usagePct > 85.0 Then
            verdict = "BORDERLINE"
            AddWarn("VERDICT: BORDERLINE (usage > 85%). Consider next bore for reliability.")
        Else
            verdict = "PASS"
        End If

        If Fbase < 0 Then
            AddWarn("Gravity assists strongly (base force negative). Required force clamped to 0 N; consider speed control/braking.")
        End If

        ' Required area from Frequired
        Dim Areq As Double = 0.0
        If pEff > 0 AndAlso eta > 0 Then Areq = Frequired / (pEff * eta)
        If Areq <= 0 Then
            AddError("Could not compute required area. Check pressure/efficiency.")
            SetStatusError("Status: Error") : Return
        End If

        ' Required bore for EXTEND: Abore >= Areq
        Dim DreqExtendMm As Double = Math.Sqrt((4.0 * Areq) / Math.PI) * 1000.0
        Dim DstdExtendMm As Double = NextStandardBore(DreqExtendMm)

        ' Required bore for RETRACT: (Abore - Arod) >= Areq  => Abore >= Areq + Arod
        Dim DreqRetractMm As Double = 0.0
        Dim DstdRetractMm As Double = 0.0
        Dim AreqRetractTotal As Double = Areq + Arod
        If AreqRetractTotal > 0 Then
            DreqRetractMm = Math.Sqrt((4.0 * AreqRetractTotal) / Math.PI) * 1000.0
            DstdRetractMm = NextStandardBore(DreqRetractMm)
        End If

        Dim recSelectedStd As Double = If(isRetract, DstdRetractMm, DstdExtendMm)

        Dim actionShort As String = If(isRetract, "Retract", "Extend")
        Dim recSelectedText As String = If(recSelectedStd > 0, recSelectedStd.ToString("0") & " mm", "Above standard list (increase bore)")

        ' NEW: shop-floor friendly lines
        Dim stdBoreToOrder As String = ""
        If recSelectedStd > 0 Then
            stdBoreToOrder = "STANDARD BORE TO ORDER: ISO Ø" & recSelectedStd.ToString("0") & " (" & actionShort & ")"
        Else
            stdBoreToOrder = "STANDARD BORE TO ORDER: (Above standard list - increase bore)"
        End If

        Dim whyLine As String = "WHY: Ø" & boreMm.ToString("0") & " is " & verdict & " (" & usagePct.ToString("0") & "% usage)"

        ' Cycle time
        Dim cycleOn As Boolean = (chkCycle IsNot Nothing AndAlso chkCycle.Checked)
        Dim strokeMm As Double = 0.0
        Dim vExt As Double = 0.0
        Dim vRet As Double = 0.0
        Dim dwellExt As Double = 0.0
        Dim dwellRet As Double = 0.0
        Dim tExt As Double = 0.0
        Dim tRet As Double = 0.0
        Dim tTotal As Double = 0.0

        If cycleOn Then
            strokeMm = ReadD(txtStrokeMm, DEFAULT_STROKE_MM)
            vExt = ReadD(txtExtSpeed, DEFAULT_EXT_SPEED_MM_S)
            vRet = ReadD(txtRetSpeed, DEFAULT_RET_SPEED_MM_S)
            dwellExt = ReadD(txtDwellExt, DEFAULT_DWELL_EXT_S)
            dwellRet = ReadD(txtDwellRet, DEFAULT_DWELL_RET_S)

            If strokeMm <= 0 Then AddError("Cycle Time: Stroke must be > 0 mm.") : SetStatusError("Status: Error") : Return
            If vExt <= 0 Then AddError("Cycle Time: Extend Speed must be > 0 mm/s.") : SetStatusError("Status: Error") : Return
            If vRet <= 0 Then AddError("Cycle Time: Retract Speed must be > 0 mm/s.") : SetStatusError("Status: Error") : Return
            If dwellExt < 0 Then dwellExt = 0
            If dwellRet < 0 Then dwellRet = 0

            tExt = strokeMm / vExt
            tRet = strokeMm / vRet
            tTotal = tExt + dwellExt + tRet + dwellRet
        End If

        ' FINAL RESULTS
        Dim sbFinal As New StringBuilder()

        sbFinal.AppendLine(stdBoreToOrder)
        sbFinal.AppendLine(whyLine)
        sbFinal.AppendLine("")

        sbFinal.AppendLine("PASS/BORDERLINE/FAIL : " & verdict)
        sbFinal.AppendLine("Cylinder usage       : " & usagePct.ToString("0.0") & " %")
        sbFinal.AppendLine("")
        sbFinal.AppendLine("Required force (design)   : " & Frequired.ToString("0.0") & " N")
        sbFinal.AppendLine("Available force (entered) : " & Favailable.ToString("0.0") & " N")
        sbFinal.AppendLine("Margin (Avail-Req)        : " & (Favailable - Frequired).ToString("0.0") & " N")
        If pReqOk Then
            sbFinal.AppendLine("Required pressure @ entered bore : " & pReqBar.ToString("0.00") & " bar")
            sbFinal.AppendLine("Pressure margin (@ supply)       : " & pMarginBar.ToString("0.00") & " bar")
        End If
        sbFinal.AppendLine("")
        sbFinal.AppendLine("Recommended standard bore (selected action) : " & recSelectedText)
        sbFinal.AppendLine("Extend next standard bore  : " & If(DstdExtendMm > 0, DstdExtendMm.ToString("0") & " mm", "(above list)"))
        sbFinal.AppendLine("Retract next standard bore : " & If(DstdRetractMm > 0, DstdRetractMm.ToString("0") & " mm", "(above list)"))
        sbFinal.AppendLine("")
        sbFinal.AppendLine("Angle used                 : " & angleDeg.ToString("0.0") & " deg (from vertical-up)")
        sbFinal.AppendLine("Action                     : " & actionName)

        If rodAutoSuggested Then
            sbFinal.AppendLine("Rod used (auto-suggested)  : " & rodMm.ToString("0") & " mm")
        Else
            sbFinal.AppendLine("Rod used                   : " & rodMm.ToString("0") & " mm")
        End If

        If cycleOn Then
            sbFinal.AppendLine("")
            sbFinal.AppendLine("Cycle time:")
            sbFinal.AppendLine("  Extend time  : " & tExt.ToString("0.000") & " s")
            sbFinal.AppendLine("  Dwell (ext)  : " & dwellExt.ToString("0.000") & " s")
            sbFinal.AppendLine("  Retract time : " & tRet.ToString("0.000") & " s")
            sbFinal.AppendLine("  Dwell (ret)  : " & dwellRet.ToString("0.000") & " s")
            sbFinal.AppendLine("  TOTAL        : " & tTotal.ToString("0.000") & " s")
        End If

        txtFinal.Text = sbFinal.ToString()

        ' STEP-BY-STEP (with real figures)
        Dim sbCalc As New StringBuilder()
        sbCalc.AppendLine("============================================================")
        sbCalc.AppendLine("CALCULATION METHOD (STEP-BY-STEP)")
        sbCalc.AppendLine("============================================================")
        sbCalc.AppendLine("Inputs:")
        sbCalc.AppendLine("  m = " & mKg.ToString("0.###") & " kg")
        sbCalc.AppendLine("  mu = " & mu.ToString("0.###"))
        sbCalc.AppendLine("  a = " & accel.ToString("0.###") & " m/s^2")
        sbCalc.AppendLine("  theta = " & angleDeg.ToString("0.###") & " deg (from vertical-up)")
        sbCalc.AppendLine("  action = " & actionName)
        sbCalc.AppendLine("  bore = " & boreMm.ToString("0.###") & " mm")
        sbCalc.AppendLine("  rod = " & rodMm.ToString("0.###") & " mm")
        sbCalc.AppendLine("  P = " & pBar.ToString("0.###") & " bar")
        sbCalc.AppendLine("  PressureFactor = " & pFactorPct.ToString("0.###") & " %")
        sbCalc.AppendLine("  CylinderEff = " & cylEffPct.ToString("0.###") & " %")
        sbCalc.AppendLine("  PeakFactor = " & peakFactor.ToString("0.###"))
        sbCalc.AppendLine("  SafetyFactor = " & sf.ToString("0.###"))
        sbCalc.AppendLine("")

        sbCalc.AppendLine("1) Convert angle to radians:")
        sbCalc.AppendLine("   thetaRad = theta*PI/180 = " & angleDeg.ToString("0.###") & "*PI/180 = " & thetaRad.ToString("0.#####"))
        sbCalc.AppendLine("")

        sbCalc.AppendLine("2) Weight:")
        sbCalc.AppendLine("   Fg = m*g = " & mKg.ToString("0.###") & "*9.81 = " & Fg.ToString("0.###") & " N")
        sbCalc.AppendLine("")

        sbCalc.AppendLine("3) Gravity along motion axis (from vertical-up):")
        sbCalc.AppendLine("   Fgravity = Fg*cos(thetaRad) = " & Fg.ToString("0.###") & "*cos(" & thetaRad.ToString("0.#####") & ") = " & FgravityAlong.ToString("0.###") & " N")
        sbCalc.AppendLine("")

        sbCalc.AppendLine("4) Normal force & friction:")
        sbCalc.AppendLine("   N = |Fg*sin(thetaRad)| = |" & Fg.ToString("0.###") & "*sin(" & thetaRad.ToString("0.#####") & ")| = " & N.ToString("0.###") & " N")
        sbCalc.AppendLine("   Ffric = mu*N = " & mu.ToString("0.###") & "*" & N.ToString("0.###") & " = " & Ffric.ToString("0.###") & " N")
        sbCalc.AppendLine("")

        sbCalc.AppendLine("5) Inertia force:")
        sbCalc.AppendLine("   Facc = m*a = " & mKg.ToString("0.###") & "*" & accel.ToString("0.###") & " = " & Facc.ToString("0.###") & " N")
        sbCalc.AppendLine("")

        sbCalc.AppendLine("6) Base force (clamped):")
        sbCalc.AppendLine("   Fbase = Fgravity + Ffric + Facc = " & FgravityAlong.ToString("0.###") & " + " & Ffric.ToString("0.###") & " + " & Facc.ToString("0.###") & " = " & Fbase.ToString("0.###") & " N")
        sbCalc.AppendLine("   FbaseClamped = max(0, Fbase) = " & FbaseClamped.ToString("0.###") & " N")
        sbCalc.AppendLine("")

        sbCalc.AppendLine("7) Peak + Safety:")
        sbCalc.AppendLine("   Fpeak = FbaseClamped*PeakFactor = " & FbaseClamped.ToString("0.###") & "*" & peakFactor.ToString("0.###") & " = " & Fpeak.ToString("0.###") & " N")
        sbCalc.AppendLine("   Frequired = Fpeak*SF = " & Fpeak.ToString("0.###") & "*" & sf.ToString("0.###") & " = " & Frequired.ToString("0.###") & " N")
        sbCalc.AppendLine("")

        sbCalc.AppendLine("8) Effective pressure:")
        sbCalc.AppendLine("   P(Pa) = Pbar*100000 = " & pBar.ToString("0.###") & "*100000 = " & pPa.ToString("0") & " Pa")
        sbCalc.AppendLine("   pf = PressureFactor/100 = " & pFactorPct.ToString("0.###") & "/100 = " & pf.ToString("0.###"))
        sbCalc.AppendLine("   eta = CylinderEff/100 = " & cylEffPct.ToString("0.###") & "/100 = " & eta.ToString("0.###"))
        sbCalc.AppendLine("   Peff = P*pf = " & pPa.ToString("0") & "*" & pf.ToString("0.###") & " = " & pEff.ToString("0") & " Pa")
        sbCalc.AppendLine("")

        sbCalc.AppendLine("9) Areas:")
        sbCalc.AppendLine("   bore(m) = bore(mm)/1000 = " & boreMm.ToString("0.###") & "/1000 = " & boreM.ToString("0.#####") & " m")
        sbCalc.AppendLine("   rod(m)  = rod(mm)/1000  = " & rodMm.ToString("0.###") & "/1000 = " & rodM.ToString("0.#####") & " m")
        sbCalc.AppendLine("   Abore = pi*bore^2/4 = " & Abore.ToString("0.########") & " m^2")
        sbCalc.AppendLine("   Arod  = pi*rod^2/4  = " & Arod.ToString("0.########") & " m^2")
        sbCalc.AppendLine("   Aann  = Abore-Arod  = " & Aann.ToString("0.########") & " m^2")
        sbCalc.AppendLine("")

        sbCalc.AppendLine("10) Force available:")
        sbCalc.AppendLine("   If Extend: Aused = Abore")
        sbCalc.AppendLine("   If Retract: Aused = Aann")
        sbCalc.AppendLine("   Favailable = Peff*Aused*eta = " & pEff.ToString("0") & "*" & Aused.ToString("0.########") & "*" & eta.ToString("0.###") & " = " & Favailable.ToString("0.###") & " N")
        sbCalc.AppendLine("")

        sbCalc.AppendLine("11) Usage:")
        sbCalc.AppendLine("   usage% = Frequired/Favailable*100 = " & Frequired.ToString("0.###") & "/" & Favailable.ToString("0.###") & "*100 = " & usagePct.ToString("0.###") & " %")
        sbCalc.AppendLine("")

        If pReqOk Then
            sbCalc.AppendLine("12) Required pressure @ entered bore:")
            sbCalc.AppendLine("   Preq(bar) = Frequired / (100000*pf*Aused*eta) = " & pReqBar.ToString("0.###") & " bar")
            sbCalc.AppendLine("   Margin(bar) = Psupply - Preq = " & pBar.ToString("0.###") & " - " & pReqBar.ToString("0.###") & " = " & pMarginBar.ToString("0.###") & " bar")
            sbCalc.AppendLine("")
        End If

        sbCalc.AppendLine("13) Recommended standard bore:")
        sbCalc.AppendLine("   Areq = Frequired/(Peff*eta) = " & Areq.ToString("0.########") & " m^2")
        sbCalc.AppendLine("   Extend: Dreq = sqrt(4*Areq/pi) = " & DreqExtendMm.ToString("0.###") & " mm -> next standard = " & If(DstdExtendMm > 0, DstdExtendMm.ToString("0") & " mm", "(above list)"))
        sbCalc.AppendLine("   Retract: Dreq = sqrt(4*(Areq + Arod)/pi) = " & DreqRetractMm.ToString("0.###") & " mm -> next standard = " & If(DstdRetractMm > 0, DstdRetractMm.ToString("0") & " mm", "(above list)"))
        sbCalc.AppendLine("   SHOP FLOOR: " & stdBoreToOrder)
        sbCalc.AppendLine("   SHOP FLOOR: " & whyLine)
        sbCalc.AppendLine("")

        If cycleOn Then
            sbCalc.AppendLine("14) Cycle time:")
            sbCalc.AppendLine("   tExt = stroke/vExt = " & strokeMm.ToString("0.###") & "/" & vExt.ToString("0.###") & " = " & tExt.ToString("0.###") & " s")
            sbCalc.AppendLine("   tRet = stroke/vRet = " & strokeMm.ToString("0.###") & "/" & vRet.ToString("0.###") & " = " & tRet.ToString("0.###") & " s")
            sbCalc.AppendLine("   tTotal = tExt + dwellExt + tRet + dwellRet = " & tExt.ToString("0.###") & " + " & dwellExt.ToString("0.###") & " + " & tRet.ToString("0.###") & " + " & dwellRet.ToString("0.###") & " = " & tTotal.ToString("0.###") & " s")
            sbCalc.AppendLine("")
        End If

        sbCalc.AppendLine("============================================================")
        txtCalc.Text = sbCalc.ToString()

        lblStatus.Text = "Status: Calculation successful"
        If verdict = "PASS" Then
            lblStatus.ForeColor = themeOk
        ElseIf verdict = "BORDERLINE" Then
            lblStatus.ForeColor = themeWarn
        Else
            lblStatus.ForeColor = themeErr
        End If
    End Sub

    ' ============================================================
    ' ANGLE HELPER
    ' ============================================================
    Private Function GetAngleDeg() As Double
        If cmbAnglePreset Is Nothing Then Return 90.0
        Dim s As String = cmbAnglePreset.Text
        If s Is Nothing Then Return 90.0

        If s.StartsWith("0", StringComparison.OrdinalIgnoreCase) Then Return 0.0
        If s.StartsWith("45", StringComparison.OrdinalIgnoreCase) Then Return 45.0
        If s.StartsWith("90", StringComparison.OrdinalIgnoreCase) Then Return 90.0
        If s.StartsWith("180", StringComparison.OrdinalIgnoreCase) Then Return 180.0

        Dim a As Double = ReadD(txtAngleCustom, 90.0)
        If a < 0 Then a = 0
        If a > 180 Then a = 180
        Return a
    End Function

    ' ============================================================
    ' STANDARD BORE HELPER
    ' ============================================================
    Private Function NextStandardBore(requiredMm As Double) As Double
        If requiredMm <= 0 Then Return 0
        For i As Integer = 0 To STANDARD_BORES_MM.Length - 1
            If STANDARD_BORES_MM(i) >= requiredMm Then Return STANDARD_BORES_MM(i)
        Next
        Return 0
    End Function

    ' ============================================================
    ' ROD AUTO-SUGGEST (Option Strict safe)
    ' ============================================================
    Private Function SuggestRodForBore(boreMm As Double) As Double
        Dim b As Integer = CInt(Math.Round(boreMm))
        If b <= 0 Then Return DEFAULT_ROD_MM

        Dim nearest As Integer = 32
        Dim bestDiff As Integer = Integer.MaxValue

        For Each k As Integer In ROD_SUGGEST_TABLE.Keys
            Dim d As Integer = Math.Abs(k - b)
            If d < bestDiff Then
                bestDiff = d
                nearest = k
            End If
        Next

        Dim rod As Integer = CInt(Math.Round(DEFAULT_ROD_MM))
        If ROD_SUGGEST_TABLE.ContainsKey(nearest) Then
            rod = ROD_SUGGEST_TABLE(nearest)
        End If

        Return CDbl(rod)
    End Function

    ' ============================================================
    ' WARNINGS / ERRORS
    ' ============================================================
    Private Sub ClearOutputs()
        If txtFinal IsNot Nothing Then txtFinal.Clear()
        If txtCalc IsNot Nothing Then txtCalc.Clear()
        If rtbWarn IsNot Nothing Then rtbWarn.Clear()
        If lblStatus IsNot Nothing Then
            lblStatus.Text = "Status: Ready"
            lblStatus.ForeColor = themeOk
        End If
    End Sub

    Private Sub AddWarn(msg As String)
        AppendWarnLine("WARNING: " & msg, themeWarn, False)
    End Sub

    Private Sub AddError(msg As String)
        AppendWarnLine("ERROR: " & msg, themeErr, True)
    End Sub

    Private Sub AppendWarnLine(text As String, col As Color, bold As Boolean)
        If rtbWarn Is Nothing Then Exit Sub
        Try
            rtbWarn.SelectionStart = rtbWarn.TextLength
            rtbWarn.SelectionLength = 0

            Dim f As Font = rtbWarn.Font
            If bold Then f = New Font(rtbWarn.Font, FontStyle.Bold)

            rtbWarn.SelectionFont = f
            rtbWarn.SelectionColor = col
            rtbWarn.AppendText(text & vbCrLf)

            rtbWarn.SelectionColor = themeText
            rtbWarn.SelectionFont = New Font(rtbWarn.Font, FontStyle.Regular)
        Catch
        End Try
    End Sub

    Private Sub SetStatusError(text As String)
        If lblStatus Is Nothing Then Exit Sub
        lblStatus.Text = text
        lblStatus.ForeColor = themeErr
    End Sub

    ' ============================================================
    ' PDF EXPORT
    ' ============================================================
    Private Sub ExportPdf(sender As Object, e As EventArgs)
        Dim sfd As New SaveFileDialog()
        sfd.Filter = "PDF Files|*.pdf"
        If sfd.ShowDialog() <> DialogResult.OK Then Exit Sub

        Dim doc As New PrintDocument()
        AddHandler doc.PrintPage,
            Sub(s, ev)
                ev.Graphics.DrawString("Pneumatic Cylinder Calculation Report",
                                       New Font("Segoe UI", 14, FontStyle.Bold),
                                       Brushes.Black, 50, 50)

                Dim body As String = ""
                body &= txtFinal.Text & vbCrLf & vbCrLf
                body &= txtCalc.Text & vbCrLf & vbCrLf
                If rtbWarn IsNot Nothing AndAlso rtbWarn.TextLength > 0 Then body &= rtbWarn.Text

                ev.Graphics.DrawString(body, New Font("Consolas", 9), Brushes.Black, 50, 100)
            End Sub

        doc.PrinterSettings.PrintToFile = True
        doc.PrinterSettings.PrintFileName = sfd.FileName

        Try
            doc.Print()
        Catch ex As Exception
            AddError("PDF export failed: " & ex.Message & " (PDF printer may be missing).")
            SetStatusError("Status: PDF export error")
        End Try
    End Sub

    ' ============================================================
    ' RESET (blank everything + DO NOT change theme)
    ' ============================================================
    Private Sub ResetAll(sender As Object, e As EventArgs)
        SyncThemeFromCurrentUI()
        ClearOutputs()

        ClearTB(txtMassKg)
        ClearTB(txtMu)
        ClearTB(txtAccel)

        SetComboSafe(cmbAnglePreset, 2) ' 90 deg
        ClearTB(txtAngleCustom)

        SetComboSafe(cmbAction, 0)

        ClearTB(txtPressureBar)
        ClearTB(txtPressureFactor)
        ClearTB(txtCylEff)

        ClearTB(txtBoreMm)
        ClearTB(txtRodMm)

        ClearTB(txtPeakFactor)
        ClearTB(txtSF)

        SetCheckSafe(chkCycle, False)
        ClearTB(txtStrokeMm)
        ClearTB(txtExtSpeed)
        ClearTB(txtRetSpeed)
        ClearTB(txtDwellExt)
        ClearTB(txtDwellRet)

        ToggleCycle(Nothing, EventArgs.Empty)
        AnglePresetChanged(Nothing, EventArgs.Empty)

        If lblStatus IsNot Nothing Then
            lblStatus.Text = "Status: Reset (inputs cleared)"
            lblStatus.ForeColor = themeWarn
        End If
    End Sub

    Private Sub ClearTB(tb As TextBox)
        If tb Is Nothing Then Exit Sub
        Try
            tb.Text = ""
        Catch
        End Try
    End Sub

    Private Sub SetComboSafe(cb As ComboBox, index As Integer)
        If cb Is Nothing Then Exit Sub
        Try
            If cb.Items IsNot Nothing AndAlso cb.Items.Count > 0 Then
                If index < 0 Then index = 0
                If index > cb.Items.Count - 1 Then index = cb.Items.Count - 1
                cb.SelectedIndex = index
            End If
        Catch
        End Try
    End Sub

    Private Sub SetCheckSafe(chk As CheckBox, value As Boolean)
        If chk Is Nothing Then Exit Sub
        Try
            chk.Checked = value
        Catch
        End Try
    End Sub

    ' ============================================================
    ' TOOLTIP SYSTEM
    ' ============================================================
    Private Sub RegisterTip(c As Control, text As String)
        If c Is Nothing Then Exit Sub
        If tipMap Is Nothing Then Exit Sub

        Dim wrapped As String = WrapTip(text, 58)

        If tipMap.ContainsKey(c) Then
            tipMap(c) = wrapped
        Else
            tipMap.Add(c, wrapped)
        End If

        AddHandler c.Enter, AddressOf ControlEnter_ShowTip
        AddHandler c.Leave, AddressOf ControlLeave_HideTip
    End Sub

    Private Sub ControlEnter_ShowTip(sender As Object, e As EventArgs)
        If Not tipsArmed Then Exit Sub

        Dim c As Control = TryCast(sender, Control)
        If c Is Nothing Then Exit Sub
        If tipMap Is Nothing Then Exit Sub
        If Not tipMap.ContainsKey(c) Then Exit Sub

        Try
            tip.Hide(c)
            Dim x As Integer = Math.Max(10, c.Width + 6)
            tip.Show(tipMap(c), c, x, 0, 9000)
        Catch
        End Try
    End Sub

    Private Sub ControlLeave_HideTip(sender As Object, e As EventArgs)
        Dim c As Control = TryCast(sender, Control)
        If c Is Nothing Then Exit Sub
        Try
            tip.Hide(c)
        Catch
        End Try
    End Sub

    Private Function WrapTip(text As String, maxLine As Integer) As String
        If text Is Nothing Then Return ""
        Dim s As String = text.Trim()
        If s = "" Then Return ""

        Dim words() As String = s.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
        Dim sb As New StringBuilder()
        Dim line As New StringBuilder()

        For i As Integer = 0 To words.Length - 1
            Dim w As String = words(i)
            If line.Length = 0 Then
                line.Append(w)
            ElseIf line.Length + 1 + w.Length <= maxLine Then
                line.Append(" "c)
                line.Append(w)
            Else
                sb.AppendLine(line.ToString())
                line.Length = 0
                line.Append(w)
            End If
        Next

        If line.Length > 0 Then sb.Append(line.ToString())
        Return sb.ToString()
    End Function

    ' ============================================================
    ' NUMERIC READ (blank => defaultValue)
    ' ============================================================
    Private Function ReadD(tb As TextBox, defaultValue As Double) As Double
        If tb Is Nothing Then Return defaultValue
        Dim s As String = tb.Text
        If s Is Nothing Then Return defaultValue
        s = s.Trim()
        If s = "" Then Return defaultValue

        Dim v As Double = defaultValue
        Try
            s = s.Replace(","c, "."c)
            If Double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, v) Then
                Return v
            End If
        Catch
        End Try

        Return defaultValue
    End Function

End Class
