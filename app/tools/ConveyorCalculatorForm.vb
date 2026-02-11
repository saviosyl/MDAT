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

Public Class ConveyorCalculatorForm
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
    Private Const DEFAULT_LOADVAR As Double = 1.0
    Private Const DEFAULT_DRIVE_DIA_MM As Double = 100.0
    Private Const DEFAULT_EFF_PCT As Double = 85.0
    Private Const DEFAULT_SF As Double = 1.2
    Private Const DEFAULT_START_FACTOR As Double = 1.5
    Private Const DEFAULT_MOTOR_RPM As Double = 1400.0
    Private Const DEFAULT_ACCEL_TIME As Double = 1.0
    Private Const DEFAULT_KG_PER_M As Double = 2.0
    Private Const DEFAULT_STARTS_PER_HR As Double = 60.0

    ' ============================================================
    ' PANELS
    ' ============================================================
    Private pnlInput As Panel
    Private pnlResults As Panel

    ' ============================================================
    ' OUTPUT
    ' ============================================================
    Private txtFinal As TextBox
    Private txtCalc As TextBox
    Private rtbWarn As RichTextBox
    Private lblStatus As Label

    ' ============================================================
    ' INPUTS (BASIC)
    ' ============================================================
    Private txtConvLen As TextBox
    Private txtProdLen As TextBox
    Private txtGap As TextBox
    Private txtWeight As TextBox
    Private txtMu As TextBox
    Private txtPPM As TextBox
    Private txtAccLen As TextBox
    Private txtIncline As TextBox

    Private cmbType As ComboBox
    Private cmbBed As ComboBox

    ' Operating conditions / design factors
    Private cmbEnv As ComboBox
    Private txtDuty As TextBox
    Private cmbMode As ComboBox
    Private txtLoadVar As TextBox

    ' Design Factors
    Private chkDesign As CheckBox
    Private txtRollerDia As TextBox
    Private txtEff As TextBox
    Private txtSF As TextBox
    Private txtStartFactor As TextBox
    Private txtMotorRPM As TextBox

    ' Accuracy / dynamics
    Private chkAccuracy As CheckBox
    Private txtConvMassPerM As TextBox
    Private txtAccelTime As TextBox
    Private txtStartsPerHr As TextBox

    ' Expert inertia (optional)
    Private chkInertia As CheckBox
    Private cmbShaftMat As ComboBox
    Private txtDensity As TextBox
    Private txtShaftDia As TextBox
    Private txtShaftLen As TextBox
    Private txtSprocketInertia As TextBox
    Private txtExtraInertia As TextBox

    ' Buttons
    Private btnCalc As Button
    Private btnPdf As Button
    Private btnReset As Button

    ' Tooltips
    Private tip As ToolTip
    Private tipMap As Dictionary(Of Control, String)
    Private tipsArmed As Boolean = False ' prevents tooltip popping instantly on form open

    ' ============================================================
    ' INIT
    ' ============================================================
    Public Sub New()
        Me.Text = "Conveyor Calculator"
        Me.Size = New Size(1180, 720)
        Me.StartPosition = FormStartPosition.CenterParent

        ' Start with defaults (dark) if opened standalone
        Me.BackColor = themeBG

        BuildUI()

        ' Do NOT show any tooltip until user clicks inside the form
        AddHandler Me.MouseDown, AddressOf ArmTipsOnFirstClick

        ' Apply initial theme using current colors
        SyncThemeFromCurrentUI()
        ApplyThemeToThisForm()
    End Sub

    Private Sub ArmTipsOnFirstClick(sender As Object, e As MouseEventArgs)
        tipsArmed = True
    End Sub

    ' ============================================================
    ' MDAT THEME HOOK (called from MainForm before showing)
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

    ' ============================================================
    ' CRITICAL: SYNC THEME FROM CURRENT UI COLORS
    ' This prevents Reset/Calculate from repainting with dark defaults.
    ' ============================================================
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
        tip.AutoPopDelay = 12000
        tip.InitialDelay = 200
        tip.ReshowDelay = 50

        tipMap = New Dictionary(Of Control, String)()

        pnlInput = New Panel()
        pnlInput.Location = New Point(10, 10)
        pnlInput.Size = New Size(360, 660)
        pnlInput.BackColor = themePanel
        pnlInput.AutoScroll = True
        Me.Controls.Add(pnlInput)

        pnlResults = New Panel()
        pnlResults.Location = New Point(390, 10)
        pnlResults.Size = New Size(770, 660)
        pnlResults.BackColor = themeBG
        Me.Controls.Add(pnlResults)

        Dim y As Integer = 15

        AddSection("INPUT", y) : y += 26

        ' All numeric fields EMPTY by default
        txtConvLen = AddInput("Conveyor Length (mm)", y, "Total conveyor length carrying load (mm).") : y += 26
        txtProdLen = AddInput("Product Length (mm)", y, "Product length along travel direction (mm).") : y += 26
        txtGap = AddInput("Safety Gap (mm)", y, "Minimum spacing between products (mm).") : y += 26
        txtWeight = AddInput("Product Weight (g)", y, "Weight of ONE product in grams (g).") : y += 26

        AddLabel("Conveyor Type", y)
        cmbType = AddCombo(y, New String() {"Belt", "Roller", "Modular Chain"}, "Select conveyor type.")
        AddHandler cmbType.SelectedIndexChanged, AddressOf ConveyorTypeChanged
        y += 26

        AddLabel("Belt Bed Material", y)
        cmbBed = AddCombo(y, New String() {"Steel Bed", "UHMWPE", "Stainless Steel"}, "For belt conveyors: bed affects μ.")
        AddHandler cmbBed.SelectedIndexChanged, AddressOf BedMaterialChanged
        y += 26

        txtMu = AddInput("Friction Coefficient (μ)", y, "Friction coefficient μ. Auto-presets for Belt/Bed only if blank.") : y += 26
        txtPPM = AddInput("Target Products / Min", y, "Throughput (products/min). Determines speed.") : y += 26
        txtAccLen = AddInput("Accumulated Length (mm)", y, "Accumulated/backed-up length (mm). Adds moving mass.") : y += 26
        txtIncline = AddInput("Incline Angle (deg)", y, "Incline angle in degrees. 0 = horizontal.") : y += 30

        AddSection("OPERATING CONDITIONS", y) : y += 26

        AddLabel("Environment", y)
        cmbEnv = AddCombo(y, New String() {"Clean", "Dusty", "Wet", "Washdown"}, "Environment increases resistance margin.")
        y += 26

        txtDuty = AddInput("Duty Cycle (%)", y, "Approx. % of time running (informational).") : y += 26

        AddLabel("Operation Mode", y)
        cmbMode = AddCombo(y, New String() {"Continuous", "Intermittent"}, "Intermittent = many starts/stops.")
        y += 26

        txtLoadVar = AddInput("Load Variability Factor", y, "Multiplier for uneven loading. Blank => default 1.0.") : y += 30

        AddSection("DESIGN FACTORS (RECOMMENDED)", y) : y += 26

        chkDesign = AddCheck("Enable Design Factors", y, "Adds drive dia, efficiency, SF, start factor, motor RPM.")
        AddHandler chkDesign.CheckedChanged, AddressOf ToggleDesignFactors
        y += 24

        txtRollerDia = AddInput("Drive Roller Dia (mm)", y, "Drive roller / sprocket pitch dia. Blank => 100 mm.") : y += 24
        txtEff = AddInput("Drive Efficiency (%)", y, "Overall efficiency. Blank => 85%.") : y += 24
        txtSF = AddInput("Safety Factor", y, "Safety factor applied to torque/power. Blank => 1.2.") : y += 24
        txtStartFactor = AddInput("Start / Accel Factor", y, "Peak torque multiplier when inertia unknown. Blank => 1.5.") : y += 24
        txtMotorRPM = AddInput("Motor RPM", y, "Motor rated speed. Blank => 1400 rpm.") : y += 30

        AddSection("ACCURACY MODE (OPTIONAL)", y) : y += 26

        chkAccuracy = AddCheck("Enable Accuracy Mode", y, "Adds conveyor mass per meter + accel time + starts/hour.")
        AddHandler chkAccuracy.CheckedChanged, AddressOf ToggleAccuracy
        y += 24

        txtConvMassPerM = AddInput("Conveyor Mass / meter (kg/m)", y, "Moving mass (belt/chain + sliders). Blank => 2.0 kg/m.") : y += 24
        txtAccelTime = AddInput("Acceleration Time (s)", y, "Ramp time 0→speed. Blank => 1.0 s.") : y += 24
        txtStartsPerHr = AddInput("Starts per Hour", y, "For intermittent check. Blank => 60.") : y += 30

        AddSection("EXPERT: INERTIA (OPTIONAL)", y) : y += 26

        chkInertia = AddCheck("I have inertia values (expert override)", y, "Enable only if you can estimate inertia J (kg·m²).")
        AddHandler chkInertia.CheckedChanged, AddressOf ToggleInertia
        y += 24

        AddLabel("Shaft Material", y)
        cmbShaftMat = AddCombo(y, New String() {"Mild Steel (7850)", "Stainless Steel (8000)", "Aluminium (2700)", "Cast Iron (7200)", "Plastic (1100)", "Custom"}, "Auto-fill density.")
        AddHandler cmbShaftMat.SelectedIndexChanged, AddressOf ShaftMaterialChanged
        y += 26

        txtDensity = AddInput("Density (kg/m³)", y, "Auto-filled from material. Custom allows manual input.") : y += 24
        txtShaftDia = AddInput("Shaft Dia (mm)", y, "Solid shaft diameter (mm) for inertia estimate.") : y += 24
        txtShaftLen = AddInput("Shaft Length (mm)", y, "Solid shaft length (mm) for inertia estimate.") : y += 24
        txtSprocketInertia = AddInput("Sprocket/Hub Inertia J (kg·m²)", y, "If known from CAD/supplier.") : y += 24
        txtExtraInertia = AddInput("Extra Inertia J (kg·m²)", y, "Couplings, etc. if known.") : y += 30

        btnCalc = CreateButton("Calculate", themeAccent, New Point(20, y), True, AddressOf Calculate)
        btnPdf = CreateButton("Export PDF", btnGreen, New Point(130, y), False, AddressOf ExportPdf)
        btnReset = CreateButton("Reset", btnYellow, New Point(240, y), False, AddressOf ResetAll)

        pnlInput.Controls.Add(btnCalc)
        pnlInput.Controls.Add(btnPdf)
        pnlInput.Controls.Add(btnReset)

        txtFinal = AddOutputText("FINAL RESULTS", 10, 200)
        txtCalc = AddOutputText("CALCULATION METHOD (STEP-BY-STEP)", 240, 200)

        AddOutputLabel("WARNINGS / ERRORS", 470)
        rtbWarn = New RichTextBox()
        rtbWarn.Location = New Point(0, 490)
        rtbWarn.Size = New Size(760, 170)
        rtbWarn.ReadOnly = True
        rtbWarn.BorderStyle = BorderStyle.FixedSingle
        rtbWarn.Multiline = True
        rtbWarn.ScrollBars = RichTextBoxScrollBars.Vertical
        pnlResults.Controls.Add(rtbWarn)

        lblStatus = New Label()
        lblStatus.Text = "Status: Ready"
        lblStatus.ForeColor = themeOk
        lblStatus.Location = New Point(390, 680)
        lblStatus.AutoSize = True
        Me.Controls.Add(lblStatus)

        ' Defaults for toggles only
        chkDesign.Checked = True
        chkAccuracy.Checked = False
        chkInertia.Checked = False

        ToggleDesignFactors(Nothing, EventArgs.Empty)
        ToggleAccuracy(Nothing, EventArgs.Empty)
        ToggleInertia(Nothing, EventArgs.Empty)

        ConveyorTypeChanged(Nothing, EventArgs.Empty)

        If cmbShaftMat IsNot Nothing AndAlso cmbShaftMat.Items.Count > 0 Then cmbShaftMat.SelectedIndex = 0
        ShaftMaterialChanged(Nothing, EventArgs.Empty)
    End Sub

    ' ============================================================
    ' APPLY THEME (stable) - uses theme variables
    ' ============================================================
    Private Sub ApplyThemeToThisForm()
        Try
            Me.BackColor = themeBG
        Catch
        End Try

        Try
            If pnlInput IsNot Nothing Then pnlInput.BackColor = themePanel
            If pnlResults IsNot Nothing Then pnlResults.BackColor = themeBG
        Catch
        End Try

        Try
            ApplyThemeRecursive(Me)
        Catch
        End Try

        ' Buttons keep meaning colors
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

            ElseIf TypeOf c Is Label Then
                Dim lbl As Label = DirectCast(c, Label)
                lbl.BackColor = Color.Transparent
                If lbl.ForeColor = themeAccent Then
                    lbl.ForeColor = themeAccent
                Else
                    lbl.ForeColor = themeText
                End If

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
    ' AUTO MU (belt presets) - only if μ is blank
    ' ============================================================
    Private Sub ConveyorTypeChanged(sender As Object, e As EventArgs)
        Dim isBelt As Boolean = (cmbType IsNot Nothing AndAlso cmbType.Text = "Belt")
        If cmbBed IsNot Nothing Then cmbBed.Enabled = isBelt
        BedMaterialChanged(Nothing, EventArgs.Empty)
    End Sub

    Private Sub BedMaterialChanged(sender As Object, e As EventArgs)
        If cmbType Is Nothing OrElse txtMu Is Nothing OrElse cmbBed Is Nothing Then Exit Sub
        If cmbType.Text <> "Belt" Then Exit Sub

        If txtMu.Text Is Nothing OrElse txtMu.Text.Trim() = "" Then
            If cmbBed.Text = "Steel Bed" Then txtMu.Text = "0.30"
            If cmbBed.Text = "UHMWPE" Then txtMu.Text = "0.20"
            If cmbBed.Text = "Stainless Steel" Then txtMu.Text = "0.25"
        End If
    End Sub

    ' ============================================================
    ' TOGGLES
    ' ============================================================
    Private Sub ToggleDesignFactors(sender As Object, e As EventArgs)
        Dim en As Boolean = (chkDesign IsNot Nothing AndAlso chkDesign.Checked)
        SetEnabled(txtRollerDia, en)
        SetEnabled(txtEff, en)
        SetEnabled(txtSF, en)
        SetEnabled(txtStartFactor, en)
        SetEnabled(txtMotorRPM, en)
    End Sub

    Private Sub ToggleAccuracy(sender As Object, e As EventArgs)
        Dim en As Boolean = (chkAccuracy IsNot Nothing AndAlso chkAccuracy.Checked)
        SetEnabled(txtConvMassPerM, en)
        SetEnabled(txtAccelTime, en)
        SetEnabled(txtStartsPerHr, en)
    End Sub

    Private Sub ToggleInertia(sender As Object, e As EventArgs)
        Dim en As Boolean = (chkInertia IsNot Nothing AndAlso chkInertia.Checked)
        SetEnabled(cmbShaftMat, en)
        SetEnabled(txtDensity, en)
        SetEnabled(txtShaftDia, en)
        SetEnabled(txtShaftLen, en)
        SetEnabled(txtSprocketInertia, en)
        SetEnabled(txtExtraInertia, en)
    End Sub

    Private Sub SetEnabled(c As Control, en As Boolean)
        If c Is Nothing Then Exit Sub
        Try
            c.Enabled = en
        Catch
        End Try
    End Sub

    ' ============================================================
    ' SHAFT MATERIAL -> DENSITY (only auto-fill if blank)
    ' ============================================================
    Private Sub ShaftMaterialChanged(sender As Object, e As EventArgs)
        If cmbShaftMat Is Nothing OrElse txtDensity Is Nothing Then Exit Sub

        Dim sel As String = cmbShaftMat.Text
        Dim dens As Double = 7850.0
        Dim allowEdit As Boolean = False

        If sel.IndexOf("7850", StringComparison.OrdinalIgnoreCase) >= 0 Then dens = 7850.0
        If sel.IndexOf("8000", StringComparison.OrdinalIgnoreCase) >= 0 Then dens = 8000.0
        If sel.IndexOf("2700", StringComparison.OrdinalIgnoreCase) >= 0 Then dens = 2700.0
        If sel.IndexOf("7200", StringComparison.OrdinalIgnoreCase) >= 0 Then dens = 7200.0
        If sel.IndexOf("1100", StringComparison.OrdinalIgnoreCase) >= 0 Then dens = 1100.0
        If sel = "Custom" Then allowEdit = True

        txtDensity.ReadOnly = Not allowEdit

        If Not allowEdit Then
            If txtDensity.Text Is Nothing OrElse txtDensity.Text.Trim() = "" Then
                txtDensity.Text = dens.ToString("0", CultureInfo.InvariantCulture)
            End If
        End If
    End Sub

    ' ============================================================
    ' CALCULATE (keeps current theme - no forced dark repaint)
    ' ============================================================
    Private Sub Calculate(sender As Object, e As EventArgs)
        ' IMPORTANT: do not repaint with defaults
        SyncThemeFromCurrentUI()

        ClearOutputs()

        Dim Lmm As Double = ReadD(txtConvLen, 0.0)
        Dim prodMm As Double = ReadD(txtProdLen, 0.0)
        Dim gapMm As Double = ReadD(txtGap, 0.0)
        Dim ppm As Double = ReadD(txtPPM, 0.0)
        Dim weightG As Double = ReadD(txtWeight, 0.0)
        Dim accMm As Double = ReadD(txtAccLen, 0.0)
        Dim muBase As Double = ReadD(txtMu, 0.0)
        Dim thetaDeg As Double = ReadD(txtIncline, 0.0)

        If Lmm <= 0 OrElse prodMm <= 0 OrElse ppm <= 0 OrElse weightG <= 0 Then
            AddError("Please enter: Conveyor Length, Product Length, PPM and Product Weight (all must be > 0).")
            SetStatusError("Status: Error")
            Return
        End If

        Dim pitchMm As Double = prodMm + gapMm
        If pitchMm <= 0 Then
            AddError("Invalid pitch. Product Length + Safety Gap must be > 0.")
            SetStatusError("Status: Error")
            Return
        End If

        If muBase <= 0 Then
            AddError("Please enter friction coefficient μ (> 0).")
            SetStatusError("Status: Error")
            Return
        End If

        Dim Lm As Double = Lmm / 1000.0
        Dim pitchM As Double = pitchMm / 1000.0
        Dim accM As Double = Math.Max(0.0, accMm / 1000.0)
        Dim theta As Double = thetaDeg * Math.PI / 180.0

        Dim countOnConv As Integer = CInt(Math.Floor(Lm / pitchM))
        Dim countAcc As Integer = CInt(Math.Floor(accM / pitchM))
        Dim totalCount As Integer = Math.Max(0, countOnConv + countAcc)

        If totalCount <= 0 Then
            AddError("Calculated total products on conveyor is 0. Check length/pitch/accumulation.")
            SetStatusError("Status: Error")
            Return
        End If

        Dim mProd As Double = weightG / 1000.0
        Dim mProductsTotal As Double = totalCount * mProd

        Dim envFactor As Double = 1.0
        Dim envName As String = If(cmbEnv Is Nothing, "Clean", cmbEnv.Text)
        If envName = "Dusty" Then envFactor = 1.15
        If envName = "Wet" Then envFactor = 1.3
        If envName = "Washdown" Then envFactor = 1.45

        Dim loadVar As Double = ReadD(txtLoadVar, DEFAULT_LOADVAR)
        If loadVar <= 0 Then loadVar = DEFAULT_LOADVAR

        Dim mu As Double = muBase * envFactor

        Dim accuracyOn As Boolean = (chkAccuracy IsNot Nothing AndAlso chkAccuracy.Checked)
        Dim effLenM As Double = Lm + accM
        Dim massPerM As Double = 0.0
        Dim mConveyorMoving As Double = 0.0

        If accuracyOn Then
            massPerM = ReadD(txtConvMassPerM, DEFAULT_KG_PER_M)
            If massPerM < 0 Then massPerM = 0
            mConveyorMoving = massPerM * effLenM
        End If

        Dim mMovingTotal As Double = mProductsTotal + mConveyorMoving

        Dim g As Double = 9.81
        Dim Ffric As Double = mu * mMovingTotal * g * Math.Cos(theta)
        Dim Fincline As Double = mMovingTotal * g * Math.Sin(theta)
        Dim Frun As Double = (Ffric + Fincline) * loadVar
        If Frun < 0 Then Frun = 0

        Dim speedMMin As Double = ppm * pitchM
        Dim speedMS As Double = speedMMin / 60.0

        If speedMMin > 65 Then AddWarn("Conveyor speed exceeds 65 m/min. Check feasibility, stability, and safety.")
        If thetaDeg > 20 Then AddWarn("Incline > 20°. Consider backstop, traction limits, and higher safety factor.")
        If muBase > 1.5 Then AddWarn("Friction μ looks unusually high. Please verify.")

        Dim designOn As Boolean = (chkDesign IsNot Nothing AndAlso chkDesign.Checked)

        Dim driveDiaMm As Double = ReadD(txtRollerDia, DEFAULT_DRIVE_DIA_MM)
        If driveDiaMm <= 0 Then driveDiaMm = DEFAULT_DRIVE_DIA_MM

        Dim effPct As Double = ReadD(txtEff, DEFAULT_EFF_PCT)
        Dim eff As Double = effPct / 100.0
        If eff <= 0 OrElse eff > 1 Then eff = DEFAULT_EFF_PCT / 100.0

        Dim sf As Double = ReadD(txtSF, DEFAULT_SF)
        If sf <= 0 Then sf = DEFAULT_SF

        Dim startFactor As Double = ReadD(txtStartFactor, DEFAULT_START_FACTOR)
        If startFactor <= 0 Then startFactor = DEFAULT_START_FACTOR

        Dim motorRPM As Double = ReadD(txtMotorRPM, DEFAULT_MOTOR_RPM)
        If motorRPM <= 0 Then motorRPM = DEFAULT_MOTOR_RPM

        If Not designOn Then
            driveDiaMm = DEFAULT_DRIVE_DIA_MM
            eff = DEFAULT_EFF_PCT / 100.0
            sf = DEFAULT_SF
            startFactor = DEFAULT_START_FACTOR
            motorRPM = DEFAULT_MOTOR_RPM
        End If

        Dim radius As Double = (driveDiaMm / 1000.0) / 2.0

        Dim torqueRun As Double = Frun * radius
        Dim torqueDesign As Double = torqueRun * sf
        Dim torquePeakStartFactor As Double = torqueDesign * startFactor

        Dim powerRunW As Double = (Frun * speedMS) / eff
        Dim powerDesignW As Double = powerRunW * sf

        Dim driveRPM As Double = 0.0
        If driveDiaMm > 0 Then
            Dim Dm As Double = driveDiaMm / 1000.0
            driveRPM = speedMS * 60.0 / (Math.PI * Dm)
        End If

        Dim gearRatio As Double = 0.0
        If driveRPM > 0 Then gearRatio = motorRPM / driveRPM

        Dim accelTimeS As Double = ReadD(txtAccelTime, DEFAULT_ACCEL_TIME)
        Dim startsHr As Double = ReadD(txtStartsPerHr, DEFAULT_STARTS_PER_HR)
        If accelTimeS < 0 Then accelTimeS = DEFAULT_ACCEL_TIME
        If startsHr < 0 Then startsHr = DEFAULT_STARTS_PER_HR

        Dim accelExtraTorque As Double = 0.0
        If accuracyOn AndAlso accelTimeS > 0 Then
            Dim a As Double = speedMS / accelTimeS
            Dim Facc As Double = mMovingTotal * a
            accelExtraTorque = (Facc * radius) * sf
            If cmbMode IsNot Nothing AndAlso cmbMode.Text = "Intermittent" AndAlso startsHr > 300 Then
                AddWarn("High starts/hour detected. Consider thermal limits and a higher service factor.")
            End If
        End If

        Dim inertiaOn As Boolean = (chkInertia IsNot Nothing AndAlso chkInertia.Checked)
        Dim Jtotal As Double = 0.0
        Dim torqueInertia As Double = 0.0

        If inertiaOn Then
            Dim dens As Double = ReadD(txtDensity, 7850.0)
            Dim shaftDiaMm As Double = ReadD(txtShaftDia, 0.0)
            Dim shaftLenMm As Double = ReadD(txtShaftLen, 0.0)

            Dim rS As Double = (shaftDiaMm / 1000.0) / 2.0
            Dim LS As Double = shaftLenMm / 1000.0
            Dim vol As Double = Math.PI * rS * rS * LS
            Dim mShaft As Double = dens * vol
            Dim Jshaft As Double = 0.5 * mShaft * rS * rS

            Dim Jsprocket As Double = ReadD(txtSprocketInertia, 0.0)
            Dim Jextra As Double = ReadD(txtExtraInertia, 0.0)

            Jtotal = Math.Max(0.0, Jshaft) + Math.Max(0.0, Jsprocket) + Math.Max(0.0, Jextra)

            If accelTimeS > 0 AndAlso driveRPM > 0 Then
                Dim omega As Double = (2.0 * Math.PI * driveRPM) / 60.0
                Dim alpha As Double = omega / accelTimeS
                torqueInertia = (Jtotal * alpha) * sf
            Else
                AddWarn("Inertia enabled, but Acceleration Time/Drive RPM missing. Inertia torque could not be computed.")
            End If
        End If

        Dim torquePeakSuggested As Double = torquePeakStartFactor
        If accelExtraTorque > 0 Then torquePeakSuggested = Math.Max(torquePeakSuggested, torqueDesign + accelExtraTorque)
        If torqueInertia > 0 Then torquePeakSuggested = Math.Max(torquePeakSuggested, torqueDesign + torqueInertia)

        ' FINAL RESULTS
        Dim sbFinal As New StringBuilder()
        sbFinal.AppendLine("Total Products             : " & totalCount.ToString())
        sbFinal.AppendLine("Product Total Mass         : " & mProductsTotal.ToString("0.00") & " kg")
        If accuracyOn Then
            sbFinal.AppendLine("Conveyor Moving Mass       : " & mConveyorMoving.ToString("0.00") & " kg  (kg/m=" & massPerM.ToString("0.00") & ", length=" & effLenM.ToString("0.00") & " m)")
        End If
        sbFinal.AppendLine("Total Moving Mass          : " & mMovingTotal.ToString("0.00") & " kg")
        sbFinal.AppendLine("Incline Angle              : " & thetaDeg.ToString("0.00") & " deg")
        sbFinal.AppendLine("Speed                      : " & speedMMin.ToString("0.00") & " m/min  (" & speedMS.ToString("0.000") & " m/s)")
        sbFinal.AppendLine("")
        sbFinal.AppendLine("Drive Force (running)      : " & Frun.ToString("0.0") & " N")
        sbFinal.AppendLine("Torque (running @ drive)   : " & torqueRun.ToString("0.00") & " Nm")
        sbFinal.AppendLine("Torque (design SF)         : " & torqueDesign.ToString("0.00") & " Nm")
        sbFinal.AppendLine("Peak Torque (suggested)    : " & torquePeakSuggested.ToString("0.00") & " Nm")
        sbFinal.AppendLine("")
        sbFinal.AppendLine("Power (running)            : " & powerRunW.ToString("0.0") & " W  (" & (powerRunW / 1000.0).ToString("0.000") & " kW)")
        sbFinal.AppendLine("Power (design SF)          : " & powerDesignW.ToString("0.0") & " W  (" & (powerDesignW / 1000.0).ToString("0.000") & " kW)")
        sbFinal.AppendLine("")
        sbFinal.AppendLine("Motor RPM (assumed)        : " & motorRPM.ToString("0") & " rpm")
        sbFinal.AppendLine("Drive Dia Used             : " & driveDiaMm.ToString("0.0") & " mm")
        sbFinal.AppendLine("Drive RPM (required)       : " & driveRPM.ToString("0.0") & " rpm")
        If gearRatio > 0 Then
            sbFinal.AppendLine("Suggested Gear Ratio       : " & gearRatio.ToString("0.0") & " : 1")
        Else
            sbFinal.AppendLine("Suggested Gear Ratio       : (need speed & drive dia)")
        End If
        txtFinal.Text = sbFinal.ToString()

        ' FULL STEP-BY-STEP
        Dim sbCalc As New StringBuilder()
        sbCalc.AppendLine("============================================================")
        sbCalc.AppendLine("CALCULATION METHOD (STEP-BY-STEP)")
        sbCalc.AppendLine("============================================================")
        sbCalc.AppendLine("NOTE: Results are indicative. Validate with supplier/ISO/CEMA data.")
        sbCalc.AppendLine("")
        sbCalc.AppendLine("1) pitch = product + gap = " & prodMm.ToString("0.###") & " + " & gapMm.ToString("0.###") & " = " & pitchMm.ToString("0.###") & " mm")
        sbCalc.AppendLine("2) pitch(m) = pitch/1000 = " & pitchMm.ToString("0.###") & "/1000 = " & pitchM.ToString("0.#####") & " m")
        sbCalc.AppendLine("3) L(m) = " & Lmm.ToString("0.###") & "/1000 = " & Lm.ToString("0.#####") & " m")
        sbCalc.AppendLine("4) acc(m) = " & accMm.ToString("0.###") & "/1000 = " & accM.ToString("0.#####") & " m")
        sbCalc.AppendLine("")
        sbCalc.AppendLine("5) countOnConv = floor(L/pitch) = floor(" & Lm.ToString("0.#####") & "/" & pitchM.ToString("0.#####") & ") = " & countOnConv.ToString())
        sbCalc.AppendLine("6) countAcc    = floor(acc/pitch) = floor(" & accM.ToString("0.#####") & "/" & pitchM.ToString("0.#####") & ") = " & countAcc.ToString())
        sbCalc.AppendLine("7) totalCount  = " & totalCount.ToString())
        sbCalc.AppendLine("")
        sbCalc.AppendLine("8) mProd = weight(g)/1000 = " & weightG.ToString("0.###") & "/1000 = " & mProd.ToString("0.#####") & " kg")
        sbCalc.AppendLine("9) mProductsTotal = totalCount*mProd = " & totalCount.ToString() & "*" & mProd.ToString("0.#####") & " = " & mProductsTotal.ToString("0.#####") & " kg")
        If accuracyOn Then
            sbCalc.AppendLine("10) conveyor mass = (kg/m)*length = " & massPerM.ToString("0.###") & "*" & effLenM.ToString("0.#####") & " = " & mConveyorMoving.ToString("0.#####") & " kg")
        Else
            sbCalc.AppendLine("10) conveyor mass assumed 0 (Accuracy Mode OFF)")
        End If
        sbCalc.AppendLine("11) mTotalMoving = " & mMovingTotal.ToString("0.#####") & " kg")
        sbCalc.AppendLine("")
        sbCalc.AppendLine("12) μ = μbase*envFactor = " & muBase.ToString("0.###") & "*" & envFactor.ToString("0.###") & " = " & mu.ToString("0.#####"))
        sbCalc.AppendLine("13) θ(rad) = θ(deg)*π/180 = " & thetaDeg.ToString("0.###") & "*π/180 = " & theta.ToString("0.#####"))
        sbCalc.AppendLine("")
        sbCalc.AppendLine("14) F_fric = μ*m*g*cosθ = " & mu.ToString("0.#####") & "*" & mMovingTotal.ToString("0.#####") & "*9.81*cos(" & theta.ToString("0.#####") & ") = " & Ffric.ToString("0.#####") & " N")
        sbCalc.AppendLine("15) F_incline = m*g*sinθ = " & mMovingTotal.ToString("0.#####") & "*9.81*sin(" & theta.ToString("0.#####") & ") = " & Fincline.ToString("0.#####") & " N")
        sbCalc.AppendLine("16) F_run = (F_fric+F_incline)*loadVar = (" & Ffric.ToString("0.#####") & "+" & Fincline.ToString("0.#####") & ")*" & loadVar.ToString("0.###") & " = " & Frun.ToString("0.#####") & " N")
        sbCalc.AppendLine("")
        sbCalc.AppendLine("17) speed(m/min) = PPM*pitch(m) = " & ppm.ToString("0.###") & "*" & pitchM.ToString("0.#####") & " = " & speedMMin.ToString("0.#####") & " m/min")
        sbCalc.AppendLine("18) speed(m/s) = speed/60 = " & speedMMin.ToString("0.#####") & "/60 = " & speedMS.ToString("0.#####") & " m/s")
        sbCalc.AppendLine("")
        sbCalc.AppendLine("19) radius = (driveDia/1000)/2 = (" & driveDiaMm.ToString("0.###") & "/1000)/2 = " & radius.ToString("0.#####") & " m")
        sbCalc.AppendLine("20) T_run = F_run*radius = " & Frun.ToString("0.#####") & "*" & radius.ToString("0.#####") & " = " & torqueRun.ToString("0.#####") & " Nm")
        sbCalc.AppendLine("21) T_design = T_run*SF = " & torqueRun.ToString("0.#####") & "*" & sf.ToString("0.###") & " = " & torqueDesign.ToString("0.#####") & " Nm")
        sbCalc.AppendLine("22) T_peak_start = T_design*StartFactor = " & torqueDesign.ToString("0.#####") & "*" & startFactor.ToString("0.###") & " = " & torquePeakStartFactor.ToString("0.#####") & " Nm")
        If accuracyOn AndAlso accelExtraTorque > 0 Then
            sbCalc.AppendLine("23) T_acc ≈ (m*a)*radius*SF (linear accel) = " & accelExtraTorque.ToString("0.#####") & " Nm")
        End If
        If inertiaOn AndAlso torqueInertia > 0 Then
            sbCalc.AppendLine("24) T_inertia = J*alpha*SF = " & torqueInertia.ToString("0.#####") & " Nm")
        End If
        sbCalc.AppendLine("25) T_peak_suggested = " & torquePeakSuggested.ToString("0.#####") & " Nm")
        sbCalc.AppendLine("")
        sbCalc.AppendLine("26) P_run = (F_run*v)/eff = (" & Frun.ToString("0.#####") & "*" & speedMS.ToString("0.#####") & ")/" & eff.ToString("0.###") & " = " & powerRunW.ToString("0.#####") & " W")
        sbCalc.AppendLine("27) P_design = P_run*SF = " & powerRunW.ToString("0.#####") & "*" & sf.ToString("0.###") & " = " & powerDesignW.ToString("0.#####") & " W")
        sbCalc.AppendLine("")
        sbCalc.AppendLine("28) driveRPM = v*60/(π*D) = " & driveRPM.ToString("0.#####") & " rpm")
        If gearRatio > 0 Then
            sbCalc.AppendLine("29) gearRatio ≈ motorRPM/driveRPM = " & motorRPM.ToString("0.###") & "/" & driveRPM.ToString("0.#####") & " = " & gearRatio.ToString("0.#####") & " : 1")
        End If
        sbCalc.AppendLine("============================================================")
        txtCalc.Text = sbCalc.ToString()

        lblStatus.Text = "Status: Calculation successful"
        lblStatus.ForeColor = themeOk
    End Sub

    ' ============================================================
    ' WARNINGS / ERRORS (bold red for errors)
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
                ev.Graphics.DrawString("Conveyor Calculation Report",
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

        ' CRITICAL: keep whatever theme MDAT applied (light or dark)
        SyncThemeFromCurrentUI()

        ClearOutputs()

        ClearTB(txtConvLen)
        ClearTB(txtProdLen)
        ClearTB(txtGap)
        ClearTB(txtWeight)
        ClearTB(txtMu)
        ClearTB(txtPPM)
        ClearTB(txtAccLen)
        ClearTB(txtIncline)
        ClearTB(txtDuty)
        ClearTB(txtLoadVar)

        ClearTB(txtRollerDia)
        ClearTB(txtEff)
        ClearTB(txtSF)
        ClearTB(txtStartFactor)
        ClearTB(txtMotorRPM)

        ClearTB(txtConvMassPerM)
        ClearTB(txtAccelTime)
        ClearTB(txtStartsPerHr)

        ClearTB(txtDensity)
        ClearTB(txtShaftDia)
        ClearTB(txtShaftLen)
        ClearTB(txtSprocketInertia)
        ClearTB(txtExtraInertia)

        SetComboSafe(cmbType, 0)
        SetComboSafe(cmbBed, 0)
        SetComboSafe(cmbEnv, 0)
        SetComboSafe(cmbMode, 0)
        SetComboSafe(cmbShaftMat, 0)

        SetCheckSafe(chkDesign, True)
        SetCheckSafe(chkAccuracy, False)
        SetCheckSafe(chkInertia, False)

        ToggleDesignFactors(Nothing, EventArgs.Empty)
        ToggleAccuracy(Nothing, EventArgs.Empty)
        ToggleInertia(Nothing, EventArgs.Empty)

        ConveyorTypeChanged(Nothing, EventArgs.Empty)
        ShaftMaterialChanged(Nothing, EventArgs.Empty)

        ' IMPORTANT: Do NOT call ApplyThemeToThisForm() here.
        ' That is what caused the form to repaint dark in your light MDAT theme.

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
    ' HELPERS (UI)
    ' ============================================================
    Private Sub AddSection(t As String, y As Integer)
        Dim lbl As New Label()
        lbl.Text = t
        lbl.ForeColor = themeAccent
        lbl.Location = New Point(10, y)
        lbl.AutoSize = True
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

    Private Sub AddOutputLabel(t As String, y As Integer)
        Dim lbl As New Label()
        lbl.Text = t
        lbl.ForeColor = themeAccent
        lbl.Location = New Point(0, y)
        lbl.AutoSize = True
        pnlResults.Controls.Add(lbl)
    End Sub

    Private Function AddOutputText(t As String, y As Integer, h As Integer) As TextBox
        AddOutputLabel(t, y)

        Dim tb As New TextBox()
        tb.Location = New Point(0, y + 20)
        tb.Size = New Size(760, h)
        tb.Multiline = True
        tb.ReadOnly = True
        tb.ScrollBars = ScrollBars.Vertical
        pnlResults.Controls.Add(tb)

        Return tb
    End Function

    ' ============================================================
    ' TOOLTIP SYSTEM (no auto tooltip on open)
    ' ============================================================
    Private Sub RegisterTip(c As Control, text As String)
        If c Is Nothing Then Exit Sub
        If tipMap Is Nothing Then Exit Sub

        Dim wrapped As String = WrapTip(text, 52)

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
            tip.Show(tipMap(c), c, x, 0, 7000)
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
