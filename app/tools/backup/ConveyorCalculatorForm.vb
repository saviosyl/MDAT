Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Text
Imports System.Drawing.Printing
Imports System.IO

Public Class ConveyorCalculatorForm
    Inherits Form

    ' ================= THEME =================
    Private ReadOnly BG As Color = Color.FromArgb(18, 22, 30)
    Private ReadOnly PANEL As Color = Color.FromArgb(28, 34, 44)
    Private ReadOnly PURPLE As Color = Color.FromArgb(150, 90, 190)
    Private ReadOnly GREEN As Color = Color.LightGreen
    Private ReadOnly YELLOW As Color = Color.Gold
    Private ReadOnly TXT As Color = Color.White
    Private ReadOnly WARN As Color = Color.Gold
    Private ReadOnly ERR As Color = Color.OrangeRed
    Private ReadOnly OK As Color = Color.LightGreen

    ' ================= PANELS =================
    Private pnlInput As Panel
    Private pnlResults As Panel

    ' ================= OUTPUT =================
    Private txtFinal As TextBox
    Private txtCalc As TextBox
    Private txtWarn As TextBox
    Private lblStatus As Label

    ' ================= INPUTS =================
    Private txtConvLen, txtProdLen, txtGap, txtWeight As TextBox
    Private txtMu, txtPPM, txtAccLen, txtIncline As TextBox
    Private cmbType As ComboBox
    Private cmbBed As ComboBox

    ' Operating
    Private cmbEnv As ComboBox
    Private txtDuty As TextBox
    Private cmbMode As ComboBox
    Private txtLoadVar As TextBox

    ' Advanced
    Private chkAdv As CheckBox
    Private txtRoller, txtEff, txtSF As TextBox

    ' Reliability
    Private chkRel As CheckBox
    Private txtReliability As TextBox
    Private txtLifetime As TextBox

    ' Buttons
    Private btnCalc As Button
    Private btnPdf As Button
    Private btnReset As Button

    ' ================= INIT =================
    Public Sub New()
        Me.Text = "Conveyor Calculator"
        Me.Size = New Size(1180, 720)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = BG
        BuildUI()
    End Sub

    ' ================= BUILD UI =================
    Private Sub BuildUI()

        pnlInput = New Panel()
        pnlInput.Location = New Point(10, 10)
        pnlInput.Size = New Size(360, 660)
        pnlInput.BackColor = PANEL
        pnlInput.AutoScroll = True
        Me.Controls.Add(pnlInput)

        Dim y As Integer = 15

        AddSection("INPUT", y) : y += 26
        txtConvLen = AddInput("Conveyor Length (mm)", y) : y += 26
        txtProdLen = AddInput("Product Length (mm)", y) : y += 26
        txtGap = AddInput("Safety Gap (mm)", y) : y += 26
        txtWeight = AddInput("Product Weight (g)", y) : y += 26

        AddLabel("Conveyor Type", y)
        cmbType = New ComboBox()
        cmbType.Location = New Point(170, y)
        cmbType.Width = 160
        cmbType.Items.AddRange(New String() {"Belt", "Roller", "Modular Chain"})
        cmbType.SelectedIndex = 0
        AddHandler cmbType.SelectedIndexChanged, AddressOf ConveyorTypeChanged
        pnlInput.Controls.Add(cmbType)
        y += 26

        AddLabel("Belt Bed Material", y)
        cmbBed = New ComboBox()
        cmbBed.Location = New Point(170, y)
        cmbBed.Width = 160
        cmbBed.Items.AddRange(New String() {"Steel Bed", "UHMWPE", "Stainless Steel"})
        cmbBed.SelectedIndex = 0
        AddHandler cmbBed.SelectedIndexChanged, AddressOf BedMaterialChanged
        pnlInput.Controls.Add(cmbBed)
        y += 26

        txtMu = AddInput("Friction Coefficient (μ)", y) : y += 26
        txtPPM = AddInput("Target Products / Min", y) : y += 26
        txtAccLen = AddInput("Accumulated Length (mm)", y) : y += 26
        txtIncline = AddInput("Incline Angle (deg)", y) : y += 30

        AddSection("OPERATING CONDITIONS", y) : y += 26
        AddLabel("Environment", y)
        cmbEnv = New ComboBox()
        cmbEnv.Location = New Point(170, y)
        cmbEnv.Width = 160
        cmbEnv.Items.AddRange(New String() {"Clean", "Dusty", "Wet", "Washdown"})
        cmbEnv.SelectedIndex = 0
        pnlInput.Controls.Add(cmbEnv)
        y += 26

        txtDuty = AddInput("Duty Cycle (%)", y) : y += 26

        AddLabel("Operation Mode", y)
        cmbMode = New ComboBox()
        cmbMode.Location = New Point(170, y)
        cmbMode.Width = 160
        cmbMode.Items.AddRange(New String() {"Continuous", "Intermittent"})
        cmbMode.SelectedIndex = 0
        pnlInput.Controls.Add(cmbMode)
        y += 26

        txtLoadVar = AddInput("Load Variability Factor", y) : y += 30

        AddSection("ADVANCED", y) : y += 26
        chkAdv = AddCheck("Enable Advanced Factors", y) : y += 24
        txtRoller = AddInput("Roller Diameter (mm)", y) : y += 24
        txtEff = AddInput("Efficiency (%)", y) : y += 24
        txtSF = AddInput("Safety Factor", y) : y += 30

        AddSection("RELIABILITY & LIFETIME", y) : y += 26
        chkRel = AddCheck("Enable Reliability", y) : y += 24
        txtReliability = AddInput("Reliability Factor", y) : y += 24
        txtLifetime = AddInput("Expected Lifetime (hrs)", y) : y += 36

        btnCalc = New Button()
        btnCalc.Text = "Calculate"
        btnCalc.BackColor = PURPLE
        btnCalc.ForeColor = Color.White
        btnCalc.Location = New Point(20, y)
        btnCalc.Size = New Size(90, 32)
        AddHandler btnCalc.Click, AddressOf Calculate
        pnlInput.Controls.Add(btnCalc)

        btnPdf = New Button()
        btnPdf.Text = "Export PDF"
        btnPdf.BackColor = GREEN
        btnPdf.Location = New Point(130, y)
        btnPdf.Size = New Size(90, 32)
        AddHandler btnPdf.Click, AddressOf ExportPdf
        pnlInput.Controls.Add(btnPdf)

        btnReset = New Button()
        btnReset.Text = "Reset"
        btnReset.BackColor = YELLOW
        btnReset.Location = New Point(240, y)
        btnReset.Size = New Size(90, 32)
        AddHandler btnReset.Click, AddressOf ResetAll
        pnlInput.Controls.Add(btnReset)

        pnlResults = New Panel()
        pnlResults.Location = New Point(390, 10)
        pnlResults.Size = New Size(770, 660)
        pnlResults.BackColor = BG
        Me.Controls.Add(pnlResults)

        txtFinal = AddOutput("FINAL RESULTS", 10)
        txtCalc = AddOutput("CALCULATION METHOD (STEP-BY-STEP)", 240)
        txtWarn = AddOutput("WARNINGS / ERRORS", 470)
        txtWarn.ForeColor = WARN

        lblStatus = New Label()
        lblStatus.Text = "Status: Ready"
        lblStatus.ForeColor = OK
        lblStatus.Location = New Point(10, 685)
        Me.Controls.Add(lblStatus)

        ConveyorTypeChanged(Nothing, EventArgs.Empty)
    End Sub

    ' ================= AUTO MU =================
    Private Sub ConveyorTypeChanged(sender As Object, e As EventArgs)
        cmbBed.Enabled = (cmbType.Text = "Belt")
        txtMu.ReadOnly = (cmbType.Text = "Belt")
        BedMaterialChanged(Nothing, EventArgs.Empty)
    End Sub

    Private Sub BedMaterialChanged(sender As Object, e As EventArgs)
        If cmbType.Text <> "Belt" Then Exit Sub
        If cmbBed.Text = "Steel Bed" Then txtMu.Text = "0.30"
        If cmbBed.Text = "UHMWPE" Then txtMu.Text = "0.20"
        If cmbBed.Text = "Stainless Steel" Then txtMu.Text = "0.25"
    End Sub

    ' ================= CALCULATION =================
' ================= CALCULATION =================
Private Sub Calculate(sender As Object, e As EventArgs)

    txtFinal.Clear()
    txtCalc.Clear()
    txtWarn.Clear()

    Dim Lmm As Double = Val(txtConvLen.Text)
    Dim Lm As Double = Lmm / 1000.0
    Dim pitchMm As Double = Val(txtProdLen.Text) + Val(txtGap.Text)
    Dim pitchM As Double = pitchMm / 1000.0
    Dim ppm As Double = Val(txtPPM.Text)
    Dim mProd As Double = Val(txtWeight.Text) / 1000.0
    Dim accM As Double = Val(txtAccLen.Text) / 1000.0
    Dim muBase As Double = Val(txtMu.Text)
    Dim thetaDeg As Double = Val(txtIncline.Text)
    Dim theta As Double = thetaDeg * Math.PI / 180.0

    If Lm <= 0 Or pitchM <= 0 Or ppm <= 0 Then
        txtWarn.Text = "Invalid inputs detected."
        lblStatus.Text = "Status: Error"
        Exit Sub
    End If

    Dim count As Integer = CInt(Math.Floor(Lm / pitchM))
    Dim accCount As Integer = CInt(Math.Floor(accM / pitchM))
    Dim totalCount As Integer = count + accCount
    Dim mTotal As Double = totalCount * mProd

    Dim envFactor As Double = 1.0
    If cmbEnv.Text = "Dusty" Then envFactor = 1.15
    If cmbEnv.Text = "Wet" Then envFactor = 1.30
    If cmbEnv.Text = "Washdown" Then envFactor = 1.45

    Dim loadVar As Double = Val(txtLoadVar.Text)
    If loadVar <= 0 Then loadVar = 1.0

    Dim mu As Double = muBase * envFactor

    Dim Ffric As Double = mu * mTotal * 9.81 * Math.Cos(theta)
    Dim Fincline As Double = mTotal * 9.81 * Math.Sin(theta)
    Dim Ftotal As Double = (Ffric + Fincline) * loadVar

    Dim speedMMin As Double = ppm * pitchM
    Dim speedMS As Double = speedMMin / 60.0

    If speedMMin > 65 Then
        txtWarn.AppendText("WARNING: Conveyor speed exceeds 65 m/min." & vbCrLf)
    End If

    Dim radius As Double = 0.05
    If chkAdv.Checked And Val(txtRoller.Text) > 0 Then
        radius = Val(txtRoller.Text) / 2000.0
    End If

    Dim torque As Double = Ftotal * radius

    Dim eff As Double = 1.0
    If chkAdv.Checked And Val(txtEff.Text) > 0 Then eff = Val(txtEff.Text) / 100.0

    Dim sf As Double = 1.0
    If chkAdv.Checked And Val(txtSF.Text) > 0 Then sf = Val(txtSF.Text)

    Dim rel As Double = 1.0
    If chkRel.Checked And Val(txtReliability.Text) > 0 Then rel = Val(txtReliability.Text)

    Dim adjTorque As Double = torque * sf * rel
    Dim power As Double = (Ftotal * speedMS) / eff

    ' -------- FINAL RESULTS --------
    txtFinal.Text =
        "Total Products        : " & totalCount & vbCrLf &
        "Total Mass            : " & mTotal.ToString("0.00") & " kg" & vbCrLf &
        "Belt Speed            : " & speedMMin.ToString("0.00") & " m/min (" & speedMS.ToString("0.000") & " m/s)" & vbCrLf &
        "Incline Angle         : " & thetaDeg & " deg" & vbCrLf &
        "Drive Force           : " & Ftotal.ToString("0.0") & " N" & vbCrLf &
        "Torque Required       : " & adjTorque.ToString("0.00") & " Nm" & vbCrLf &
        "Motor Power           : " & power.ToString("0.0") & " W (" & (power / 1000.0).ToString("0.000") & " kW)"

    ' -------- CALCULATION METHOD (BUILD SAFE) --------
    Dim sb As New StringBuilder()

    sb.AppendLine("1) Total Mass")
    sb.AppendLine("m = product_mass * total_products")
    sb.AppendLine("m = " & mProd.ToString("0.000") & " * " & totalCount & " = " & mTotal.ToString("0.00") & " kg")
    sb.AppendLine()

    sb.AppendLine("2) Friction Force")
    sb.AppendLine("F_fric = mu * m * g * cos(theta)")
    sb.AppendLine("mu = " & mu.ToString("0.00") & ", theta = " & thetaDeg & " deg")
    sb.AppendLine("F_fric = " & Ffric.ToString("0.0") & " N")
    sb.AppendLine()

    sb.AppendLine("3) Incline Force")
    sb.AppendLine("F_incline = m * g * sin(theta)")
    sb.AppendLine("F_incline = " & Fincline.ToString("0.0") & " N")
    sb.AppendLine()

    sb.AppendLine("4) Total Drive Force")
    sb.AppendLine("F_total = (F_fric + F_incline) * load_variability")
    sb.AppendLine("F_total = " & Ftotal.ToString("0.0") & " N")
    sb.AppendLine()

    sb.AppendLine("5) Torque")
    sb.AppendLine("T = F * r")
    sb.AppendLine("T = " & Ftotal.ToString("0.0") & " * " & radius.ToString("0.000"))
    sb.AppendLine("T = " & adjTorque.ToString("0.00") & " Nm")
    sb.AppendLine()

    sb.AppendLine("6) Power")
    sb.AppendLine("P = F * v / efficiency")
    sb.AppendLine("P = " & power.ToString("0.0") & " W (" & (power / 1000.0).ToString("0.000") & " kW)")
    sb.AppendLine()
    sb.AppendLine("--------------------------------")
    sb.AppendLine("Calculations are indicative only.")
    sb.AppendLine("Verify against ISO, CEMA, and manufacturer data.")
    sb.AppendLine("XXXXX accepts no liability.")

    txtCalc.Text = sb.ToString()

    lblStatus.Text = "Status: Calculation successful"

End Sub

' ================= PDF EXPORT =================
Private Sub ExportPdf(sender As Object, e As EventArgs)
    Dim sfd As New SaveFileDialog()
    sfd.Filter = "PDF Files|*.pdf"
    If sfd.ShowDialog() <> DialogResult.OK Then Exit Sub

    Dim doc As New PrintDocument()
    AddHandler doc.PrintPage, Sub(s, ev)
        ev.Graphics.DrawString("Conveyor Calculation Report",
                               New Font("Segoe UI", 14, FontStyle.Bold),
                               Brushes.Black, 50, 50)
        ev.Graphics.DrawString(txtFinal.Text & vbCrLf & vbCrLf & txtCalc.Text,
                               New Font("Consolas", 9),
                               Brushes.Black, 50, 100)
    End Sub

    doc.PrinterSettings.PrintToFile = True
    doc.PrinterSettings.PrintFileName = sfd.FileName
    doc.Print()
End Sub

Private Sub ResetAll(sender As Object, e As EventArgs)
    For Each c As Control In pnlInput.Controls
        If TypeOf c Is TextBox Then DirectCast(c, TextBox).Clear()
    Next
    txtFinal.Clear()
    txtCalc.Clear()
    txtWarn.Clear()
    lblStatus.Text = "Status: Reset"
End Sub


    ' ================= HELPERS =================
    Private Sub AddSection(t As String, y As Integer)
        pnlInput.Controls.Add(New Label() With {.Text = t, .ForeColor = PURPLE, .Location = New Point(10, y), .AutoSize = True})
    End Sub

    Private Function AddInput(t As String, y As Integer) As TextBox
        AddLabel(t, y)
        Dim tb As New TextBox()
        tb.Location = New Point(170, y - 2)
        tb.Width = 160
        pnlInput.Controls.Add(tb)
        Return tb
    End Function

    Private Sub AddLabel(t As String, y As Integer)
        pnlInput.Controls.Add(New Label() With {.Text = t, .ForeColor = TXT, .Location = New Point(14, y), .AutoSize = True})
    End Sub

    Private Function AddCheck(t As String, y As Integer) As CheckBox
        Dim cb As New CheckBox()
        cb.Text = t
        cb.ForeColor = TXT
        cb.Location = New Point(14, y)
        pnlInput.Controls.Add(cb)
        Return cb
    End Function

    Private Function AddOutput(t As String, y As Integer) As TextBox
        pnlResults.Controls.Add(New Label() With {.Text = t, .ForeColor = PURPLE, .Location = New Point(0, y), .AutoSize = True})
        Dim tb As New TextBox()
        tb.Location = New Point(0, y + 20)
        tb.Size = New Size(760, 200)
        tb.Multiline = True
        tb.ReadOnly = True
        tb.BackColor = PANEL
        tb.ForeColor = TXT
        tb.ScrollBars = ScrollBars.Vertical
        pnlResults.Controls.Add(tb)
        Return tb
    End Function

End Class
