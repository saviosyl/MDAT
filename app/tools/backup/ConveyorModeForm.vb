 Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports System.Drawing.Printing

Public Class ConveyorModeForm
    Inherits Form

    '================ THEME =================
    Private ReadOnly BG As Color = Color.FromArgb(18, 22, 30)
    Private ReadOnly PANEL As Color = Color.FromArgb(28, 34, 44)
    Private ReadOnly ACCENT As Color = Color.FromArgb(0, 170, 210)
    Private ReadOnly WARN As Color = Color.Orange
    Private ReadOnly ERR As Color = Color.Red
    Private ReadOnly TXT As Color = Color.White

    '================ UI =================
    Private chkAdvanced As CheckBox
    Private pnlAdvanced As Panel
    Private txtResults As TextBox
    Private txtCalc As TextBox
    Private txtWarnings As TextBox

    ' Inputs
    Private txtPL, txtGap, txtWeight, txtMu, txtPPM, txtAccLen, txtIncline As TextBox
    Private txtRoller, txtRPM, txtEff, txtSF As TextBox
    Private cmbType As ComboBox

    '====================================================
    Public Sub New()
        Me.Text = "Conveyor Calculator"
        Me.Size = New Size(1050, 720)
        Me.BackColor = BG
        Me.StartPosition = FormStartPosition.CenterParent

        BuildUI()
    End Sub

    '====================================================
    Private Sub BuildUI()

        Dim y As Integer = 20

        AddLabel("Product Length (mm):", 20, y) : txtPL = AddBox(200, y) : y += 36
        AddLabel("Safety Gap (mm):", 20, y) : txtGap = AddBox(200, y) : y += 36
        AddLabel("Product Weight (g):", 20, y) : txtWeight = AddBox(200, y) : y += 36

        AddLabel("Conveyor Type:", 20, y)
        cmbType = New ComboBox()
        cmbType.Items.AddRange(New String() {"Belt", "Roller", "Modular Chain"})
        cmbType.SelectedIndex = 0
        cmbType.Location = New Point(200, y)
        Me.Controls.Add(cmbType)
        y += 36

        AddLabel("Friction Coefficient (μ):", 20, y) : txtMu = AddBox(200, y) : y += 36
        AddLabel("Target Products / Min:", 20, y) : txtPPM = AddBox(200, y) : y += 36
        AddLabel("Accumulated Length (mm):", 20, y) : txtAccLen = AddBox(200, y) : y += 36
        AddLabel("Incline Angle (deg):", 20, y) : txtIncline = AddBox(200, y) : y += 36

        chkAdvanced = New CheckBox()
        chkAdvanced.Text = "Enable Advanced Calculations"
        chkAdvanced.ForeColor = TXT
        chkAdvanced.Location = New Point(20, y)
        AddHandler chkAdvanced.CheckedChanged, AddressOf ToggleAdvanced
        Me.Controls.Add(chkAdvanced)
        y += 30

        pnlAdvanced = New Panel()
        pnlAdvanced.Location = New Point(20, y)
        pnlAdvanced.Size = New Size(360, 160)
        pnlAdvanced.Enabled = False
        pnlAdvanced.Visible = False
        Me.Controls.Add(pnlAdvanced)

        Dim ay As Integer = 0
        AddLabel("Roller Diameter (mm):", 0, ay, pnlAdvanced) : txtRoller = AddBox(180, ay, pnlAdvanced) : ay += 32
        AddLabel("Motor RPM:", 0, ay, pnlAdvanced) : txtRPM = AddBox(180, ay, pnlAdvanced) : ay += 32
        AddLabel("Drive Efficiency (%):", 0, ay, pnlAdvanced) : txtEff = AddBox(180, ay, pnlAdvanced) : ay += 32
        AddLabel("Safety Factor:", 0, ay, pnlAdvanced) : txtSF = AddBox(180, ay, pnlAdvanced)

        Dim btnCalc As New Button()
        btnCalc.Text = "Calculate"
        btnCalc.Location = New Point(20, 560)
        AddHandler btnCalc.Click, AddressOf Calculate
        Me.Controls.Add(btnCalc)

        txtResults = CreateOutputBox(420, 20, "FINAL RESULTS")
        txtCalc = CreateOutputBox(420, 260, "CALCULATION METHOD")
        txtWarnings = CreateOutputBox(420, 520, "WARNINGS")
        txtWarnings.ForeColor = WARN
    End Sub

    '====================================================
    Private Sub ToggleAdvanced(sender As Object, e As EventArgs)
        pnlAdvanced.Visible = chkAdvanced.Checked
        pnlAdvanced.Enabled = chkAdvanced.Checked
    End Sub

    '====================================================
 '====================================================
Private Sub Calculate(sender As Object, e As EventArgs)

    txtResults.Clear()
    txtCalc.Clear()
    txtWarnings.Clear()

    Dim PL As Double = Val(txtPL.Text)
    Dim GAP As Double = Val(txtGap.Text)
    Dim Wg As Double = Val(txtWeight.Text) / 1000.0
    Dim mu As Double = Val(txtMu.Text)
    Dim ppm As Double = Val(txtPPM.Text)
    Dim accLen As Double = Val(txtAccLen.Text) / 1000.0
    Dim incline As Double = Val(txtIncline.Text) * Math.PI / 180.0

    If PL <= 0 Or ppm <= 0 Then
        txtWarnings.ForeColor = ERR
        txtWarnings.Text = "ERROR: Invalid input values."
        Exit Sub
    End If

    Dim pitch As Double = (PL + GAP) / 1000.0
    Dim speed_m_min As Double = ppm * pitch
    Dim speed_m_s As Double = speed_m_min / 60.0

    Dim accCount As Integer = 0
    If accLen > 0 Then accCount = CInt(Math.Floor(accLen / pitch))

    Dim mass As Double = accCount * Wg
    Dim forceF As Double = mu * mass * 9.81
    Dim forceI As Double = mass * 9.81 * Math.Sin(incline)
    Dim totalF As Double = forceF + forceI

    ' ================= FINAL RESULTS =================
    txtResults.AppendText("Pitch              : " & pitch.ToString("0.000") & " m" & vbCrLf)
    txtResults.AppendText("Speed              : " & speed_m_min.ToString("0.00") & " m/min" & vbCrLf)
    txtResults.AppendText("Accumulation       : " & accCount & " products" & vbCrLf)
    txtResults.AppendText("Total Force        : " & totalF.ToString("0.0") & " N" & vbCrLf)

    ' ================= CALCULATION METHOD =================
    Dim calcText As String = ""

    calcText &= "MASS CALCULATION:" & vbCrLf
    calcText &= "m = accumulation × product_mass" & vbCrLf
    calcText &= "m = " & accCount & " × " & Wg.ToString("0.000") & " = " & mass.ToString("0.00") & " kg" & vbCrLf & vbCrLf

    calcText &= "FRICTION FORCE:" & vbCrLf
    calcText &= "F_friction = μ × m × g" & vbCrLf
    calcText &= "F_friction = " & mu.ToString("0.00") & " × " & mass.ToString("0.00") & " × 9.81" & vbCrLf
    calcText &= "F_friction = " & forceF.ToString("0.0") & " N" & vbCrLf & vbCrLf

    calcText &= "INCLINE FORCE:" & vbCrLf
    calcText &= "F_incline = m × g × sin(θ)" & vbCrLf
    calcText &= "F_incline = " & mass.ToString("0.00") & " × 9.81 × sin(" & (incline * 180 / Math.PI).ToString("0.0") & "°)" & vbCrLf
    calcText &= "F_incline = " & forceI.ToString("0.0") & " N" & vbCrLf & vbCrLf

    calcText &= "TOTAL DRIVE FORCE:" & vbCrLf
    calcText &= "F_total = F_friction + F_incline" & vbCrLf
    calcText &= "F_total = " & totalF.ToString("0.0") & " N" & vbCrLf & vbCrLf

    txtCalc.AppendText(calcText)

    ' ================= ADVANCED =================
    If chkAdvanced.Checked Then
        Dim D As Double = Val(txtRoller.Text) / 1000.0
        Dim rpm As Double = Val(txtRPM.Text)
        Dim eff As Double = Val(txtEff.Text) / 100.0
        Dim sf As Double = Val(txtSF.Text)

        Dim torque As Double = totalF * (D / 2.0)
        Dim mechW As Double = totalF * speed_m_s
        Dim motorW As Double = (mechW / eff) * sf
        Dim motorkW As Double = motorW / 1000.0

        Dim rollerRPM As Double = (speed_m_s / (Math.PI * D)) * 60.0
        Dim gearbox As Double = rpm / rollerRPM

        txtResults.AppendText("Torque             : " & torque.ToString("0.00") & " Nm" & vbCrLf)
        txtResults.AppendText("Motor Power        : " & motorkW.ToString("0.000") & " kW" & vbCrLf)
        txtResults.AppendText("Gearbox Ratio      : " & gearbox.ToString("0.0") & ":1" & vbCrLf)
    End If

End Sub


    '====================================================
    Private Function AddBox(x As Integer, y As Integer, Optional parent As Control = Nothing) As TextBox
        Dim t As New TextBox()
        t.Location = New Point(x, y)
        t.Width = 120
        If parent Is Nothing Then Me.Controls.Add(t) Else parent.Controls.Add(t)
        Return t
    End Function

    Private Sub AddLabel(txt As String, x As Integer, y As Integer, Optional parent As Control = Nothing)
        Dim l As New Label()
        l.Text = txt
        l.ForeColor = Color.FromName(TXT)
        l.Location = New Point(x, y + 4)
        l.AutoSize = True
        If parent Is Nothing Then Me.Controls.Add(l) Else parent.Controls.Add(l)
    End Sub

    Private Function CreateOutputBox(x As Integer, y As Integer, title As String) As TextBox
        Dim t As New TextBox()
        t.Location = New Point(x, y)
        t.Size = New Size(580, 200)
        t.Multiline = True
        t.ScrollBars = ScrollBars.Vertical
        t.BackColor = PANEL
        t.ForeColor = TXT
        t.ReadOnly = True
        Me.Controls.Add(t)
        t.AppendText(title & vbCrLf & New String("-"c, 30) & vbCrLf)
        Return t
    End Function

End Class
