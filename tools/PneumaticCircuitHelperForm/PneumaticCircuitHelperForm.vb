Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Imports System.Text
Imports System.IO
Imports System.Globalization
Imports System.Collections.Generic
Imports System.Drawing.Printing

Public Class PneumaticCircuitHelperForm
    Inherits Form

    ' ============================================================
    ' THEME (MDAT-compatible)
    ' ============================================================
    Private themeBG As Color = Color.FromArgb(18, 22, 30)
    Private themePanel As Color = Color.FromArgb(28, 34, 44)
    Private themeAccent As Color = Color.FromArgb(150, 90, 190)
    Private themeText As Color = Color.Gainsboro

    Private themeWarn As Color = Color.Gold
    Private themeErr As Color = Color.OrangeRed
    Private themeOk As Color = Color.LightGreen

    Private btnGreen As Color = Color.LightGreen
    Private btnYellow As Color = Color.Gold

    ' ============================================================
    ' ROOT LAYOUT (responsive)
    ' ============================================================
    Private splitMain As SplitContainer
    Private pnlInput As Panel
    Private tlRight As TableLayoutPanel

    ' input content host (manual vertical layout)
    Private inputContent As Panel
    Private btnBar As FlowLayoutPanel

    ' FIX: preferred splitter minimums
    Private _preferredPanel1Min As Integer = 360
    Private _preferredPanel2Min As Integer = 650

    ' ============================================================
    ' OUTPUT CONTROLS
    ' ============================================================
    Private txtFinal As TextBox
    Private txtCalc As TextBox
    Private txtBom As TextBox
    Private rtbWarn As RichTextBox
    Private lblStatus As Label

    ' preview
    Private pnlPreviewToolbar As Panel
    Private picPreview As PictureBox
    Private cmbPreviewMode As ComboBox

    ' ============================================================
    ' PHASE-2 PROJECT MODE CONTROLS
    ' ============================================================
    Private cmbProjectMode As ComboBox
    Private txtJobName As TextBox
    Private lstCylinderJobs As ListBox
    Private btnAddJob As Button
    Private btnUpdateJob As Button
    Private btnRemoveJob As Button
    Private btnLoadJob As Button
    Private btnNewFromCurrent As Button
    Private lblJobsInfo As Label

    Private jobs As List(Of CylinderJob) = New List(Of CylinderJob)()
    Private currentJobIndex As Integer = -1

    ' ============================================================
    ' INPUT CONTROLS
    ' ============================================================
    Private cmbCylinderType As ComboBox
    Private txtBore As TextBox
    Private txtStroke As TextBox
    Private txtQty As TextBox
    Private cmbOrientation As ComboBox
    Private cmbMotionType As ComboBox

    Private cmbOperationMode As ComboBox
    Private cmbFailSafe As ComboBox
    Private cmbSpeedPriority As ComboBox
    Private cmbCushioning As ComboBox

    Private cmbVendor As ComboBox
    Private cmbVoltage As ComboBox

    Private txtPressure As TextBox
    Private txtTubeLen As TextBox
    Private txtCyclesPerMin As TextBox
    Private cmbAirQuality As ComboBox
    Private cmbEnvironment As ComboBox

    Private chkFRL As CheckBox
    Private chkBranchReg As CheckBox
    Private chkQuickExhaust As CheckBox
    Private chkSensors As CheckBox
    Private chkManualOverride As CheckBox
    Private chkDumpValve As CheckBox

    ' Buttons
    Private btnCalc As Button
    Private btnPreview As Button
    Private btnExportDxfSmart As Button
    Private btnExportDxfStd As Button
    Private btnPdf As Button
    Private btnReset As Button

    ' Tooltips
    Private tip As ToolTip
    Private tipMap As Dictionary(Of Control, String)
    Private tipsArmed As Boolean = False

    ' ============================================================
    ' LAST CALC STATE (shared by preview / DXF)
    ' ============================================================
    Private lastCalcOk As Boolean = False
    Private lastIsDoubleActing As Boolean = True
    Private lastCylinderLabel As String = "CYL-01"
    Private lastValveLabel As String = "SV-01"
    Private lastPortSize As String = "G1/8"
    Private lastTubeOD As String = "6 mm"
    Private lastLineTags As List(Of String) = New List(Of String)()
    Private lastBOMLines As List(Of String) = New List(Of String)()
    Private lastSummaryTitle As String = "Pneumatic Circuit"
    Private lastDrawSensors As Boolean = False
    Private lastDrawQuickExhaust As Boolean = False
    Private lastDrawFRL As Boolean = False
    Private lastDrawBranchReg As Boolean = False
    Private lastPreviewMode As String = "Standard"

    ' ============================================================
    ' PRINT / PDF (Option B - Microsoft Print to PDF, multi-page)
    ' ============================================================
    Private _printLines As List(Of String) = New List(Of String)()
    Private _printLineIndex As Integer = 0
    Private _printHeaderTitle As String = "MetaMech Pneumatic Circuit Helper Report"
    Private _printHeaderGenerated As String = ""
    Private _printBodyFont As Font = Nothing
    Private _printHeaderFont As Font = Nothing
    Private _printSubFont As Font = Nothing

    ' ============================================================
    ' MODEL (Phase-2 per-cylinder jobs)
    ' ============================================================
    Private Class CylinderJob
        Public Property JobName As String
        Public Property CylinderType As String
        Public Property Bore As String
        Public Property Stroke As String
        Public Property Qty As String
        Public Property Orientation As String
        Public Property MotionType As String

        Public Property OperationMode As String
        Public Property FailSafe As String
        Public Property SpeedPriority As String
        Public Property Cushioning As String

        Public Property Vendor As String
        Public Property Voltage As String

        Public Property Pressure As String
        Public Property TubeLen As String
        Public Property CyclesPerMin As String
        Public Property AirQuality As String
        Public Property Environment As String

        Public Property AddFRL As Boolean
        Public Property AddBranchReg As Boolean
        Public Property AddQuickExhaust As Boolean
        Public Property AddSensors As Boolean
        Public Property AddManualOverride As Boolean
        Public Property AddDumpValve As Boolean

        Public Function DisplayText(ByVal index1 As Integer) As String
            Dim nm As String = If(String.IsNullOrWhiteSpace(JobName), "Cylinder " & index1.ToString(CultureInfo.InvariantCulture), JobName.Trim())
            Dim typ As String = If(String.IsNullOrWhiteSpace(CylinderType), "DA", CylinderType)
            Dim b As String = If(String.IsNullOrWhiteSpace(Bore), "?", Bore)
            Dim s As String = If(String.IsNullOrWhiteSpace(Stroke), "?", Stroke)
            Return index1.ToString(CultureInfo.InvariantCulture) & ". " & nm & "  |  " & typ & "  |  Ø" & b & " x " & s
        End Function

        Public Function ShallowClone() As CylinderJob
            Dim j As New CylinderJob()
            j.JobName = Me.JobName
            j.CylinderType = Me.CylinderType
            j.Bore = Me.Bore
            j.Stroke = Me.Stroke
            j.Qty = Me.Qty
            j.Orientation = Me.Orientation
            j.MotionType = Me.MotionType
            j.OperationMode = Me.OperationMode
            j.FailSafe = Me.FailSafe
            j.SpeedPriority = Me.SpeedPriority
            j.Cushioning = Me.Cushioning
            j.Vendor = Me.Vendor
            j.Voltage = Me.Voltage
            j.Pressure = Me.Pressure
            j.TubeLen = Me.TubeLen
            j.CyclesPerMin = Me.CyclesPerMin
            j.AirQuality = Me.AirQuality
            j.Environment = Me.Environment
            j.AddFRL = Me.AddFRL
            j.AddBranchReg = Me.AddBranchReg
            j.AddQuickExhaust = Me.AddQuickExhaust
            j.AddSensors = Me.AddSensors
            j.AddManualOverride = Me.AddManualOverride
            j.AddDumpValve = Me.AddDumpValve
            Return j
        End Function
    End Class

    ' ============================================================
    ' CTOR
    ' ============================================================
    Public Sub New()
        Me.Text = "Pneumatic Circuit Helper"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Size = New Size(1500, 900)
        Me.MinimumSize = New Size(1180, 760)
        Me.BackColor = themeBG

        BuildUI()

        AddHandler Me.MouseDown, AddressOf ArmTipsOnFirstClick
        AddHandler Me.Shown, AddressOf FormShown_AdjustLayout
        AddHandler Me.Resize, AddressOf FormResize_AdjustLayout

        SyncThemeFromCurrentUI()
        ApplyThemeToThisForm()
        RenderPreviewPlaceholder("Click Calculate, then Preview.")
        UpdateJobsPanelState()
    End Sub

    Public Sub ApplyMDATTheme(bg As Color, panel As Color, accent As Color, isDark As Boolean)
        themeBG = bg
        themePanel = panel
        themeAccent = accent
        If isDark Then
            themeText = Color.Gainsboro
        Else
            themeText = Color.FromArgb(35, 35, 35)
        End If
        ApplyThemeToThisForm()
    End Sub

    Private Sub ArmTipsOnFirstClick(sender As Object, e As MouseEventArgs)
        tipsArmed = True
    End Sub

    Private Sub FormShown_AdjustLayout(sender As Object, e As EventArgs)
        SafeSetSplitter(splitMain, 470)
    End Sub

    Private Sub FormResize_AdjustLayout(sender As Object, e As EventArgs)
        If splitMain IsNot Nothing Then
            SafeSetSplitter(splitMain, splitMain.SplitterDistance)
        End If
        If picPreview IsNot Nothing AndAlso picPreview.Width > 10 AndAlso picPreview.Height > 10 Then
            If lastCalcOk Then
                DrawPreview()
            End If
        End If
    End Sub

    Private Sub SafeSetSplitter(ByVal sc As SplitContainer, ByVal desired As Integer)
        If sc Is Nothing Then Exit Sub
        If Not sc.IsHandleCreated Then Exit Sub

        Dim totalWidth As Integer = sc.ClientSize.Width
        If totalWidth <= 0 Then Exit Sub

        Dim p1Min As Integer = Math.Max(0, _preferredPanel1Min)
        Dim p2Min As Integer = Math.Max(0, _preferredPanel2Min)

        Dim available As Integer = totalWidth - sc.SplitterWidth
        If available <= 0 Then Exit Sub

        If p1Min + p2Min > available Then
            Dim overflow As Integer = (p1Min + p2Min) - available
            Dim cut1 As Integer = overflow \ 2
            Dim cut2 As Integer = overflow - cut1

            p1Min = Math.Max(180, p1Min - cut1)
            p2Min = Math.Max(260, p2Min - cut2)

            If p1Min + p2Min > available Then
                p1Min = Math.Max(0, available \ 3)
                p2Min = Math.Max(0, available \ 3)
            End If
        End If

        Try
            sc.Panel1MinSize = p1Min
            sc.Panel2MinSize = p2Min
        Catch
            Exit Sub
        End Try

        Dim minVal As Integer = sc.Panel1MinSize
        Dim maxVal As Integer = totalWidth - sc.Panel2MinSize - sc.SplitterWidth

        If maxVal < minVal Then
            Try
                sc.Panel1MinSize = 0
                sc.Panel2MinSize = 0
                minVal = 0
                maxVal = totalWidth - sc.SplitterWidth
            Catch
                Exit Sub
            End Try
            If maxVal < minVal Then Exit Sub
        End If

        If desired < minVal Then desired = minVal
        If desired > maxVal Then desired = maxVal

        Try
            sc.SplitterDistance = desired
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

        splitMain = New SplitContainer()
        splitMain.Dock = DockStyle.Fill
        splitMain.Orientation = Orientation.Vertical
        splitMain.SplitterWidth = 6

        splitMain.Panel1MinSize = 0
        splitMain.Panel2MinSize = 0

        Try
            splitMain.SplitterDistance = Math.Max(100, Math.Min(470, Me.ClientSize.Width \ 3))
        Catch
        End Try

        splitMain.BackColor = themeBG
        Me.Controls.Add(splitMain)

        ' LEFT SIDE
        pnlInput = New Panel()
        pnlInput.Dock = DockStyle.Fill
        pnlInput.BackColor = themePanel
        splitMain.Panel1.Controls.Add(pnlInput)

        btnBar = New FlowLayoutPanel()
        btnBar.Dock = DockStyle.Bottom
        btnBar.Height = 84
        btnBar.WrapContents = True
        btnBar.FlowDirection = FlowDirection.LeftToRight
        btnBar.Padding = New Padding(8, 8, 8, 8)
        btnBar.BackColor = themePanel
        pnlInput.Controls.Add(btnBar)

        inputContent = New Panel()
        inputContent.Dock = DockStyle.Fill
        inputContent.AutoScroll = True
        inputContent.BackColor = themePanel
        pnlInput.Controls.Add(inputContent)

        BuildInputContent()

        ' RIGHT SIDE
        tlRight = New TableLayoutPanel()
        tlRight.Dock = DockStyle.Fill
        tlRight.BackColor = themeBG
        tlRight.ColumnCount = 1
        tlRight.RowCount = 6
        tlRight.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        tlRight.RowStyles.Add(New RowStyle(SizeType.Percent, 22.0F)) ' Final
        tlRight.RowStyles.Add(New RowStyle(SizeType.Percent, 20.0F)) ' Calc logic
        tlRight.RowStyles.Add(New RowStyle(SizeType.Percent, 17.0F)) ' BOM
        tlRight.RowStyles.Add(New RowStyle(SizeType.Percent, 16.0F)) ' Warnings
        tlRight.RowStyles.Add(New RowStyle(SizeType.Percent, 25.0F)) ' Preview
        tlRight.RowStyles.Add(New RowStyle(SizeType.Absolute, 28.0F)) ' Status
        splitMain.Panel2.Controls.Add(tlRight)

        txtFinal = CreateOutputTextBox()
        txtCalc = CreateOutputTextBox()
        txtBom = CreateOutputTextBox()

        rtbWarn = New RichTextBox()
        rtbWarn.Dock = DockStyle.Fill
        rtbWarn.ReadOnly = True
        rtbWarn.BorderStyle = BorderStyle.FixedSingle
        rtbWarn.Multiline = True
        rtbWarn.ScrollBars = RichTextBoxScrollBars.Vertical
        rtbWarn.BackColor = themePanel
        rtbWarn.ForeColor = themeText
        rtbWarn.Font = New Font("Consolas", 9.0F, FontStyle.Regular)

        tlRight.Controls.Add(CreateDockGroup("FINAL CIRCUIT / PROJECT RECOMMENDATION", txtFinal), 0, 0)
        tlRight.Controls.Add(CreateDockGroup("CIRCUIT LOGIC / WHY THIS SELECTION", txtCalc), 0, 1)
        tlRight.Controls.Add(CreateDockGroup("DRAFT BOM (COPY-READY)", txtBom), 0, 2)
        tlRight.Controls.Add(CreateDockGroup("WARNINGS / NOTES", rtbWarn), 0, 3)
        tlRight.Controls.Add(CreatePreviewGroup(), 0, 4)

        lblStatus = New Label()
        lblStatus.Dock = DockStyle.Fill
        lblStatus.AutoSize = False
        lblStatus.TextAlign = ContentAlignment.MiddleLeft
        lblStatus.Padding = New Padding(4, 0, 0, 0)
        lblStatus.Text = "Status: Ready"
        lblStatus.ForeColor = themeOk
        lblStatus.BackColor = themeBG
        tlRight.Controls.Add(lblStatus, 0, 5)
    End Sub

    Private Sub BuildInputContent()
        Dim y As Integer = 12

        ' -------------------- PHASE-2 PROJECT MODE --------------------
        AddSection("PROJECT MODE (PHASE-2)", y) : y += 26

        AddLabel("Mode", y)
        cmbProjectMode = AddCombo(y, New String() {"Single Circuit (Representative)", "Multi Cylinder - Independent"}, "Single = one circuit. Multi = each cylinder/job stored separately and merged in report/BOM.")
        y += 26

        txtJobName = AddInput("Job Name / Tag", y, "Name for current cylinder job (e.g., Clamp-01, Stopper-02, Lift-01).")
        y += 26

        lblJobsInfo = New Label()
        lblJobsInfo.Text = "Jobs: 0 (Multi mode)"
        lblJobsInfo.ForeColor = themeWarn
        lblJobsInfo.AutoSize = True
        lblJobsInfo.Location = New Point(14, y + 4)
        inputContent.Controls.Add(lblJobsInfo)

        btnAddJob = CreateMiniButton("Add", AddressOf AddJobFromCurrent)
        btnAddJob.Location = New Point(215, y)
        inputContent.Controls.Add(btnAddJob)

        btnUpdateJob = CreateMiniButton("Update", AddressOf UpdateSelectedJobFromCurrent)
        btnUpdateJob.Location = New Point(280, y)
        inputContent.Controls.Add(btnUpdateJob)
        y += 28

        btnLoadJob = CreateMiniButton("Load", AddressOf LoadSelectedJobToInputs)
        btnLoadJob.Location = New Point(215, y)
        inputContent.Controls.Add(btnLoadJob)

        btnRemoveJob = CreateMiniButton("Remove", AddressOf RemoveSelectedJob)
        btnRemoveJob.Location = New Point(280, y)
        inputContent.Controls.Add(btnRemoveJob)
        y += 28

        btnNewFromCurrent = CreateMiniButton("Duplicate->New", AddressOf DuplicateCurrentIntoNewJob)
        btnNewFromCurrent.Location = New Point(215, y)
        btnNewFromCurrent.Width = 145
        inputContent.Controls.Add(btnNewFromCurrent)
        y += 30

        lstCylinderJobs = New ListBox()
        lstCylinderJobs.Location = New Point(14, y)
        lstCylinderJobs.Size = New Size(346, 110)
        lstCylinderJobs.HorizontalScrollbar = True
        AddHandler lstCylinderJobs.SelectedIndexChanged, AddressOf JobsList_SelectedIndexChanged
        inputContent.Controls.Add(lstCylinderJobs)
        y += 120

        ' -------------------- CYLINDER / MOTION --------------------
        AddSection("CYLINDER / MOTION", y) : y += 26
        AddLabel("Cylinder Type", y)
        cmbCylinderType = AddCombo(y, New String() {"Double Acting", "Single Acting (Spring Return)"}, "Single-acting = air extend, spring return.")
        y += 26

        txtBore = AddInput("Bore (mm)", y, "Cylinder bore diameter in mm.") : y += 26
        txtStroke = AddInput("Stroke (mm)", y, "Cylinder stroke in mm.") : y += 26
        txtQty = AddInput("Cylinder Qty", y, "Number of identical cylinders for THIS job.") : y += 26

        AddLabel("Orientation", y)
        cmbOrientation = AddCombo(y, New String() {"Horizontal", "Vertical Up", "Vertical Down"}, "Cylinder orientation.")
        y += 26

        AddLabel("Motion Type", y)
        cmbMotionType = AddCombo(y, New String() {"General", "Clamp", "Push", "Stopper", "Slide", "Gate", "Lift"}, "Application hint.")
        y += 30

        ' -------------------- CONTROL / LOGIC --------------------
        AddSection("CONTROL / LOGIC", y) : y += 26

        AddLabel("Operation Mode", y)
        cmbOperationMode = AddCombo(y, New String() {"Manual", "Semi-Auto", "Full Auto (PLC)"}, "Control mode affects valve/coil notes.")
        y += 26

        AddLabel("Fail-safe State", y)
        cmbFailSafe = AddCombo(y, New String() {"Retracted on power loss", "Extended on power loss", "Hold last position (not guaranteed)"}, "Fail-safe preference.")
        y += 26

        AddLabel("Speed Priority", y)
        cmbSpeedPriority = AddCombo(y, New String() {"Smooth motion", "Fast cycle", "Precise stop"}, "Influences FCV / QEV recommendations.")
        y += 26

        AddLabel("Cushioning", y)
        cmbCushioning = AddCombo(y, New String() {"Not required", "Preferred", "Required"}, "Cylinder cushioning preference.")
        y += 30

        ' -------------------- VENDOR BOM --------------------
        AddSection("VENDOR BOM RECOMMENDATION", y) : y += 26

        AddLabel("Vendor", y)
        cmbVendor = AddCombo(y, New String() {"Generic", "Festo", "SMC"}, "Select vendor for recommended BOM series / sample part numbers.")
        y += 26

        AddLabel("Coil Voltage", y)
        cmbVoltage = AddCombo(y, New String() {"24VDC", "230VAC", "110VAC"}, "Valve coil voltage for BOM recommendation.")
        y += 30

        ' -------------------- AIR / ENVIRONMENT --------------------
        AddSection("AIR SUPPLY / ENVIRONMENT", y) : y += 26

        txtPressure = AddInput("Supply Pressure (bar)", y, "Typical machine air supply pressure.") : y += 26
        txtTubeLen = AddInput("Tube Length (m)", y, "Approx total tubing length for estimate.") : y += 26
        txtCyclesPerMin = AddInput("Cycles / min", y, "Machine cycle rate (approx).") : y += 26

        AddLabel("Air Quality", y)
        cmbAirQuality = AddCombo(y, New String() {"General dry air", "Clean dry air", "Oil-lubricated", "Unknown"}, "Air quality note.")
        y += 26

        AddLabel("Environment", y)
        cmbEnvironment = AddCombo(y, New String() {"Clean", "Dusty", "Wet", "Washdown"}, "Environmental severity.")
        y += 30

        ' -------------------- OPTIONS --------------------
        AddSection("OPTIONS / SAFETY", y) : y += 26

        chkFRL = AddCheck("Add FRL unit", y, "Main filter/regulator/lubricator (or filter-regulator) set.") : y += 24
        chkBranchReg = AddCheck("Add branch regulator + gauge", y, "Adds local pressure adjustment branch.") : y += 24
        chkQuickExhaust = AddCheck("Add quick exhaust valve", y, "Useful for fast extension/retraction in some cases.") : y += 24
        chkSensors = AddCheck("Add end sensors (S1/S2)", y, "Cylinder position sensors (if magnetic piston).") : y += 24
        chkManualOverride = AddCheck("Manual override on valve", y, "Prefer manual override on solenoid valve.") : y += 24
        chkDumpValve = AddCheck("Soft-start / dump valve", y, "Add dump/soft-start valve for machine safety/service.") : y += 24

        y += 10

        ' buttons
        btnCalc = CreateButton("Calculate", themeAccent, True, AddressOf CalculateCircuit)
        btnPreview = CreateButton("Preview", Color.FromArgb(90, 170, 255), False, AddressOf PreviewCircuit)
        btnExportDxfSmart = CreateButton("DXF Smart", btnGreen, False, AddressOf ExportDxfSmart)
        btnExportDxfStd = CreateButton("DXF Std", btnGreen, False, AddressOf ExportDxfStandard)
        btnPdf = CreateButton("Export PDF", btnGreen, False, AddressOf ExportPdf)
        btnReset = CreateButton("Reset", btnYellow, False, AddressOf ResetAll)

        btnBar.Controls.Add(btnCalc)
        btnBar.Controls.Add(btnPreview)
        btnBar.Controls.Add(btnExportDxfSmart)
        btnBar.Controls.Add(btnExportDxfStd)
        btnBar.Controls.Add(btnPdf)
        btnBar.Controls.Add(btnReset)

        ' defaults
        SetComboSafe(cmbProjectMode, 0)
        SetComboSafe(cmbCylinderType, 0)
        SetComboSafe(cmbOrientation, 0)
        SetComboSafe(cmbMotionType, 0)
        SetComboSafe(cmbOperationMode, 2)
        SetComboSafe(cmbFailSafe, 0)
        SetComboSafe(cmbSpeedPriority, 0)
        SetComboSafe(cmbCushioning, 1)
        SetComboSafe(cmbVendor, 0)
        SetComboSafe(cmbVoltage, 0)
        SetComboSafe(cmbAirQuality, 0)
        SetComboSafe(cmbEnvironment, 0)

        txtJobName.Text = "Cylinder-01"
        txtBore.Text = "20"
        txtStroke.Text = "100"
        txtQty.Text = "1"
        txtPressure.Text = "6"
        txtTubeLen.Text = "2"
        txtCyclesPerMin.Text = "5"

        chkFRL.Checked = True
        chkBranchReg.Checked = False
        chkQuickExhaust.Checked = False
        chkSensors.Checked = True
        chkManualOverride.Checked = True
        chkDumpValve.Checked = False

        AddHandler cmbProjectMode.SelectedIndexChanged, AddressOf ProjectModeChanged

        inputContent.AutoScrollMinSize = New Size(0, y + 20)
    End Sub

    Private Function CreateMiniButton(ByVal t As String, ByVal h As EventHandler) As Button
        Dim b As New Button()
        b.Text = t
        b.Width = 60
        b.Height = 24
        b.FlatStyle = FlatStyle.Flat
        b.FlatAppearance.BorderSize = 1
        b.BackColor = Color.FromArgb(60, 70, 90)
        b.ForeColor = Color.White
        AddHandler b.Click, h
        Return b
    End Function

    Private Function CreateOutputTextBox() As TextBox
        Dim tb As New TextBox()
        tb.Dock = DockStyle.Fill
        tb.Multiline = True
        tb.ScrollBars = ScrollBars.Vertical
        tb.ReadOnly = True
        tb.BorderStyle = BorderStyle.FixedSingle
        tb.BackColor = themePanel
        tb.ForeColor = themeText
        tb.Font = New Font("Consolas", 9.0F, FontStyle.Regular)
        Return tb
    End Function

    Private Function CreateDockGroup(title As String, bodyControl As Control) As Panel
        Dim host As New Panel()
        host.Dock = DockStyle.Fill
        host.Margin = New Padding(4, 4, 4, 4)
        host.Padding = New Padding(0)
        host.BackColor = themeBG

        Dim lbl As New Label()
        lbl.Text = title
        lbl.Dock = DockStyle.Top
        lbl.Height = 18
        lbl.ForeColor = themeAccent
        lbl.BackColor = themeBG
        lbl.TextAlign = ContentAlignment.MiddleLeft

        Dim panelBody As New Panel()
        panelBody.Dock = DockStyle.Fill
        panelBody.BackColor = themeBG
        bodyControl.Dock = DockStyle.Fill
        panelBody.Controls.Add(bodyControl)

        host.Controls.Add(panelBody)
        host.Controls.Add(lbl)
        Return host
    End Function

    Private Function CreatePreviewGroup() As Panel
        Dim host As New Panel()
        host.Dock = DockStyle.Fill
        host.Margin = New Padding(4, 4, 4, 4)
        host.BackColor = themeBG

        Dim lbl As New Label()
        lbl.Text = "CIRCUIT PREVIEW"
        lbl.Dock = DockStyle.Top
        lbl.Height = 18
        lbl.ForeColor = themeAccent
        lbl.BackColor = themeBG
        lbl.TextAlign = ContentAlignment.MiddleLeft
        host.Controls.Add(lbl)

        pnlPreviewToolbar = New Panel()
        pnlPreviewToolbar.Dock = DockStyle.Top
        pnlPreviewToolbar.Height = 28
        pnlPreviewToolbar.BackColor = themeBG
        host.Controls.Add(pnlPreviewToolbar)

        Dim lblMode As New Label()
        lblMode.Text = "Preview Mode:"
        lblMode.AutoSize = True
        lblMode.Location = New Point(0, 6)
        lblMode.ForeColor = themeText
        pnlPreviewToolbar.Controls.Add(lblMode)

        cmbPreviewMode = New ComboBox()
        cmbPreviewMode.DropDownStyle = ComboBoxStyle.DropDownList
        cmbPreviewMode.Items.AddRange(New Object() {"Standard", "Smart"})
        cmbPreviewMode.SelectedIndex = 0
        cmbPreviewMode.Width = 110
        cmbPreviewMode.Location = New Point(92, 2)
        AddHandler cmbPreviewMode.SelectedIndexChanged, AddressOf PreviewModeChanged
        pnlPreviewToolbar.Controls.Add(cmbPreviewMode)

        picPreview = New PictureBox()
        picPreview.Dock = DockStyle.Fill
        picPreview.BackColor = Color.Black
        picPreview.BorderStyle = BorderStyle.FixedSingle
        picPreview.SizeMode = PictureBoxSizeMode.Zoom
        host.Controls.Add(picPreview)

        Return host
    End Function

    Private Sub PreviewModeChanged(sender As Object, e As EventArgs)
        If cmbPreviewMode Is Nothing Then Exit Sub
        lastPreviewMode = cmbPreviewMode.Text
        If lastCalcOk Then DrawPreview()
    End Sub

    ' ============================================================
    ' THEME APPLY
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
    End Sub

    Private Sub ApplyThemeToThisForm()
        Try
            Me.BackColor = themeBG
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
            If btnPreview IsNot Nothing Then btnPreview.ForeColor = Color.Black
            If btnExportDxfSmart IsNot Nothing Then btnExportDxfSmart.ForeColor = Color.Black
            If btnExportDxfStd IsNot Nothing Then btnExportDxfStd.ForeColor = Color.Black
            If btnPdf IsNot Nothing Then btnPdf.ForeColor = Color.Black
            If btnReset IsNot Nothing Then btnReset.ForeColor = Color.Black
        Catch
        End Try
    End Sub

    Private Sub ApplyThemeRecursive(root As Control)
        If root Is Nothing Then Exit Sub

        For Each c As Control In root.Controls
            If TypeOf c Is Panel Then
                Dim p As Panel = DirectCast(c, Panel)
                If p Is pnlInput OrElse p Is inputContent OrElse p Is btnBar Then
                    p.BackColor = themePanel
                ElseIf p Is pnlPreviewToolbar Then
                    p.BackColor = themeBG
                Else
                    If p.BackColor <> Color.Black Then p.BackColor = themeBG
                End If
            ElseIf TypeOf c Is Label Then
                Dim l As Label = DirectCast(c, Label)
                If l.ForeColor <> themeAccent Then l.ForeColor = themeText
            ElseIf TypeOf c Is TextBox Then
                Dim tb As TextBox = DirectCast(c, TextBox)
                tb.BackColor = If(tb.ReadOnly, themePanel, themeBG)
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
            ElseIf TypeOf c Is SplitContainer Then
                c.BackColor = themeBG
            ElseIf TypeOf c Is ListBox Then
                Dim lb As ListBox = DirectCast(c, ListBox)
                lb.BackColor = themeBG
                lb.ForeColor = themeText
            End If

            If c.HasChildren Then ApplyThemeRecursive(c)
        Next
    End Sub

    ' ============================================================
    ' INPUT HELPERS
    ' ============================================================
    Private Sub AddSection(t As String, y As Integer)
        Dim lbl As New Label()
        lbl.Text = t
        lbl.ForeColor = themeAccent
        lbl.Location = New Point(10, y)
        lbl.AutoSize = True
        inputContent.Controls.Add(lbl)
    End Sub

    Private Sub AddLabel(t As String, y As Integer)
        Dim lbl As New Label()
        lbl.Text = t
        lbl.ForeColor = themeText
        lbl.Location = New Point(14, y + 3)
        lbl.AutoSize = True
        inputContent.Controls.Add(lbl)
    End Sub

    Private Function AddInput(t As String, y As Integer, tipText As String) As TextBox
        AddLabel(t, y)
        Dim tb As New TextBox()
        tb.Location = New Point(215, y)
        tb.Width = 145
        inputContent.Controls.Add(tb)

        RegisterTip(tb, tipText)
        AddHandler tb.MouseDown, AddressOf ArmTipsOnFirstClick
        Return tb
    End Function

    Private Function AddCombo(y As Integer, items() As String, tipText As String) As ComboBox
        Dim cb As New ComboBox()
        cb.Location = New Point(215, y)
        cb.Width = 145
        cb.DropDownStyle = ComboBoxStyle.DropDownList
        cb.Items.AddRange(items)
        If cb.Items.Count > 0 Then cb.SelectedIndex = 0
        inputContent.Controls.Add(cb)

        RegisterTip(cb, tipText)
        AddHandler cb.MouseDown, AddressOf ArmTipsOnFirstClick
        Return cb
    End Function

    Private Function AddCheck(t As String, y As Integer, tipText As String) As CheckBox
        Dim chk As New CheckBox()
        chk.Text = t
        chk.ForeColor = themeText
        chk.Location = New Point(14, y)
        chk.AutoSize = True
        inputContent.Controls.Add(chk)

        RegisterTip(chk, tipText)
        AddHandler chk.MouseDown, AddressOf ArmTipsOnFirstClick
        Return chk
    End Function

    Private Function CreateButton(t As String, c As Color, isPrimary As Boolean, h As EventHandler) As Button
        Dim b As New Button()
        b.Text = t
        b.BackColor = c
        b.ForeColor = If(isPrimary, Color.White, Color.Black)
        b.FlatStyle = FlatStyle.Flat
        b.FlatAppearance.BorderSize = 1
        b.Width = 102
        b.Height = 32
        b.Margin = New Padding(4, 4, 4, 4)
        AddHandler b.Click, h
        AddHandler b.MouseDown, AddressOf ArmTipsOnFirstClick
        Return b
    End Function

    Private Sub RegisterTip(c As Control, text As String)
        If c Is Nothing OrElse tipMap Is Nothing Then Exit Sub

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
        If Not tipMap.ContainsKey(c) Then Exit Sub
        Try
            tip.Hide(c)
            tip.Show(tipMap(c), c, Math.Max(10, c.Width + 6), 0, 7000)
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
                line.Append(" "c).Append(w)
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
    ' PHASE-2 JOBS (MULTI CYLINDER INDEPENDENT)
    ' ============================================================
    Private Sub ProjectModeChanged(sender As Object, e As EventArgs)
        UpdateJobsPanelState()
        If IsMultiMode() Then
            SetStatusWarn("Status: Multi Cylinder mode active (add/update jobs, then Calculate)")
        Else
            SetStatusOk("Status: Single Circuit mode active")
        End If
    End Sub

    Private Function IsMultiMode() As Boolean
        If cmbProjectMode Is Nothing Then Return False
        Return (cmbProjectMode.SelectedIndex = 1)
    End Function

    Private Sub UpdateJobsPanelState()
        Dim multi As Boolean = IsMultiMode()

        If lstCylinderJobs IsNot Nothing Then lstCylinderJobs.Enabled = multi
        If btnAddJob IsNot Nothing Then btnAddJob.Enabled = multi
        If btnUpdateJob IsNot Nothing Then btnUpdateJob.Enabled = multi
        If btnRemoveJob IsNot Nothing Then btnRemoveJob.Enabled = multi
        If btnLoadJob IsNot Nothing Then btnLoadJob.Enabled = multi
        If btnNewFromCurrent IsNot Nothing Then btnNewFromCurrent.Enabled = multi

        If lblJobsInfo IsNot Nothing Then
            lblJobsInfo.Text = "Jobs: " & jobs.Count.ToString(CultureInfo.InvariantCulture) & If(multi, " (Multi mode)", " (Ignored in Single mode)")
            lblJobsInfo.ForeColor = If(multi, themeOk, themeWarn)
        End If
    End Sub

    Private Function CaptureJobFromInputs() As CylinderJob
        Dim j As New CylinderJob()
        j.JobName = SafeText(txtJobName, "Cylinder-" & (jobs.Count + 1).ToString(CultureInfo.InvariantCulture))
        j.CylinderType = SafeComboText(cmbCylinderType)
        j.Bore = SafeText(txtBore, "20")
        j.Stroke = SafeText(txtStroke, "100")
        j.Qty = SafeText(txtQty, "1")
        j.Orientation = SafeComboText(cmbOrientation)
        j.MotionType = SafeComboText(cmbMotionType)

        j.OperationMode = SafeComboText(cmbOperationMode)
        j.FailSafe = SafeComboText(cmbFailSafe)
        j.SpeedPriority = SafeComboText(cmbSpeedPriority)
        j.Cushioning = SafeComboText(cmbCushioning)

        j.Vendor = SafeComboText(cmbVendor)
        j.Voltage = SafeComboText(cmbVoltage)

        j.Pressure = SafeText(txtPressure, "6")
        j.TubeLen = SafeText(txtTubeLen, "2")
        j.CyclesPerMin = SafeText(txtCyclesPerMin, "5")
        j.AirQuality = SafeComboText(cmbAirQuality)
        j.Environment = SafeComboText(cmbEnvironment)

        j.AddFRL = chkFRL.Checked
        j.AddBranchReg = chkBranchReg.Checked
        j.AddQuickExhaust = chkQuickExhaust.Checked
        j.AddSensors = chkSensors.Checked
        j.AddManualOverride = chkManualOverride.Checked
        j.AddDumpValve = chkDumpValve.Checked
        Return j
    End Function

    Private Sub ApplyJobToInputs(ByVal j As CylinderJob)
        If j Is Nothing Then Exit Sub

        txtJobName.Text = j.JobName
        SelectComboByText(cmbCylinderType, j.CylinderType)
        txtBore.Text = j.Bore
        txtStroke.Text = j.Stroke
        txtQty.Text = j.Qty
        SelectComboByText(cmbOrientation, j.Orientation)
        SelectComboByText(cmbMotionType, j.MotionType)

        SelectComboByText(cmbOperationMode, j.OperationMode)
        SelectComboByText(cmbFailSafe, j.FailSafe)
        SelectComboByText(cmbSpeedPriority, j.SpeedPriority)
        SelectComboByText(cmbCushioning, j.Cushioning)

        SelectComboByText(cmbVendor, j.Vendor)
        SelectComboByText(cmbVoltage, j.Voltage)

        txtPressure.Text = j.Pressure
        txtTubeLen.Text = j.TubeLen
        txtCyclesPerMin.Text = j.CyclesPerMin
        SelectComboByText(cmbAirQuality, j.AirQuality)
        SelectComboByText(cmbEnvironment, j.Environment)

        chkFRL.Checked = j.AddFRL
        chkBranchReg.Checked = j.AddBranchReg
        chkQuickExhaust.Checked = j.AddQuickExhaust
        chkSensors.Checked = j.AddSensors
        chkManualOverride.Checked = j.AddManualOverride
        chkDumpValve.Checked = j.AddDumpValve
    End Sub

    Private Sub RefreshJobsList()
        If lstCylinderJobs Is Nothing Then Exit Sub
        lstCylinderJobs.BeginUpdate()
        lstCylinderJobs.Items.Clear()
        For i As Integer = 0 To jobs.Count - 1
            lstCylinderJobs.Items.Add(jobs(i).DisplayText(i + 1))
        Next
        lstCylinderJobs.EndUpdate()

        If currentJobIndex >= 0 AndAlso currentJobIndex < lstCylinderJobs.Items.Count Then
            lstCylinderJobs.SelectedIndex = currentJobIndex
        End If

        UpdateJobsPanelState()
    End Sub

    Private Sub AddJobFromCurrent(sender As Object, e As EventArgs)
        If Not IsMultiMode() Then
            AddWarn("Switch to 'Multi Cylinder - Independent' mode first.")
            Exit Sub
        End If

        Dim j As CylinderJob = CaptureJobFromInputs()
        jobs.Add(j)
        currentJobIndex = jobs.Count - 1
        RefreshJobsList()
        SetStatusOk("Status: Added cylinder job (" & j.JobName & ")")
    End Sub

    Private Sub UpdateSelectedJobFromCurrent(sender As Object, e As EventArgs)
        If Not IsMultiMode() Then
            AddWarn("Switch to 'Multi Cylinder - Independent' mode first.")
            Exit Sub
        End If
        If currentJobIndex < 0 OrElse currentJobIndex >= jobs.Count Then
            AddWarn("Select a job to update.")
            Exit Sub
        End If

        jobs(currentJobIndex) = CaptureJobFromInputs()
        RefreshJobsList()
        SetStatusOk("Status: Updated selected job")
    End Sub

    Private Sub RemoveSelectedJob(sender As Object, e As EventArgs)
        If Not IsMultiMode() Then Exit Sub
        If currentJobIndex < 0 OrElse currentJobIndex >= jobs.Count Then
            AddWarn("Select a job to remove.")
            Exit Sub
        End If

        jobs.RemoveAt(currentJobIndex)
        If jobs.Count = 0 Then
            currentJobIndex = -1
        ElseIf currentJobIndex > jobs.Count - 1 Then
            currentJobIndex = jobs.Count - 1
        End If
        RefreshJobsList()
        SetStatusWarn("Status: Job removed")
    End Sub

    Private Sub LoadSelectedJobToInputs(sender As Object, e As EventArgs)
        If currentJobIndex < 0 OrElse currentJobIndex >= jobs.Count Then
            AddWarn("Select a job to load.")
            Exit Sub
        End If
        ApplyJobToInputs(jobs(currentJobIndex))
        SetStatusOk("Status: Loaded selected job into input fields")
    End Sub

    Private Sub DuplicateCurrentIntoNewJob(sender As Object, e As EventArgs)
        If Not IsMultiMode() Then Exit Sub
        Dim j As CylinderJob = CaptureJobFromInputs()
        j.JobName = j.JobName & "-COPY"
        jobs.Add(j)
        currentJobIndex = jobs.Count - 1
        RefreshJobsList()
        SetStatusOk("Status: Duplicated current inputs as new job")
    End Sub

    Private Sub JobsList_SelectedIndexChanged(sender As Object, e As EventArgs)
        If lstCylinderJobs Is Nothing Then Exit Sub
        currentJobIndex = lstCylinderJobs.SelectedIndex
    End Sub

    ' ============================================================
    ' CALC / LOGIC
    ' ============================================================
    Private Sub CalculateCircuit(sender As Object, e As EventArgs)
        SyncThemeFromCurrentUI()
        ClearOutputs()

        If IsMultiMode() Then
            CalculateProjectMode()
        Else
            CalculateSingleModeFromCurrentInputs(True)
        End If
    End Sub

    Private Sub CalculateProjectMode()
        If jobs.Count = 0 Then
            AddWarn("Multi mode is selected, but no jobs were added. Add jobs or switch to Single mode.")
            SetStatusWarn("Status: No jobs in Multi mode")
            lastCalcOk = False
            Exit Sub
        End If

        Dim sbFinal As New StringBuilder()
        Dim sbCalc As New StringBuilder()
        Dim sbBom As New StringBuilder()

        Dim totalCylQty As Integer = 0
        Dim totalValves As Integer = 0
        Dim totalSensors As Integer = 0
        Dim totalFCV As Integer = 0

        Dim commonFRL As Boolean = False
        Dim commonDump As Boolean = False

        sbFinal.AppendLine("PROJECT MODE: Multi Cylinder - Independent")
        sbFinal.AppendLine("Jobs Count                : " & jobs.Count.ToString(CultureInfo.InvariantCulture))
        sbFinal.AppendLine("Approach                  : Each job is treated as an independent circuit branch")
        sbFinal.AppendLine("Common supply             : Shared FRL / dump valve may be used at machine level (if selected)")
        sbFinal.AppendLine("DXF Preview/Export        : Shows current active/loaded circuit only (single branch preview)")
        sbFinal.AppendLine("")

        sbCalc.AppendLine("PHASE-2 LOGIC")
        sbCalc.AppendLine("1) Each job stores its own cylinder/valve/control settings.")
        sbCalc.AppendLine("2) Calculation merges project summary + per-job BOM lines.")
        sbCalc.AppendLine("3) Independent jobs typically require one valve per job (or manifold station per job).")
        sbCalc.AppendLine("4) Qty in each job = number of identical cylinders for that specific job.")
        sbCalc.AppendLine("5) FRL / dump valve may be centralized, not repeated per branch in final machine design.")
        sbCalc.AppendLine("")

        sbBom.AppendLine("PROJECT BOM (MERGED - DRAFT)")
        sbBom.AppendLine("NOTE: Review common items (FRL / dump valve / manifold) to avoid duplicates.")
        sbBom.AppendLine("")

        Dim anyFRL As Boolean = False
        Dim anyDump As Boolean = False
        Dim anyBranchReg As Boolean = False
        Dim anyQuickExhaust As Boolean = False

        For i As Integer = 0 To jobs.Count - 1
            Dim j As CylinderJob = jobs(i)
            Dim vendor As String = If(String.IsNullOrWhiteSpace(j.Vendor), "Generic", j.Vendor.Trim())

            Dim bore As Integer = SafeInt(j.Bore, 20)
            Dim stroke As Integer = SafeInt(j.Stroke, 100)
            Dim qty As Integer = Math.Max(1, SafeInt(j.Qty, 1))
            Dim isDA As Boolean = (j.CylinderType.IndexOf("Double", StringComparison.OrdinalIgnoreCase) >= 0)

            Dim portSize As String = RecommendPortSize(bore)
            Dim tubeOD As String = RecommendTubeOD(bore, If(String.IsNullOrWhiteSpace(j.SpeedPriority), "Smooth motion", j.SpeedPriority))

            totalCylQty += qty
            totalValves += 1
            totalFCV += If(isDA, 2, 1)
            If j.AddSensors Then
                totalSensors += If(isDA, 2, 1) * qty
            End If

            If j.AddFRL Then anyFRL = True
            If j.AddDumpValve Then anyDump = True
            If j.AddBranchReg Then anyBranchReg = True
            If j.AddQuickExhaust Then anyQuickExhaust = True

            sbFinal.AppendLine("Job " & (i + 1).ToString(CultureInfo.InvariantCulture) & " : " & If(String.IsNullOrWhiteSpace(j.JobName), "Cylinder-" & (i + 1).ToString(CultureInfo.InvariantCulture), j.JobName))
            sbFinal.AppendLine("  Type/Size               : " & If(isDA, "DA", "SA") & "  Ø" & bore.ToString(CultureInfo.InvariantCulture) & " x " & stroke.ToString(CultureInfo.InvariantCulture) & "  (Qty " & qty.ToString(CultureInfo.InvariantCulture) & ")")
            sbFinal.AppendLine("  Motion / Orientation    : " & j.MotionType & " / " & j.Orientation)
            sbFinal.AppendLine("  Valve / Port / Tube     : " & If(isDA, "5/2", "3/2") & " / " & portSize & " / " & tubeOD)
            sbFinal.AppendLine("  Vendor / Voltage        : " & vendor & " / " & j.Voltage)
            sbFinal.AppendLine("")

            sbCalc.AppendLine("Job " & (i + 1).ToString(CultureInfo.InvariantCulture) & ": " & If(String.IsNullOrWhiteSpace(j.JobName), "Unnamed", j.JobName))
            sbCalc.AppendLine(" - Independent branch circuit")
            sbCalc.AppendLine(" - Valve: " & If(isDA, "5/2", "3/2"))
            sbCalc.AppendLine(" - Meter-out flow control: " & If(isDA, "2 pcs", "1 pc"))
            sbCalc.AppendLine(" - Recommended port/tube: " & portSize & " / " & tubeOD)
            sbCalc.AppendLine("")

            sbBom.AppendLine("------------------------------------------------------------")
            sbBom.AppendLine("JOB " & (i + 1).ToString(CultureInfo.InvariantCulture) & ": " & If(String.IsNullOrWhiteSpace(j.JobName), "Cylinder-" & (i + 1).ToString(CultureInfo.InvariantCulture), j.JobName))
            sbBom.AppendLine("------------------------------------------------------------")
            sbBom.AppendLine(BuildVendorRecommendedBomTextEx( _
                vendor, j.Voltage, isDA, bore, stroke, qty, _
                j.AddFRL, j.AddBranchReg, j.AddQuickExhaust, j.AddSensors, j.AddManualOverride, j.AddDumpValve, _
                portSize, tubeOD))
            sbBom.AppendLine("")
        Next

        commonFRL = anyFRL
        commonDump = anyDump

        sbFinal.AppendLine("PROJECT SUMMARY")
        sbFinal.AppendLine("Total Cylinders (sum qty) : " & totalCylQty.ToString(CultureInfo.InvariantCulture))
        sbFinal.AppendLine("Estimated Valves          : " & totalValves.ToString(CultureInfo.InvariantCulture) & " (or manifold stations)")
        sbFinal.AppendLine("Estimated FCVs            : " & totalFCV.ToString(CultureInfo.InvariantCulture))
        sbFinal.AppendLine("Estimated Sensors         : " & totalSensors.ToString(CultureInfo.InvariantCulture))
        sbFinal.AppendLine("Common FRL suggested      : " & If(commonFRL, "Yes (shared machine supply)", "Optional"))
        sbFinal.AppendLine("Common Dump Valve         : " & If(commonDump, "Yes (shared machine safety/service)", "Optional"))
        sbFinal.AppendLine("Branch Regulators used    : " & If(anyBranchReg, "Some jobs request branch regulation", "No"))
        sbFinal.AppendLine("Quick Exhaust used        : " & If(anyQuickExhaust, "Some jobs request QEV", "No"))

        txtFinal.Text = sbFinal.ToString()
        txtCalc.Text = sbCalc.ToString()
        txtBom.Text = sbBom.ToString()

        lastBOMLines = New List(Of String)(txtBom.Text.Split(New String() {vbCrLf}, StringSplitOptions.None))
        lastSummaryTitle = "Pneumatic Project - Multi Cylinder (" & jobs.Count.ToString(CultureInfo.InvariantCulture) & " jobs)"
        lastCalcOk = True

        If currentJobIndex >= 0 AndAlso currentJobIndex < jobs.Count Then
            ApplyJobToInputs(jobs(currentJobIndex))
        ElseIf jobs.Count > 0 Then
            currentJobIndex = 0
            ApplyJobToInputs(jobs(0))
            RefreshJobsList()
        End If

        BuildLastPreviewStateFromCurrentInputs()
        DrawPreview()
        AddWarn("Multi mode active: preview/DXF currently show one selected branch only (not full multi-branch schematic yet).")
        SetStatusOk("Status: Project recommendation generated (Multi Cylinder mode)")
    End Sub

    Private Sub CalculateSingleModeFromCurrentInputs(ByVal writeOutputs As Boolean)
        Dim bore As Integer = CInt(Math.Round(ReadD(txtBore, 0)))
        Dim stroke As Integer = CInt(Math.Round(ReadD(txtStroke, 0)))
        Dim qty As Integer = CInt(Math.Round(ReadD(txtQty, 1)))
        Dim pressureBar As Double = ReadD(txtPressure, 6.0)
        Dim cpm As Double = ReadD(txtCyclesPerMin, 5.0)

        If bore <= 0 OrElse stroke <= 0 OrElse qty <= 0 Then
            AddError("Please enter valid Bore, Stroke and Cylinder Qty (all > 0).")
            SetStatusError("Status: Error")
            lastCalcOk = False
            Exit Sub
        End If

        If pressureBar <= 0 Then
            AddError("Supply pressure must be > 0 bar.")
            SetStatusError("Status: Error")
            lastCalcOk = False
            Exit Sub
        End If

        BuildLastPreviewStateFromCurrentInputs()

        Dim operationMode As String = If(cmbOperationMode Is Nothing, "Full Auto (PLC)", cmbOperationMode.Text)
        Dim failSafe As String = If(cmbFailSafe Is Nothing, "Retracted on power loss", cmbFailSafe.Text)
        Dim speedPriority As String = If(cmbSpeedPriority Is Nothing, "Smooth motion", cmbSpeedPriority.Text)
        Dim cushioning As String = If(cmbCushioning Is Nothing, "Preferred", cmbCushioning.Text)
        Dim orientation As String = If(cmbOrientation Is Nothing, "Horizontal", cmbOrientation.Text)
        Dim motionType As String = If(cmbMotionType Is Nothing, "General", cmbMotionType.Text)
        Dim env As String = If(cmbEnvironment Is Nothing, "Clean", cmbEnvironment.Text)
        Dim vendor As String = If(cmbVendor Is Nothing, "Generic", cmbVendor.Text)
        Dim coilV As String = If(cmbVoltage Is Nothing, "24VDC", cmbVoltage.Text)
        Dim airQ As String = If(cmbAirQuality Is Nothing, "General dry air", cmbAirQuality.Text)

        Dim valveFunction As String = If(lastIsDoubleActing, "5/2 directional valve", "3/2 directional valve")
        Dim valveActuation As String = RecommendValveActuation( _
            operationMode, _
            failSafe, _
            lastIsDoubleActing, _
            orientation, _
            motionType)

        Dim coilControl As String = RecommendCoilControl(operationMode, coilV)

        lastLineTags = BuildLineTags(lastIsDoubleActing, chkQuickExhaust.Checked)

        If pressureBar > 8 Then AddWarn("Pressure > 8 bar. Verify component ratings and regulator settings.")
        If pressureBar < 4 Then AddWarn("Low pressure may reduce cylinder force and speed.")
        If env = "Wet" OrElse env = "Washdown" Then AddWarn("Wet/washdown environment: verify corrosion resistance, seals and IP ratings.")
        If chkSensors.Checked AndAlso Not lastIsDoubleActing Then AddWarn("Single-acting cylinders may use one or two sensors depending application and cylinder options.")
        If speedPriority = "Fast cycle" AndAlso Not chkQuickExhaust.Checked Then AddWarn("Fast cycle selected: consider enabling Quick Exhaust Valve.")
        If cpm > 30 Then AddWarn("High cycles/min detected. Verify valve Cv, tubing, and thermal duty.")
        If orientation = "Vertical Up" OrElse orientation = "Vertical Down" Then AddWarn("Vertical motion: verify load holding / safety requirements.")
        If airQ = "Unknown" Then AddWarn("Air quality unknown. Add FRL / filtration and validate air prep.")

        If writeOutputs Then
            Dim sbFinal As New StringBuilder()
            sbFinal.AppendLine("Mode                      : Single Circuit (Representative)")
            sbFinal.AppendLine("Circuit Type              : " & If(lastIsDoubleActing, "Double-Acting Cylinder Control (Option A)", "Single-Acting Cylinder Control (Spring Return)"))
            sbFinal.AppendLine("Cylinder                  : " & If(lastIsDoubleActing, "Double-acting ", "Single-acting ") & "Ø" & bore.ToString(CultureInfo.InvariantCulture) & " x " & stroke.ToString(CultureInfo.InvariantCulture) & " mm (Qty " & qty.ToString(CultureInfo.InvariantCulture) & ")")
            sbFinal.AppendLine("Valve                     : " & valveFunction)
            sbFinal.AppendLine("Actuation                 : " & valveActuation)
            sbFinal.AppendLine("Coil / Control            : " & coilControl)
            sbFinal.AppendLine("Port Size (starting point): " & lastPortSize)
            sbFinal.AppendLine("Tubing OD (starting point): " & lastTubeOD)
            sbFinal.AppendLine("Flow Controls             : " & If(lastIsDoubleActing, "Meter-out on both cylinder ports (A and B)", "Meter-out on working port"))
            sbFinal.AppendLine("Quick Exhaust             : " & If(chkQuickExhaust.Checked, "Yes (speed-critical use)", "Optional"))
            sbFinal.AppendLine("FRL                       : " & If(chkFRL.Checked, "Mini FRL recommended + gauge", "Optional"))
            sbFinal.AppendLine("Branch Regulator          : " & If(chkBranchReg.Checked, "Branch regulator + gauge recommended", "Optional"))
            sbFinal.AppendLine("Sensors                   : " & If(chkSensors.Checked, "Add end sensors (S1/S2) if cylinder supports it", "Not included"))
            sbFinal.AppendLine("Cushioning                : " & cushioning)
            sbFinal.AppendLine("Operation Mode            : " & operationMode)
            sbFinal.AppendLine("Fail-safe Preference      : " & failSafe)
            sbFinal.AppendLine("Orientation / Motion      : " & orientation & " / " & motionType)
            sbFinal.AppendLine("Supply / Air / Env        : " & pressureBar.ToString("0.##", CultureInfo.InvariantCulture) & " bar / " & airQ & " / " & env)
            sbFinal.AppendLine("Vendor BOM Mode           : " & vendor)
            txtFinal.Text = sbFinal.ToString()

            Dim sbCalc As New StringBuilder()
            sbCalc.AppendLine("1) Baseline circuit = " & If(lastIsDoubleActing, "Double-acting cylinder + 5/2 valve (Option A).", "Single-acting spring-return cylinder + 3/2 valve."))
            sbCalc.AppendLine("2) Meter-out means using a flow control valve on the exhaust side of the cylinder port (speed control).")
            sbCalc.AppendLine("3) " & If(lastIsDoubleActing, "Flow controls default to meter-out on both A/B ports for stable speed control.", "One flow control on working port; exhaust silencing on valve/exhaust path."))
            sbCalc.AppendLine("4) Valve actuation selected from operation mode + fail-safe preference.")
            sbCalc.AppendLine("5) Port/tube size are starting recommendations only; verify by required force, speed, Cv and cycle rate.")
            sbCalc.AppendLine("6) Standard DXF preview/export uses symbol-library style + orthogonal routing.")
            sbCalc.AppendLine("7) Smart DXF preview/export uses compact logic layout for quick communication.")
            sbCalc.AppendLine("8) Vendor BOM shows recommended SERIES + sample PN style; final configuration must be verified in vendor configurator.")
            sbCalc.AppendLine("9) Selected layout: Supply left -> valve middle -> cylinder right/down.")
            txtCalc.Text = sbCalc.ToString()

            txtBom.Text = BuildVendorRecommendedBomText(lastIsDoubleActing, bore, stroke, qty, chkFRL.Checked, chkBranchReg.Checked, chkQuickExhaust.Checked, chkSensors.Checked, chkManualOverride.Checked, chkDumpValve.Checked, lastPortSize, lastTubeOD)
            lastBOMLines = New List(Of String)(txtBom.Text.Split(New String() {vbCrLf}, StringSplitOptions.None))

            lastSummaryTitle = "Pneumatic Circuit - " & If(lastIsDoubleActing, "DA", "SA") & " / " & vendor
        End If

        lastCalcOk = True
        SetStatusOk("Status: Circuit recommendation generated")
        DrawPreview()
    End Sub

    Private Sub BuildLastPreviewStateFromCurrentInputs()
        Dim bore As Integer = CInt(Math.Round(ReadD(txtBore, 20)))
        Dim speedPriority As String = If(cmbSpeedPriority Is Nothing, "Smooth motion", cmbSpeedPriority.Text)

        lastIsDoubleActing = True
        If cmbCylinderType IsNot Nothing Then
            lastIsDoubleActing = (cmbCylinderType.Text.IndexOf("Double", StringComparison.OrdinalIgnoreCase) >= 0)
        End If

        lastPortSize = RecommendPortSize(bore)
        lastTubeOD = RecommendTubeOD(bore, speedPriority)

        lastDrawFRL = chkFRL.Checked
        lastDrawBranchReg = chkBranchReg.Checked
        lastDrawQuickExhaust = chkQuickExhaust.Checked
        lastDrawSensors = chkSensors.Checked

        Dim tagName As String = SafeText(txtJobName, "CYL-01")
        If tagName.Trim() = "" Then tagName = "CYL-01"
        lastCylinderLabel = tagName.ToUpperInvariant()
        lastValveLabel = "SV-01"
        lastLineTags = BuildLineTags(lastIsDoubleActing, chkQuickExhaust.Checked)
    End Sub

    Private Function RecommendPortSize(ByVal bore As Integer) As String
        If bore <= 16 Then Return "M5"
        If bore <= 32 Then Return "G1/8"
        If bore <= 63 Then Return "G1/4"
        Return "G3/8"
    End Function

    Private Function RecommendTubeOD(ByVal bore As Integer, ByVal speedPriority As String) As String
        Dim fast As Boolean = (speedPriority = "Fast cycle")
        If bore <= 16 Then Return If(fast, "6 mm", "4 mm")
        If bore <= 32 Then Return If(fast, "8 mm", "6 mm")
        If bore <= 63 Then Return If(fast, "10 mm", "8 mm")
        Return "12 mm"
    End Function

    Private Function RecommendValveActuation( _
        ByVal opMode As String, _
        ByVal failSafe As String, _
        ByVal isDA As Boolean, _
        ByVal orientation As String, _
        ByVal motionType As String) As String

        Dim isPLC As Boolean = (opMode.IndexOf("PLC", StringComparison.OrdinalIgnoreCase) >= 0)
        Dim isSemi As Boolean = (opMode.IndexOf("Semi", StringComparison.OrdinalIgnoreCase) >= 0)
        Dim wantsHoldLast As Boolean = (failSafe.IndexOf("Hold last", StringComparison.OrdinalIgnoreCase) >= 0)
        Dim wantsRetractOnLoss As Boolean = (failSafe.IndexOf("Retracted", StringComparison.OrdinalIgnoreCase) >= 0)
        Dim wantsExtendOnLoss As Boolean = (failSafe.IndexOf("Extended", StringComparison.OrdinalIgnoreCase) >= 0)

        Dim isVertical As Boolean = False
        If orientation IsNot Nothing Then
            isVertical = (orientation.IndexOf("Vertical", StringComparison.OrdinalIgnoreCase) >= 0)
        End If

        Dim isClampLike As Boolean = False
        If motionType IsNot Nothing Then
            Dim mt As String = motionType.Trim().ToUpperInvariant()
            isClampLike = (mt = "CLAMP")
        End If

        If Not isDA Then
            If isPLC Then
                If wantsHoldLast Then
                    Return "3/2 single solenoid + spring return (Hold-last not typical for single-acting; verify requirement)"
                End If

                If wantsExtendOnLoss Then
                    Return "3/2 valve selection depends on NC/NO fail state (verify normal position for power-loss behavior)"
                End If

                Return "3/2 single solenoid + spring return"
            End If

            If isSemi Then
                Return "3/2 solenoid valve with manual override + spring return"
            End If

            Return "3/2 manual / mechanically actuated valve (spring return)"
        End If

        If Not isPLC AndAlso Not isSemi Then
            If wantsHoldLast Then
                Return "5/2 double solenoid (bistable) OR manual detent valve (verify safety)"
            End If
            Return "5/2 single solenoid + spring return (or manual valve with spring return)"
        End If

        If isSemi Then
            If wantsHoldLast Then
                Return "5/2 double solenoid (bistable) with manual override (verify safety)"
            End If
            Return "5/2 single solenoid + spring return with manual override"
        End If

        If isPLC Then
            If wantsHoldLast Then
                If isVertical Then
                    Return "5/2 double solenoid (bistable) + WARNING: vertical load may need lock/check/rod-lock (verify safety)"
                End If
                Return "5/2 double solenoid (bistable)"
            End If

            If wantsRetractOnLoss OrElse wantsExtendOnLoss Then
                If isVertical Then
                    Return "5/2 single solenoid + spring return (monostable) + WARNING: verify load-holding safety for vertical axis"
                End If
                Return "5/2 single solenoid + spring return (monostable)"
            End If

            If isClampLike Then
                If isVertical Then
                    Return "5/2 single solenoid + spring return (clamp default) + verify safe state and load holding"
                End If
                Return "5/2 single solenoid + spring return (clamp default)"
            End If

            If isVertical Then
                Return "5/2 single solenoid + spring return (default) + WARNING: vertical motion requires safety review"
            End If

            Return "5/2 single solenoid + spring return (default)"
        End If

        Return "5/2 single solenoid + spring return"
    End Function

    Private Function RecommendCoilControl(ByVal opMode As String, ByVal coilV As String) As String
        If opMode.IndexOf("PLC", StringComparison.OrdinalIgnoreCase) >= 0 Then
            Return coilV & " coil recommended"
        End If
        Return "Manual or " & coilV & " solenoid option"
    End Function

    Private Function BuildLineTags(ByVal isDA As Boolean, ByVal withQev As Boolean) As List(Of String)
        Dim tags As New List(Of String)()
        tags.Add("P001")
        tags.Add("P002")
        If isDA Then tags.Add("P003")
        tags.Add("P004")
        If isDA Then tags.Add("P005")
        If withQev Then tags.Add("QEV-01")
        Return tags
    End Function

    ' ============================================================
    ' VENDOR BOM ENGINE
    ' ============================================================
    Private Function BuildVendorRecommendedBomText( _
        ByVal isDoubleActing As Boolean, _
        ByVal boreMm As Integer, _
        ByVal strokeMm As Integer, _
        ByVal qty As Integer, _
        ByVal includeFRL As Boolean, _
        ByVal includeBranchReg As Boolean, _
        ByVal includeQuickExhaust As Boolean, _
        ByVal includeSensors As Boolean, _
        ByVal includeManualOverride As Boolean, _
        ByVal includeDumpValve As Boolean, _
        ByVal portSize As String, _
        ByVal tubeOD As String) As String

        Dim vendor As String = "Generic"
        If cmbVendor IsNot Nothing AndAlso cmbVendor.Text IsNot Nothing Then vendor = cmbVendor.Text.Trim()

        Dim coilV As String = "24VDC"
        If cmbVoltage IsNot Nothing AndAlso cmbVoltage.Text IsNot Nothing AndAlso cmbVoltage.Text.Trim() <> "" Then
            coilV = cmbVoltage.Text.Trim()
        End If

        Return BuildVendorRecommendedBomTextEx(vendor, coilV, isDoubleActing, boreMm, strokeMm, qty, includeFRL, includeBranchReg, includeQuickExhaust, includeSensors, includeManualOverride, includeDumpValve, portSize, tubeOD)
    End Function

    Private Function BuildVendorRecommendedBomTextEx( _
        ByVal vendor As String, _
        ByVal coilV As String, _
        ByVal isDoubleActing As Boolean, _
        ByVal boreMm As Integer, _
        ByVal strokeMm As Integer, _
        ByVal qty As Integer, _
        ByVal includeFRL As Boolean, _
        ByVal includeBranchReg As Boolean, _
        ByVal includeQuickExhaust As Boolean, _
        ByVal includeSensors As Boolean, _
        ByVal includeManualOverride As Boolean, _
        ByVal includeDumpValve As Boolean, _
        ByVal portSize As String, _
        ByVal tubeOD As String) As String

        Select Case vendor.Trim().ToUpperInvariant()
            Case "FESTO"
                Return BuildFestoBom(isDoubleActing, boreMm, strokeMm, qty, includeFRL, includeBranchReg, includeQuickExhaust, includeSensors, includeManualOverride, includeDumpValve, portSize, tubeOD, coilV)
            Case "SMC"
                Return BuildSmcBom(isDoubleActing, boreMm, strokeMm, qty, includeFRL, includeBranchReg, includeQuickExhaust, includeSensors, includeManualOverride, includeDumpValve, portSize, tubeOD, coilV)
            Case Else
                Return BuildGenericBom(isDoubleActing, boreMm, strokeMm, qty, includeFRL, includeBranchReg, includeQuickExhaust, includeSensors, includeManualOverride, includeDumpValve, portSize, tubeOD, coilV)
        End Select
    End Function

    Private Function BuildGenericBom( _
        ByVal isDA As Boolean, ByVal boreMm As Integer, ByVal strokeMm As Integer, ByVal qty As Integer, _
        ByVal includeFRL As Boolean, ByVal includeBranchReg As Boolean, ByVal includeQuickExhaust As Boolean, _
        ByVal includeSensors As Boolean, ByVal includeManualOverride As Boolean, ByVal includeDumpValve As Boolean, _
        ByVal portSize As String, ByVal tubeOD As String, ByVal coilV As String) As String

        Dim sb As New StringBuilder()
        sb.AppendLine("VENDOR BOM (GENERIC - COPY READY)")
        sb.AppendLine("NOTE: Starting recommendation. Final selection depends on force/speed/Cv/environment.")
        sb.AppendLine("")

        If isDA Then
            sb.AppendLine(qty.ToString(CultureInfo.InvariantCulture) & "x CYL-01  Double-acting pneumatic cylinder Ø" & boreMm.ToString(CultureInfo.InvariantCulture) & " x " & strokeMm.ToString(CultureInfo.InvariantCulture) & ", magnetic piston")
            sb.AppendLine("1x SV-01   5/2 directional valve, single solenoid + spring return, " & coilV & ", " & portSize)
            sb.AppendLine("2x FCV-01/02 Meter-out flow control valves (speed control)")
            sb.AppendLine("2x SIL-01/02 Silencers for valve exhaust ports (R/S)")
        Else
            sb.AppendLine(qty.ToString(CultureInfo.InvariantCulture) & "x CYL-01  Single-acting cylinder (spring return) Ø" & boreMm.ToString(CultureInfo.InvariantCulture) & " x " & strokeMm.ToString(CultureInfo.InvariantCulture) & ", magnetic piston if sensors required")
            sb.AppendLine("1x SV-01   3/2 directional valve, single solenoid + spring return, " & coilV & ", " & portSize)
            sb.AppendLine("1x FCV-01  Meter-out flow control valve (working port speed control)")
            sb.AppendLine("1x SIL-01  Silencer (exhaust)")
        End If

        If includeFRL Then sb.AppendLine("1x FRL-01  Filter/Regulator + Gauge (mini FRL or FR)")
        If includeBranchReg Then sb.AppendLine("1x REG-01  Branch regulator + gauge")
        If includeQuickExhaust Then sb.AppendLine("1x QEV-01  Quick exhaust valve (near cylinder if required)")
        If includeDumpValve Then sb.AppendLine("1x DUMP-01 Soft-start / dump valve")
        If includeSensors Then sb.AppendLine("Sensors    Add sensors + brackets/cables (qty depends on cylinder type and machine logic)")
        If includeManualOverride Then sb.AppendLine("1x NOTE    Manual override requested on valve")

        sb.AppendLine("Tubing set: " & tubeOD & " OD (cut lengths to suit)")
        sb.AppendLine("Fittings: Push-in fittings to suit " & tubeOD & " tube and " & portSize & " ports")

        Return sb.ToString()
    End Function

    Private Function BuildFestoBom( _
        ByVal isDA As Boolean, ByVal boreMm As Integer, ByVal strokeMm As Integer, ByVal qty As Integer, _
        ByVal includeFRL As Boolean, ByVal includeBranchReg As Boolean, ByVal includeQuickExhaust As Boolean, _
        ByVal includeSensors As Boolean, ByVal includeManualOverride As Boolean, ByVal includeDumpValve As Boolean, _
        ByVal portSize As String, ByVal tubeOD As String, ByVal coilV As String) As String

        Dim sb As New StringBuilder()
        sb.AppendLine("VENDOR BOM (FESTO - RECOMMENDED)")
        sb.AppendLine("NOTE: Series + sample PN style shown. Verify exact configuration in Festo configurator.")
        sb.AppendLine("")

        If isDA Then
            sb.AppendLine(qty.ToString(CultureInfo.InvariantCulture) & "x CYL-01  Festo ADN compact cylinder (double-acting), Ø" & boreMm.ToString(CultureInfo.InvariantCulture) & " x " & strokeMm.ToString(CultureInfo.InvariantCulture) & " mm")
            sb.AppendLine("    Sample PN style: ADN-" & boreMm.ToString(CultureInfo.InvariantCulture) & "-" & strokeMm.ToString(CultureInfo.InvariantCulture) & " (verify magnetic piston / cushioning / mounting)")
            sb.AppendLine("1x SV-01   Festo VUVG valve, 5/2 monostable, " & coilV & ", " & portSize)
            sb.AppendLine("2x FCV     Festo flow control valves (meter-out on A/B) - verify thread " & portSize)
            sb.AppendLine("2x SIL     Festo silencers for exhaust R/S")
        Else
            sb.AppendLine(qty.ToString(CultureInfo.InvariantCulture) & "x CYL-01  Festo single-acting spring-return cylinder, Ø" & boreMm.ToString(CultureInfo.InvariantCulture) & " x " & strokeMm.ToString(CultureInfo.InvariantCulture) & " mm")
            sb.AppendLine("    NOTE: Select exact single-acting Festo family based on mounting and sensing needs.")
            sb.AppendLine("1x SV-01   Festo 3/2 monostable valve, " & coilV & ", " & portSize)
            sb.AppendLine("1x FCV     Festo flow control valve (working port)")
            sb.AppendLine("1x SIL     Festo silencer")
        End If

        If includeFRL Then sb.AppendLine("1x FRL-01  Festo FR/FRL unit + gauge (size by flow)")
        If includeBranchReg Then sb.AppendLine("1x REG-01  Festo branch regulator + gauge")
        If includeQuickExhaust Then sb.AppendLine("1x QEV-01  Festo quick exhaust valve")
        If includeDumpValve Then sb.AppendLine("1x DUMP-01 Festo soft-start / dump valve (series by port size)")
        If includeSensors Then sb.AppendLine("Sensors    Festo cylinder sensors + mounting kit (verify slot/profile compatibility)")
        If includeManualOverride Then sb.AppendLine("1x NOTE    Prefer valve manual override option")

        sb.AppendLine("Tubing: " & tubeOD & " OD (Festo tubing, verify length/material)")
        sb.AppendLine("Fittings: Festo push-in fittings to suit " & tubeOD & " and " & portSize)

        Return sb.ToString()
    End Function

    Private Function BuildSmcBom( _
        ByVal isDA As Boolean, ByVal boreMm As Integer, ByVal strokeMm As Integer, ByVal qty As Integer, _
        ByVal includeFRL As Boolean, ByVal includeBranchReg As Boolean, ByVal includeQuickExhaust As Boolean, _
        ByVal includeSensors As Boolean, ByVal includeManualOverride As Boolean, ByVal includeDumpValve As Boolean, _
        ByVal portSize As String, ByVal tubeOD As String, ByVal coilV As String) As String

        Dim sb As New StringBuilder()
        sb.AppendLine("VENDOR BOM (SMC - RECOMMENDED)")
        sb.AppendLine("NOTE: Series + sample PN style shown. Verify exact configuration in SMC configurator.")
        sb.AppendLine("")

        If isDA Then
            sb.AppendLine(qty.ToString(CultureInfo.InvariantCulture) & "x CYL-01  SMC CQ2/CDQ2 compact cylinder (double-acting), Ø" & boreMm.ToString(CultureInfo.InvariantCulture) & " x " & strokeMm.ToString(CultureInfo.InvariantCulture) & " mm")
            sb.AppendLine("    Sample PN style: CDQ2" & boreMm.ToString(CultureInfo.InvariantCulture) & "-" & strokeMm.ToString(CultureInfo.InvariantCulture) & " (verify magnet, mounting, switch type)")
            sb.AppendLine("1x SV-01   SMC SY series 5-port solenoid valve (5/2), " & coilV & ", " & portSize)
            sb.AppendLine("2x FCV     SMC speed controllers (meter-out at A/B)")
            sb.AppendLine("2x SIL     SMC silencers for exhaust ports")
        Else
            sb.AppendLine(qty.ToString(CultureInfo.InvariantCulture) & "x CYL-01  SMC single-acting spring-return compact cylinder (select exact CQ2/compatible family), Ø" & boreMm.ToString(CultureInfo.InvariantCulture) & " x " & strokeMm.ToString(CultureInfo.InvariantCulture) & " mm")
            sb.AppendLine("1x SV-01   SMC 3-port solenoid valve (3/2 monostable), " & coilV & ", " & portSize)
            sb.AppendLine("1x FCV     SMC speed controller (working port)")
            sb.AppendLine("1x SIL     SMC silencer")
        End If

        If includeFRL Then sb.AppendLine("1x FRL-01  SMC air preparation unit + gauge (size by flow)")
        If includeBranchReg Then sb.AppendLine("1x REG-01  SMC branch regulator + gauge")
        If includeQuickExhaust Then sb.AppendLine("1x QEV-01  SMC quick exhaust valve")
        If includeDumpValve Then sb.AppendLine("1x DUMP-01 SMC soft-start / dump valve (verify series)")
        If includeSensors Then sb.AppendLine("Sensors    SMC auto switches + mounting brackets")
        If includeManualOverride Then sb.AppendLine("1x NOTE    Prefer valve manual override option")

        sb.AppendLine("Tubing: " & tubeOD & " OD (SMC tubing, verify length/material)")
        sb.AppendLine("Fittings: SMC one-touch fittings to suit " & tubeOD & " and " & portSize)

        Return sb.ToString()
    End Function

    ' ============================================================
    ' PREVIEW (GDI+)
    ' ============================================================
    Private Sub PreviewCircuit(sender As Object, e As EventArgs)
        If Not lastCalcOk Then
            AddWarn("Please calculate first before preview.")
            SetStatusWarn("Status: Preview requires calculation")
            RenderPreviewPlaceholder("Please calculate first.")
            Exit Sub
        End If
        DrawPreview()
        SetStatusOk("Status: Preview updated")
    End Sub

    Private Sub DrawPreview()
        If picPreview Is Nothing Then Exit Sub

        Dim w As Integer = Math.Max(400, picPreview.ClientSize.Width)
        Dim h As Integer = Math.Max(220, picPreview.ClientSize.Height)

        Dim bmp As New Bitmap(w, h)
        Dim g As Graphics = Graphics.FromImage(bmp)
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.Clear(Color.Black)

        Dim penLine As New Pen(Color.Lime, 1.4F)
        Dim penWhite As New Pen(Color.White, 1.0F)
        Dim penBorder As New Pen(Color.White, 0.8F)
        Dim brushYellow As Brush = Brushes.Yellow
        Dim brushRed As Brush = Brushes.Red
        Dim brushWhite As Brush = Brushes.White

        Dim fTitle As New Font("Segoe UI", 9.0F, FontStyle.Regular)
        Dim fTag As New Font("Consolas", 8.0F, FontStyle.Bold)
        Dim fLbl As New Font("Consolas", 8.0F, FontStyle.Regular)

        Dim header As String = "MetaMech Pneumatic Circuit Preview - " & lastPreviewMode
        If IsMultiMode() Then
            header &= " (Selected Branch / Representative)"
        End If
        g.DrawString(header, fTitle, brushYellow, 8, 6)

        Dim topY As Integer = 34
        Dim leftX As Integer = 20
        Dim rightX As Integer = w - 20
        Dim centerY As Integer = topY + Math.Max(80, (h - topY - 30) \ 2)

        If String.Equals(lastPreviewMode, "Smart", StringComparison.OrdinalIgnoreCase) Then
            DrawSmartPreview(g, penLine, penWhite, penBorder, fTag, fLbl, brushRed, brushWhite, leftX, centerY, rightX)
        Else
            DrawStandardPreview(g, penLine, penWhite, penBorder, fTag, fLbl, brushRed, brushWhite, leftX, topY, centerY, rightX, h - 10)
        End If

        g.Dispose()

        Dim oldImg As Image = picPreview.Image
        picPreview.Image = bmp
        If oldImg IsNot Nothing Then oldImg.Dispose()
    End Sub

    Private Sub RenderPreviewPlaceholder(msg As String)
        If picPreview Is Nothing Then Exit Sub
        Dim w As Integer = Math.Max(300, picPreview.ClientSize.Width)
        Dim h As Integer = Math.Max(180, picPreview.ClientSize.Height)

        Dim bmp As New Bitmap(w, h)
        Dim g As Graphics = Graphics.FromImage(bmp)
        g.Clear(Color.Black)
        g.DrawString(msg, New Font("Segoe UI", 10.0F, FontStyle.Regular), Brushes.Gainsboro, 20, 20)
        g.Dispose()

        Dim oldImg As Image = picPreview.Image
        picPreview.Image = bmp
        If oldImg IsNot Nothing Then oldImg.Dispose()
    End Sub

    Private Sub DrawSmartPreview(g As Graphics, penLine As Pen, penWhite As Pen, penBorder As Pen, fTag As Font, fLbl As Font, brRed As Brush, brWhite As Brush, leftX As Integer, centerY As Integer, rightX As Integer)
        Dim xFRL As Integer = leftX + 20
        Dim yFRL As Integer = centerY - 28
        Dim frlW As Integer = 90
        Dim frlH As Integer = 56

        Dim xValve As Integer = xFRL + 120
        Dim yValve As Integer = centerY - 34
        Dim valveW As Integer = If(lastIsDoubleActing, 140, 92)
        Dim valveH As Integer = 68

        Dim xCyl As Integer = Math.Min(rightX - 170, xValve + 230)
        Dim yCyl As Integer = centerY - 26
        Dim cylW As Integer = 160
        Dim cylH As Integer = 52

        g.DrawLine(penLine, leftX, centerY, xFRL, centerY)
        If lastDrawFRL Then
            DrawFRLSymbol(g, penWhite, fLbl, xFRL, yFRL, frlW, frlH)
            g.DrawString("FRL-01", fLbl, brWhite, xFRL + 12, yFRL - 14)
            g.DrawLine(penLine, xFRL + frlW, centerY, xValve, centerY)
            g.DrawString("P001", fTag, brRed, xFRL + frlW + 8, centerY - 16)
        Else
            g.DrawLine(penLine, xFRL, centerY, xValve, centerY)
            g.DrawString("P001", fTag, brRed, xFRL + 10, centerY - 16)
        End If

        If lastIsDoubleActing Then
            DrawValve52Symbol(g, penWhite, fLbl, xValve, yValve, valveW, valveH)
        Else
            DrawValve32Symbol(g, penWhite, fLbl, xValve, yValve, valveW, valveH)
        End If
        g.DrawString(lastValveLabel, fLbl, brWhite, xValve + 40, yValve - 14)
        g.DrawString("Y1", fLbl, brWhite, xValve - 20, yValve + 28)

        If lastIsDoubleActing Then
            Dim yA As Integer = yValve + 16
            Dim yB As Integer = yValve + valveH - 16

            Dim xFCV As Integer = xValve + valveW + 32
            Dim yFCVA As Integer = yA - 8
            Dim yFCVB As Integer = yB - 8

            g.DrawLine(penLine, xValve + valveW, yA, xFCV - 4, yA)
            g.DrawLine(penLine, xValve + valveW, yB, xFCV - 4, yB)
            g.DrawString("P002", fTag, brRed, xValve + valveW + 4, yA - 16)
            g.DrawString("P003", fTag, brRed, xValve + valveW + 4, yB - 16)

            DrawFlowControl(g, penWhite, fLbl, xFCV, yFCVA, "FCV-01")
            DrawFlowControl(g, penWhite, fLbl, xFCV, yFCVB, "FCV-02")

            g.DrawLine(penLine, xFCV + 26, yA, xCyl, yCyl + 14)
            g.DrawLine(penLine, xFCV + 26, yB, xCyl, yCyl + cylH - 14)
        Else
            Dim yA As Integer = centerY
            Dim xFCV As Integer = xValve + valveW + 32
            g.DrawLine(penLine, xValve + valveW, yA, xFCV - 4, yA)
            g.DrawString("P002", fTag, brRed, xValve + valveW + 4, yA - 16)
            DrawFlowControl(g, penWhite, fLbl, xFCV, yA - 8, "FCV-01")
            g.DrawLine(penLine, xFCV + 26, yA, xCyl, yCyl + cylH \ 2)
        End If

        DrawCylinder(g, penWhite, fLbl, xCyl, yCyl, cylW, cylH, lastIsDoubleActing)
        g.DrawString(lastCylinderLabel, fLbl, brWhite, xCyl + 25, yCyl - 14)

        If lastDrawSensors Then
            g.DrawString("S1", fLbl, brWhite, xCyl + 24, yCyl - 28)
            g.DrawRectangle(penWhite, xCyl + 20, yCyl - 14, 14, 10)
            g.DrawString("S2", fLbl, brWhite, xCyl + 86, yCyl - 28)
            g.DrawRectangle(penWhite, xCyl + 82, yCyl - 14, 14, 10)
        End If

        DrawExhaustBranch(g, penLine, penWhite, fLbl, xValve + 26, yValve + valveH, "P004")
        If lastIsDoubleActing Then
            DrawExhaustBranch(g, penLine, penWhite, fLbl, xValve + valveW - 26, yValve + valveH, "P005")
        End If
    End Sub

    Private Sub DrawStandardPreview(g As Graphics, penLine As Pen, penWhite As Pen, penBorder As Pen, fTag As Font, fLbl As Font, brRed As Brush, brWhite As Brush, leftX As Integer, topY As Integer, centerY As Integer, rightX As Integer, bottomY As Integer)
        g.DrawRectangle(penBorder, leftX, topY + 10, rightX - leftX, bottomY - (topY + 20))
        g.DrawString("MetaMech Pneumatic Circuit - STANDARD", fLbl, Brushes.Yellow, leftX + 8, topY + 16)
        g.DrawString("Layout: Supply left -> valve middle -> cylinder right/down", fLbl, Brushes.Yellow, leftX + 8, topY + 30)
        g.DrawString("Type: " & If(lastIsDoubleActing, "Double Acting", "Single Acting (Spring Return)"), fLbl, Brushes.Yellow, leftX + 8, topY + 44)

        Dim yLine As Integer = topY + 72
        Dim xFRL As Integer = leftX + 35
        Dim xValve As Integer = xFRL + 180
        Dim xCyl As Integer = xValve + 290

        g.DrawLine(penLine, leftX + 12, yLine, xCyl + 120, yLine)

        If lastDrawFRL Then
            DrawFRLSymbol(g, penWhite, fLbl, xFRL, yLine - 28, 95, 56)
            g.DrawString("FRL-01", fLbl, brWhite, xFRL + 18, yLine - 42)
        End If

        If lastDrawBranchReg Then
            g.DrawEllipse(penWhite, xFRL + 20, yLine + 40, 14, 14)
            g.DrawString("REG-01", fLbl, brWhite, xFRL + 40, yLine + 39)
        End If

        Dim xDrop As Integer = xValve + 45
        g.DrawLine(penLine, xDrop, yLine, xDrop, yLine + 18)
        g.DrawLine(penLine, xDrop, yLine + 18, xValve, yLine + 18)
        g.DrawString("P001", fTag, brRed, xDrop - 40, yLine - 14)

        Dim yValve As Integer = yLine + 18
        If lastIsDoubleActing Then
            DrawValve52Symbol(g, penWhite, fLbl, xValve, yValve, 120, 64)
        Else
            DrawValve32Symbol(g, penWhite, fLbl, xValve, yValve, 86, 64)
        End If
        g.DrawString("SV-01", fLbl, brWhite, xValve + 36, yValve - 14)
        g.DrawString("Y1", fLbl, brWhite, xValve - 18, yValve + 26)

        If lastIsDoubleActing Then
            Dim valveW As Integer = 120
            Dim xValveRight As Integer = xValve + valveW
            Dim yA As Integer = yValve + 16
            Dim yB As Integer = yValve + 48

            Dim xFCV As Integer = xValveRight + 38
            DrawFlowControl(g, penWhite, fLbl, xFCV, yA - 8, "FCV-01")
            DrawFlowControl(g, penWhite, fLbl, xFCV, yB - 8, "FCV-02")

            g.DrawLine(penLine, xValveRight, yA, xFCV - 4, yA)
            g.DrawLine(penLine, xValveRight, yB, xFCV - 4, yB)
            g.DrawString("P002", fTag, brRed, xValveRight + 6, yA - 16)
            g.DrawString("P003", fTag, brRed, xValveRight + 6, yB - 16)

            Dim xBranch As Integer = xFCV + 30
            g.DrawLine(penLine, xFCV + 26, yA, xBranch, yA)
            g.DrawLine(penLine, xFCV + 26, yB, xBranch, yB)

            Dim yCylTop As Integer = yValve + 26
            Dim yCyl As Integer = yCylTop + 6
            g.DrawLine(penLine, xBranch, yA, xBranch, yCyl + 14)
            g.DrawLine(penLine, xBranch, yB, xBranch, yCyl + 46)
            g.DrawLine(penLine, xBranch, yCyl + 14, xCyl, yCyl + 14)
            g.DrawLine(penLine, xBranch, yCyl + 46, xCyl, yCyl + 46)

            DrawCylinder(g, penWhite, fLbl, xCyl, yCyl, 150, 60, True)

            If lastDrawQuickExhaust Then
                g.DrawString("QEV-01", fLbl, brWhite, xBranch - 2, yA - 20)
            End If
        Else
            Dim valveW As Integer = 86
            Dim xValveRight As Integer = xValve + valveW
            Dim yA As Integer = yValve + 32
            Dim xFCV As Integer = xValveRight + 38

            g.DrawLine(penLine, xValveRight, yA, xFCV - 4, yA)
            g.DrawString("P002", fTag, brRed, xValveRight + 4, yA - 16)
            DrawFlowControl(g, penWhite, fLbl, xFCV, yA - 8, "FCV-01")

            Dim xBranch As Integer = xFCV + 34
            Dim yCyl As Integer = yValve + 24

            g.DrawLine(penLine, xFCV + 26, yA, xBranch, yA)
            g.DrawLine(penLine, xBranch, yA, xBranch, yCyl + 30)
            g.DrawLine(penLine, xBranch, yCyl + 30, xCyl, yCyl + 30)

            DrawCylinder(g, penWhite, fLbl, xCyl, yCyl, 150, 60, False)
        End If

        g.DrawString("CYL-01", fLbl, brWhite, xCyl + 44, yLine + 54)
        If lastDrawSensors Then
            g.DrawString("S1", fLbl, brWhite, xCyl + 22, yLine + 40)
            g.DrawRectangle(penWhite, xCyl + 18, yLine + 52, 14, 10)
            g.DrawString("S2", fLbl, brWhite, xCyl + 88, yLine + 40)
            g.DrawRectangle(penWhite, xCyl + 84, yLine + 52, 14, 10)
        End If

        DrawExhaustBranch(g, penLine, penWhite, fLbl, xValve + 26, yValve + 64, "P004")
        If lastIsDoubleActing Then
            DrawExhaustBranch(g, penLine, penWhite, fLbl, xValve + 94, yValve + 64, "P005")
        End If
    End Sub

    ' ============================================================
    ' DRAW SYMBOL HELPERS
    ' ============================================================
    Private Sub DrawFRLSymbol(g As Graphics, p As Pen, f As Font, x As Integer, y As Integer, w As Integer, h As Integer)
        Dim cellW As Integer = w \ 3
        g.DrawRectangle(p, x, y, w, h)
        g.DrawLine(p, x + cellW, y, x + cellW, y + h)
        g.DrawLine(p, x + 2 * cellW, y, x + 2 * cellW, y + h)
        g.DrawString("F", f, Brushes.White, x + 12, y + h \ 2 - 8)
        g.DrawString("R", f, Brushes.White, x + cellW + 10, y + h \ 2 - 8)
        g.DrawString("L", f, Brushes.White, x + 2 * cellW + 10, y + h \ 2 - 8)
    End Sub

    Private Sub DrawValve52Symbol(g As Graphics, p As Pen, f As Font, x As Integer, y As Integer, w As Integer, h As Integer)
        Dim cellW As Integer = w \ 2
        g.DrawRectangle(p, x, y, w, h)
        g.DrawLine(p, x + cellW, y, x + cellW, y + h)

        g.DrawLine(p, x + 12, y + h - 14, x + cellW - 12, y + 14)
        g.DrawLine(p, x + 12, y + h - 14, x + 22, y + h - 22)
        g.DrawLine(p, x + 12, y + h - 14, x + 20, y + h - 14)

        g.DrawLine(p, x + cellW + 12, y + 14, x + w - 12, y + h - 14)
        g.DrawLine(p, x + w - 20, y + h - 14, x + w - 12, y + h - 14)
        g.DrawLine(p, x + w - 20, y + h - 22, x + w - 12, y + h - 14)

        g.DrawString("R", f, Brushes.Yellow, x + 10, y + h + 2)
        g.DrawString("S", f, Brushes.Yellow, x + w - 18, y + h + 2)
        g.DrawString("A", f, Brushes.Yellow, x + w + 3, y + 10)
        g.DrawString("B", f, Brushes.Yellow, x + w + 3, y + h - 18)
    End Sub

    Private Sub DrawValve32Symbol(g As Graphics, p As Pen, f As Font, x As Integer, y As Integer, w As Integer, h As Integer)
        Dim cellW As Integer = w \ 2
        g.DrawRectangle(p, x, y, w, h)
        g.DrawLine(p, x + cellW, y, x + cellW, y + h)

        g.DrawLine(p, x + 10, y + 12, x + cellW - 10, y + h - 12)
        g.DrawLine(p, x + 10, y + h - 12, x + cellW - 10, y + 12)

        g.DrawLine(p, x + cellW + 10, y + h \ 2, x + w - 14, y + h \ 2)
        g.DrawLine(p, x + w - 22, y + h \ 2 - 6, x + w - 14, y + h \ 2)
        g.DrawLine(p, x + w - 22, y + h \ 2 + 6, x + w - 14, y + h \ 2)

        g.DrawString("R", f, Brushes.Yellow, x + 8, y + h + 2)
        g.DrawString("A", f, Brushes.Yellow, x + w + 3, y + h \ 2 - 8)
    End Sub

    Private Sub DrawFlowControl(g As Graphics, p As Pen, f As Font, x As Integer, y As Integer, labelText As String)
        Dim pts() As Point = {
            New Point(x, y + 8),
            New Point(x + 8, y),
            New Point(x + 20, y),
            New Point(x + 28, y + 8),
            New Point(x + 20, y + 16),
            New Point(x + 8, y + 16)
        }
        g.DrawPolygon(p, pts)
        g.DrawLine(p, x - 4, y + 8, x + 4, y + 8)
        g.DrawLine(p, x + 24, y + 8, x + 32, y + 8)
        g.DrawLine(p, x + 2, y + 18, x + 26, y - 2)
        g.DrawString(labelText, f, Brushes.Yellow, x + 2, y - 13)
    End Sub

    Private Sub DrawCylinder(g As Graphics, p As Pen, f As Font, x As Integer, y As Integer, w As Integer, h As Integer, isDA As Boolean)
        g.DrawRectangle(p, x, y, w, h)
        Dim rodX As Integer = x + w - 30
        g.DrawLine(p, rodX, y + 6, rodX, y + h - 6)
        g.DrawLine(p, x + w, y + h \ 2, x + w + 55, y + h \ 2)

        If Not isDA Then
            Dim sx As Integer = x + 12
            Dim sy As Integer = y + h \ 2
            g.DrawLine(p, sx, sy, sx + 8, sy - 8)
            g.DrawLine(p, sx + 8, sy - 8, sx + 16, sy + 8)
            g.DrawLine(p, sx + 16, sy + 8, sx + 24, sy - 8)
            g.DrawLine(p, sx + 24, sy - 8, sx + 32, sy + 8)
        End If
    End Sub

    Private Sub DrawExhaustBranch(g As Graphics, pLine As Pen, pWhite As Pen, f As Font, x As Integer, yTop As Integer, tagText As String)
        g.DrawLine(pLine, x, yTop, x, yTop + 28)
        g.DrawString(tagText, f, Brushes.Red, x - 16, yTop + 30)
        g.DrawRectangle(pWhite, x - 8, yTop + 34, 16, 8)
        g.DrawLine(pWhite, x - 6, yTop + 40, x + 6, yTop + 36)
        g.DrawLine(pWhite, x - 6, yTop + 36, x + 6, yTop + 40)
    End Sub

    ' ============================================================
    ' DXF EXPORTS (simple ASCII DXF)
    ' ============================================================
    Private Sub ExportDxfSmart(sender As Object, e As EventArgs)
        ExportDxfCommon(False)
    End Sub

    Private Sub ExportDxfStandard(sender As Object, e As EventArgs)
        ExportDxfCommon(True)
    End Sub

    Private Sub ExportDxfCommon(ByVal isStandard As Boolean)
        If Not lastCalcOk Then
            AddWarn("Please calculate first before exporting DXF.")
            SetStatusWarn("Status: DXF export requires calculation")
            Exit Sub
        End If

        If IsMultiMode() Then
            AddWarn("Multi mode: DXF export currently exports the selected/active branch only (not full multi-branch schematic).")
        End If

        Dim sfd As New SaveFileDialog()
        sfd.Filter = "DXF Files|*.dxf"
        sfd.FileName = If(isStandard, "PneumaticCircuit_Standard.dxf", "PneumaticCircuit_Smart.dxf")
        If sfd.ShowDialog() <> DialogResult.OK Then Exit Sub

        Try
            Dim d As New SimpleDxfWriter()
            d.StartFile()

            If isStandard Then
                WriteStandardDxf(d)
            Else
                WriteSmartDxf(d)
            End If

            d.EndFile()
            File.WriteAllText(sfd.FileName, d.ToString(), Encoding.ASCII)

            SetStatusOk("Status: DXF exported -> " & Path.GetFileName(sfd.FileName))
        Catch ex As Exception
            AddError("DXF export failed: " & ex.Message)
            SetStatusError("Status: DXF export error")
        End Try
    End Sub

    Private Sub WriteSmartDxf(d As SimpleDxfWriter)
        Dim leftX As Double = 10
        Dim centerY As Double = 60

        Dim xFRL As Double = 20
        Dim yFRL As Double = centerY - 10
        Dim xValve As Double = 70
        Dim yValve As Double = centerY - 12
        Dim xCyl As Double = 150
        Dim yCyl As Double = centerY - 10

        d.Text(8, 108, "MetaMech Pneumatic Circuit - Smart Export", 2.5, "TEXT_Y")
        If IsMultiMode() Then d.Text(8, 104, "Multi mode: selected branch export only", 1.8, "TEXT_Y")

        If lastDrawFRL Then
            DxfRect(d, xFRL, yFRL, 24, 20, "SYM_W")
            d.Line(xFRL + 8, yFRL, xFRL + 8, yFRL + 20, "SYM_W")
            d.Line(xFRL + 16, yFRL, xFRL + 16, yFRL + 20, "SYM_W")
            d.Text(xFRL + 2, yFRL + 8, "F", 2.5, "TEXT_W")
            d.Text(xFRL + 10, yFRL + 8, "R", 2.5, "TEXT_W")
            d.Text(xFRL + 18, yFRL + 8, "L", 2.5, "TEXT_W")
            d.Text(xFRL + 2, yFRL + 24, "FRL-01", 2.2, "TEXT_Y")
            d.Line(leftX, centerY, xFRL, centerY, "AIR_G")
            d.Line(xFRL + 24, centerY, xValve, centerY, "AIR_G")
            d.Text(xFRL + 27, centerY + 4, "P001", 2.2, "TEXT_R")
        Else
            d.Line(leftX, centerY, xValve, centerY, "AIR_G")
            d.Text(xFRL + 6, centerY + 4, "P001", 2.2, "TEXT_R")
        End If

        If lastIsDoubleActing Then
            DxfRect(d, xValve, yValve, 40, 24, "SYM_W")
            d.Line(xValve + 20, yValve, xValve + 20, yValve + 24, "SYM_W")
            d.Text(xValve + 10, yValve + 27, "SV-01", 2.2, "TEXT_Y")
            d.Text(xValve - 6, yValve + 8, "Y1", 2.0, "TEXT_Y")

            d.Line(xValve + 40, yValve + 18, 122, 72, "AIR_G")
            d.Line(xValve + 40, yValve + 6, 122, 48, "AIR_G")
            d.Text(xValve + 44, yValve + 20, "P002", 2.0, "TEXT_R")
            d.Text(xValve + 44, yValve + 8, "P003", 2.0, "TEXT_R")
            DxfFlowCtrl(d, 122, 70, "FCV-01")
            DxfFlowCtrl(d, 122, 46, "FCV-02")

            DxfCylinder(d, xCyl, yCyl, 46, 20, True)
            d.Text(xCyl + 8, yCyl + 24, lastCylinderLabel, 2.0, "TEXT_Y")
            If lastDrawSensors Then
                DxfRect(d, xCyl + 10, yCyl + 24, 4, 3, "SYM_W")
                DxfRect(d, xCyl + 22, yCyl + 24, 4, 3, "SYM_W")
                d.Text(xCyl + 10, yCyl + 29, "S1", 1.8, "TEXT_Y")
                d.Text(xCyl + 22, yCyl + 29, "S2", 1.8, "TEXT_Y")
            End If
        Else
            DxfRect(d, xValve, yValve, 28, 24, "SYM_W")
            d.Line(xValve + 14, yValve, xValve + 14, yValve + 24, "SYM_W")
            d.Text(xValve + 3, yValve + 27, "SV-01", 2.2, "TEXT_Y")
            d.Text(xValve - 6, yValve + 8, "Y1", 2.0, "TEXT_Y")

            d.Line(xValve + 28, centerY, 122, centerY, "AIR_G")
            d.Text(xValve + 32, centerY + 4, "P002", 2.0, "TEXT_R")
            DxfFlowCtrl(d, 122, centerY - 2, "FCV-01")

            DxfCylinder(d, xCyl, yCyl, 46, 20, False)
            d.Text(xCyl + 8, yCyl + 24, lastCylinderLabel, 2.0, "TEXT_Y")
        End If

        d.Line(xValve + 8, yValve, xValve + 8, yValve - 14, "AIR_G")
        DxfSilencer(d, xValve + 5, yValve - 18)
        d.Text(xValve + 4, yValve - 24, "P004", 1.8, "TEXT_R")
        If lastIsDoubleActing Then
            d.Line(xValve + 32, yValve, xValve + 32, yValve - 14, "AIR_G")
            DxfSilencer(d, xValve + 29, yValve - 18)
            d.Text(xValve + 28, yValve - 24, "P005", 1.8, "TEXT_R")
        End If
    End Sub

    Private Sub WriteStandardDxf(d As SimpleDxfWriter)
        d.Text(5, 140, "MetaMech Pneumatic Circuit - STANDARD DXF (Symbol Library)", 2.6, "TEXT_Y")
        d.Text(5, 135, "Layout: Supply left -> valve middle -> cylinder right/down", 1.9, "TEXT_Y")
        If IsMultiMode() Then d.Text(5, 130, "Multi mode: selected branch export only", 1.8, "TEXT_Y")

        DxfRect(d, 3, 20, 250, 108, "BORDER")
        d.Line(8, 116, 246, 116, "AIR_G")

        If lastDrawFRL Then
            DxfRect(d, 28, 102, 28, 14, "SYM_W")
            d.Line(28 + 9.3, 102, 28 + 9.3, 116, "SYM_W")
            d.Line(28 + 18.6, 102, 28 + 18.6, 116, "SYM_W")
            d.Text(31, 108, "F", 2.0, "TEXT_Y")
            d.Text(40, 108, "R", 2.0, "TEXT_Y")
            d.Text(49, 108, "L", 2.0, "TEXT_Y")
            d.Text(30, 119, "FRL-01", 1.9, "TEXT_Y")
        End If

        If lastDrawBranchReg Then
            d.Circle(38, 94, 2.2, "SYM_W")
            d.Text(42, 93, "REG-01", 1.8, "TEXT_Y")
        End If

        d.Line(78, 116, 78, 100, "AIR_G")
        d.Line(78, 100, 96, 100, "AIR_G")
        d.Text(70, 118, "P001", 1.8, "TEXT_R")

        If lastIsDoubleActing Then
            DxfRect(d, 96, 90, 34, 20, "SYM_W")
            d.Line(113, 90, 113, 110, "SYM_W")
            d.Text(104, 113, "SV-01", 1.9, "TEXT_Y")
            d.Text(90, 99, "Y1", 1.8, "TEXT_Y")

            d.Line(130, 105, 148, 105, "AIR_G")
            d.Line(130, 95, 148, 95, "AIR_G")
            d.Text(133, 107, "P002", 1.8, "TEXT_R")
            d.Text(133, 97, "P003", 1.8, "TEXT_R")

            DxfFlowCtrl(d, 148, 103, "FCV-01")
            DxfFlowCtrl(d, 148, 93, "FCV-02")

            d.Line(156, 105, 172, 105, "AIR_G")
            d.Line(156, 95, 172, 95, "AIR_G")
            d.Line(172, 105, 172, 77, "AIR_G")
            d.Line(172, 95, 172, 63, "AIR_G")
            d.Line(172, 77, 188, 77, "AIR_G")
            d.Line(172, 63, 188, 63, "AIR_G")

            DxfCylinder(d, 188, 56, 42, 28, True)
            d.Text(192, 86, lastCylinderLabel, 1.9, "TEXT_Y")
            If lastDrawSensors Then
                DxfRect(d, 196, 86, 4, 3, "SYM_W")
                DxfRect(d, 210, 86, 4, 3, "SYM_W")
                d.Text(196, 90, "S1", 1.6, "TEXT_Y")
                d.Text(210, 90, "S2", 1.6, "TEXT_Y")
            End If
            If lastDrawQuickExhaust Then
                d.Text(168, 108, "QEV-01", 1.8, "TEXT_Y")
            End If
        Else
            DxfRect(d, 96, 90, 26, 20, "SYM_W")
            d.Line(109, 90, 109, 110, "SYM_W")
            d.Text(99, 113, "SV-01", 1.9, "TEXT_Y")
            d.Text(90, 99, "Y1", 1.8, "TEXT_Y")

            d.Line(122, 100, 148, 100, "AIR_G")
            d.Text(126, 102, "P002", 1.8, "TEXT_R")

            DxfFlowCtrl(d, 148, 98, "FCV-01")
            d.Line(156, 100, 176, 100, "AIR_G")
            d.Line(176, 100, 176, 70, "AIR_G")
            d.Line(176, 70, 188, 70, "AIR_G")

            DxfCylinder(d, 188, 56, 42, 28, False)
            d.Text(192, 86, lastCylinderLabel, 1.9, "TEXT_Y")
        End If

        d.Line(104, 90, 104, 76, "AIR_G")
        DxfSilencer(d, 101, 72)
        d.Text(101, 68, "P004", 1.6, "TEXT_R")

        If lastIsDoubleActing Then
            d.Line(122, 90, 122, 76, "AIR_G")
            DxfSilencer(d, 119, 72)
            d.Text(119, 68, "P005", 1.6, "TEXT_R")
        End If
    End Sub

    Private Sub DxfRect(d As SimpleDxfWriter, x As Double, y As Double, w As Double, h As Double, layerName As String)
        d.Line(x, y, x + w, y, layerName)
        d.Line(x + w, y, x + w, y + h, layerName)
        d.Line(x + w, y + h, x, y + h, layerName)
        d.Line(x, y + h, x, y, layerName)
    End Sub

    Private Sub DxfFlowCtrl(d As SimpleDxfWriter, x As Double, y As Double, labelText As String)
        d.Line(x - 4, y, x, y, "AIR_G")
        d.Line(x, y, x + 4, y + 4, "SYM_W")
        d.Line(x + 4, y + 4, x + 10, y + 4, "SYM_W")
        d.Line(x + 10, y + 4, x + 14, y, "SYM_W")
        d.Line(x + 14, y, x + 10, y - 4, "SYM_W")
        d.Line(x + 10, y - 4, x + 4, y - 4, "SYM_W")
        d.Line(x + 4, y - 4, x, y, "SYM_W")
        d.Line(x + 1, y - 6, x + 13, y + 6, "SYM_W")
        d.Line(x + 14, y, x + 18, y, "AIR_G")
        d.Text(x + 1, y + 6, labelText, 1.6, "TEXT_Y")
    End Sub

    Private Sub DxfCylinder(d As SimpleDxfWriter, x As Double, y As Double, w As Double, h As Double, isDA As Boolean)
        DxfRect(d, x, y, w, h, "SYM_W")
        d.Line(x + w - 10, y + 2, x + w - 10, y + h - 2, "SYM_W")
        d.Line(x + w, y + h / 2, x + w + 16, y + h / 2, "SYM_W")
        If Not isDA Then
            d.Line(x + 4, y + h / 2, x + 8, y + h / 2 + 3, "SYM_W")
            d.Line(x + 8, y + h / 2 + 3, x + 12, y + h / 2 - 3, "SYM_W")
            d.Line(x + 12, y + h / 2 - 3, x + 16, y + h / 2 + 3, "SYM_W")
            d.Line(x + 16, y + h / 2 + 3, x + 20, y + h / 2 - 3, "SYM_W")
        End If
    End Sub

    Private Sub DxfSilencer(d As SimpleDxfWriter, x As Double, y As Double)
        DxfRect(d, x, y, 5, 2.5, "SYM_W")
        d.Line(x + 0.5, y + 0.3, x + 4.5, y + 2.0, "SYM_W")
        d.Line(x + 0.5, y + 2.0, x + 4.5, y + 0.3, "SYM_W")
    End Sub

    ' ============================================================
    ' PDF EXPORT (Option B: Microsoft Print to PDF) - MULTI PAGE
    ' ============================================================
    Private Sub ExportPdf(sender As Object, e As EventArgs)
        If Not lastCalcOk Then
            AddWarn("Please calculate first before PDF export.")
            SetStatusWarn("Status: PDF export requires calculation")
            Exit Sub
        End If

        Try
            Dim doc As New PrintDocument()
            doc.DocumentName = "Pneumatic Circuit Helper Report"

            Dim pdfPrinterName As String = FindMicrosoftPrintToPdfPrinter()
            If pdfPrinterName <> "" Then
                doc.PrinterSettings.PrinterName = pdfPrinterName
            End If

            Dim pd As New PrintDialog()
            pd.AllowSomePages = False
            pd.AllowSelection = False
            pd.UseEXDialog = True
            pd.Document = doc

            If pd.ShowDialog() <> DialogResult.OK Then
                SetStatusWarn("Status: PDF export cancelled")
                Exit Sub
            End If

            If doc.PrinterSettings.PrinterName Is Nothing OrElse doc.PrinterSettings.PrinterName.Trim() = "" Then
                AddWarn("No printer selected.")
                SetStatusWarn("Status: PDF export cancelled")
                Exit Sub
            End If

            If doc.PrinterSettings.PrinterName.IndexOf("Microsoft Print to PDF", StringComparison.OrdinalIgnoreCase) < 0 Then
                AddWarn("Selected printer is not 'Microsoft Print to PDF'. Output may not be a PDF.")
            End If

            PreparePrintReportLines()
            _printLineIndex = 0
            _printHeaderGenerated = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)

            doc.PrintController = New StandardPrintController()

            AddHandler doc.PrintPage, AddressOf PrintPneumaticReportPage
            AddHandler doc.EndPrint, AddressOf EndPrintCleanup

            doc.Print()

            SetStatusOk("Status: Sent to printer -> " & doc.PrinterSettings.PrinterName & " (Save PDF dialog may appear)")
        Catch ex As Exception
            AddError("PDF export failed: " & ex.Message)
            SetStatusError("Status: PDF export error")
        End Try
    End Sub

    Private Function FindMicrosoftPrintToPdfPrinter() As String
        Try
            For Each pNameObj As String In PrinterSettings.InstalledPrinters
                If pNameObj IsNot Nothing AndAlso pNameObj.IndexOf("Microsoft Print to PDF", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return pNameObj
                End If
            Next
        Catch
        End Try
        Return ""
    End Function

    Private Sub PreparePrintReportLines()
        Dim raw As New StringBuilder()

        raw.AppendLine("MetaMech Pneumatic Circuit Helper Report")
        raw.AppendLine("Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture))
        raw.AppendLine("")

        raw.AppendLine("[FINAL RECOMMENDATION]")
        raw.AppendLine(If(txtFinal Is Nothing, "", txtFinal.Text))
        raw.AppendLine("")

        raw.AppendLine("[LOGIC]")
        raw.AppendLine(If(txtCalc Is Nothing, "", txtCalc.Text))
        raw.AppendLine("")

        raw.AppendLine("[BOM]")
        raw.AppendLine(If(txtBom Is Nothing, "", txtBom.Text))
        raw.AppendLine("")

        raw.AppendLine("[WARNINGS]")
        raw.AppendLine(If(rtbWarn Is Nothing, "", rtbWarn.Text))

        Dim normalized As String = raw.ToString().Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        Dim arr() As String = normalized.Split(New String() {vbLf}, StringSplitOptions.None)

        _printLines = New List(Of String)()
        For i As Integer = 0 To arr.Length - 1
            _printLines.Add(arr(i))
        Next
        If _printLines.Count = 0 Then _printLines.Add("")
    End Sub

    Private Sub PrintPneumaticReportPage(sender As Object, ev As PrintPageEventArgs)
        If _printHeaderFont Is Nothing Then _printHeaderFont = New Font("Segoe UI", 13.0F, FontStyle.Bold)
        If _printSubFont Is Nothing Then _printSubFont = New Font("Segoe UI", 8.75F, FontStyle.Regular)
        If _printBodyFont Is Nothing Then _printBodyFont = New Font("Consolas", 8.6F, FontStyle.Regular)

        Dim g As Graphics = ev.Graphics
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

        Dim left As Single = ev.MarginBounds.Left
        Dim top As Single = ev.MarginBounds.Top
        Dim width As Single = ev.MarginBounds.Width
        Dim bottom As Single = ev.MarginBounds.Bottom

        Dim y As Single = top

        g.DrawString(_printHeaderTitle, _printHeaderFont, Brushes.Black, left, y)
        y += _printHeaderFont.GetHeight(g) + 2.0F

        g.DrawString("Generated: " & _printHeaderGenerated, _printSubFont, Brushes.Black, left, y)
        y += _printSubFont.GetHeight(g) + 2.0F

        Dim modeText As String = If(IsMultiMode(), _
            "Mode: Multi Cylinder - Independent (report summarizes project + selected branch preview)", _
            "Mode: Single Circuit (Representative)")
        g.DrawString(modeText, _printSubFont, Brushes.Black, left, y)
        y += _printSubFont.GetHeight(g) + 6.0F

        g.DrawLine(Pens.Black, left, y, left + width, y)
        y += 6.0F

        Dim footerHeight As Single = _printSubFont.GetHeight(g) + 8.0F
        Dim lineHeight As Single = _printBodyFont.GetHeight(g) + 1.0F
        If lineHeight < 10.0F Then lineHeight = 10.0F

        Dim availableTextHeight As Single = (bottom - y) - footerHeight
        If availableTextHeight < lineHeight Then
            ev.HasMorePages = False
            Exit Sub
        End If

        Dim maxLinesThisPage As Integer = CInt(Math.Floor(availableTextHeight / lineHeight))
        If maxLinesThisPage < 1 Then maxLinesThisPage = 1

        Dim countOnPage As Integer = 0

        While _printLineIndex < _printLines.Count AndAlso countOnPage < maxLinesThisPage
            Dim sourceLine As String = _printLines(_printLineIndex)
            Dim wrapped As List(Of String) = WrapPrintLineToWidth(g, sourceLine, _printBodyFont, width)

            Dim wi As Integer
            For wi = 0 To wrapped.Count - 1
                If countOnPage >= maxLinesThisPage Then Exit For
                g.DrawString(wrapped(wi), _printBodyFont, Brushes.Black, left, y)
                y += lineHeight
                countOnPage += 1
            Next

            If wrapped.Count > 0 Then
                If wi < wrapped.Count Then
                    Dim remain As New StringBuilder()
                    For r As Integer = wi To wrapped.Count - 1
                        If remain.Length > 0 Then remain.AppendLine()
                        remain.Append(wrapped(r))
                    Next
                    _printLines(_printLineIndex) = remain.ToString()
                    Exit While
                End If
            End If

            _printLineIndex += 1
        End While

        Dim footerY As Single = bottom - footerHeight + 2.0F
        Dim footerLeftText As String = "Printed via: " & SafePrinterNameForFooter(sender)
        Dim footerRightText As String = "MetaMech Pneumatic Circuit Helper"

        g.DrawString(footerLeftText, _printSubFont, Brushes.Black, left, footerY)
        Dim footerRightSize As SizeF = g.MeasureString(footerRightText, _printSubFont)
        g.DrawString(footerRightText, _printSubFont, Brushes.Black, left + width - footerRightSize.Width, footerY)

        ev.HasMorePages = (_printLineIndex < _printLines.Count)
    End Sub

    Private Function SafePrinterNameForFooter(sender As Object) As String
        Try
            Dim doc As PrintDocument = TryCast(sender, PrintDocument)
            If doc IsNot Nothing AndAlso doc.PrinterSettings IsNot Nothing AndAlso doc.PrinterSettings.PrinterName IsNot Nothing Then
                Return doc.PrinterSettings.PrinterName
            End If
        Catch
        End Try
        Return "Printer"
    End Function

    Private Function WrapPrintLineToWidth(ByVal g As Graphics, ByVal inputLine As String, ByVal f As Font, ByVal maxWidth As Single) As List(Of String)
        Dim result As New List(Of String)()

        If inputLine Is Nothing Then
            result.Add("")
            Return result
        End If

        If inputLine.IndexOf(vbLf, StringComparison.Ordinal) >= 0 OrElse inputLine.IndexOf(vbCr, StringComparison.Ordinal) >= 0 Then
            Dim normalized As String = inputLine.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim parts() As String = normalized.Split(New String() {vbLf}, StringSplitOptions.None)
            For i As Integer = 0 To parts.Length - 1
                Dim subLines As List(Of String) = WrapPrintLineToWidth(g, parts(i), f, maxWidth)
                For j As Integer = 0 To subLines.Count - 1
                    result.Add(subLines(j))
                Next
            Next
            Return result
        End If

        Dim indentCount As Integer = 0
        While indentCount < inputLine.Length AndAlso inputLine(indentCount) = " "c
            indentCount += 1
        End While
        Dim indent As String = New String(" "c, indentCount)

        If g.MeasureString(inputLine, f).Width <= maxWidth Then
            result.Add(inputLine)
            Return result
        End If

        Dim remaining As String = inputLine

        While remaining.Length > 0
            If g.MeasureString(remaining, f).Width <= maxWidth Then
                result.Add(remaining)
                Exit While
            End If

            Dim cut As Integer = remaining.Length - 1
            Dim found As Boolean = False

            While cut > 1
                Dim test As String = remaining.Substring(0, cut)
                If g.MeasureString(test, f).Width <= maxWidth Then
                    found = True
                    Exit While
                End If
                cut -= 1
            End While

            If Not found Then
                cut = Math.Max(1, Math.Min(remaining.Length, 10))
            End If

            Dim breakPos As Integer = cut
            Dim sp As Integer = remaining.LastIndexOf(" "c, Math.Max(0, cut - 1))
            If sp > 0 Then breakPos = sp

            Dim linePart As String = remaining.Substring(0, breakPos).TrimEnd()
            If linePart = "" Then
                linePart = remaining.Substring(0, Math.Min(cut, remaining.Length))
                breakPos = linePart.Length
            End If

            result.Add(linePart)

            If breakPos >= remaining.Length Then
                remaining = ""
            Else
                remaining = remaining.Substring(breakPos).TrimStart()
                If remaining.Length > 0 AndAlso indentCount > 0 Then
                    remaining = indent & remaining
                End If
            End If
        End While

        If result.Count = 0 Then result.Add("")
        Return result
    End Function

    Private Sub EndPrintCleanup(sender As Object, e As PrintEventArgs)
        Try
            _printLineIndex = 0
            If _printLines Is Nothing Then
                _printLines = New List(Of String)()
            Else
                _printLines.Clear()
            End If

            If _printBodyFont IsNot Nothing Then
                _printBodyFont.Dispose()
                _printBodyFont = Nothing
            End If
            If _printHeaderFont IsNot Nothing Then
                _printHeaderFont.Dispose()
                _printHeaderFont = Nothing
            End If
            If _printSubFont IsNot Nothing Then
                _printSubFont.Dispose()
                _printSubFont = Nothing
            End If

            Dim doc As PrintDocument = TryCast(sender, PrintDocument)
            If doc IsNot Nothing Then
                RemoveHandler doc.PrintPage, AddressOf PrintPneumaticReportPage
                RemoveHandler doc.EndPrint, AddressOf EndPrintCleanup
            End If
        Catch
        End Try
    End Sub

    ' ============================================================
    ' RESET
    ' ============================================================
    Private Sub ResetAll(sender As Object, e As EventArgs)
        SyncThemeFromCurrentUI()
        ClearOutputs()

        txtJobName.Text = "Cylinder-01"
        ClearTB(txtBore) : txtBore.Text = "20"
        ClearTB(txtStroke) : txtStroke.Text = "100"
        ClearTB(txtQty) : txtQty.Text = "1"
        ClearTB(txtPressure) : txtPressure.Text = "6"
        ClearTB(txtTubeLen) : txtTubeLen.Text = "2"
        ClearTB(txtCyclesPerMin) : txtCyclesPerMin.Text = "5"

        SetComboSafe(cmbProjectMode, 0)
        SetComboSafe(cmbCylinderType, 0)
        SetComboSafe(cmbOrientation, 0)
        SetComboSafe(cmbMotionType, 0)
        SetComboSafe(cmbOperationMode, 2)
        SetComboSafe(cmbFailSafe, 0)
        SetComboSafe(cmbSpeedPriority, 0)
        SetComboSafe(cmbCushioning, 1)
        SetComboSafe(cmbVendor, 0)
        SetComboSafe(cmbVoltage, 0)
        SetComboSafe(cmbAirQuality, 0)
        SetComboSafe(cmbEnvironment, 0)
        SetComboSafe(cmbPreviewMode, 0)

        chkFRL.Checked = True
        chkBranchReg.Checked = False
        chkQuickExhaust.Checked = False
        chkSensors.Checked = True
        chkManualOverride.Checked = True
        chkDumpValve.Checked = False

        jobs.Clear()
        currentJobIndex = -1
        RefreshJobsList()

        lastCalcOk = False
        lastLineTags = New List(Of String)()
        lastBOMLines = New List(Of String)()

        RenderPreviewPlaceholder("Reset complete. Click Calculate, then Preview.")
        SetStatusWarn("Status: Reset (inputs restored)")
    End Sub

    ' ============================================================
    ' OUTPUT HELPERS
    ' ============================================================
    Private Sub ClearOutputs()
        If txtFinal IsNot Nothing Then txtFinal.Clear()
        If txtCalc IsNot Nothing Then txtCalc.Clear()
        If txtBom IsNot Nothing Then txtBom.Clear()
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
            rtbWarn.SelectionFont = New Font(rtbWarn.Font, If(bold, FontStyle.Bold, FontStyle.Regular))
            rtbWarn.SelectionColor = col
            rtbWarn.AppendText(text & vbCrLf)
            rtbWarn.SelectionColor = themeText
            rtbWarn.SelectionFont = New Font(rtbWarn.Font, FontStyle.Regular)
        Catch
        End Try
    End Sub

    Private Sub SetStatusOk(text As String)
        If lblStatus Is Nothing Then Exit Sub
        lblStatus.Text = text
        lblStatus.ForeColor = themeOk
    End Sub

    Private Sub SetStatusWarn(text As String)
        If lblStatus Is Nothing Then Exit Sub
        lblStatus.Text = text
        lblStatus.ForeColor = themeWarn
    End Sub

    Private Sub SetStatusError(text As String)
        If lblStatus Is Nothing Then Exit Sub
        lblStatus.Text = text
        lblStatus.ForeColor = themeErr
    End Sub

    ' ============================================================
    ' GENERIC HELPERS
    ' ============================================================
    Private Function ReadD(tb As TextBox, defaultValue As Double) As Double
        If tb Is Nothing Then Return defaultValue
        Dim s As String = tb.Text
        If s Is Nothing Then Return defaultValue
        s = s.Trim()
        If s = "" Then Return defaultValue

        Dim v As Double
        s = s.Replace(","c, "."c)
        If Double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, v) Then
            Return v
        End If
        Return defaultValue
    End Function

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

    Private Sub SelectComboByText(cb As ComboBox, textValue As String)
        If cb Is Nothing Then Exit Sub
        If textValue Is Nothing Then Exit Sub
        For i As Integer = 0 To cb.Items.Count - 1
            Dim s As String = Convert.ToString(cb.Items(i), CultureInfo.InvariantCulture)
            If String.Equals(s, textValue, StringComparison.OrdinalIgnoreCase) Then
                cb.SelectedIndex = i
                Exit Sub
            End If
        Next
        If cb.Items.Count > 0 AndAlso cb.SelectedIndex < 0 Then cb.SelectedIndex = 0
    End Sub

    Private Function SafeText(tb As TextBox, fallbackValue As String) As String
        If tb Is Nothing Then Return fallbackValue
        Dim s As String = If(tb.Text, "").Trim()
        If s = "" Then Return fallbackValue
        Return s
    End Function

    Private Function SafeComboText(cb As ComboBox) As String
        If cb Is Nothing Then Return ""
        If cb.Text Is Nothing Then Return ""
        Return cb.Text.Trim()
    End Function

    Private Function SafeInt(ByVal s As String, ByVal defaultValue As Integer) As Integer
        If s Is Nothing Then Return defaultValue
        Dim v As Integer
        If Integer.TryParse(s.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, v) Then
            Return v
        End If
        Dim d As Double
        If Double.TryParse(s.Trim().Replace(","c, "."c), NumberStyles.Any, CultureInfo.InvariantCulture, d) Then
            Return CInt(Math.Round(d))
        End If
        Return defaultValue
    End Function

End Class

' ============================================================
' SIMPLE ASCII DXF WRITER (R12-ish entities)
' ============================================================
Friend Class SimpleDxfWriter
    Private ReadOnly sb As StringBuilder

    Public Sub New()
        sb = New StringBuilder()
    End Sub

    Public Sub StartFile()
        sb.Length = 0
        AppendPair(0, "SECTION")
        AppendPair(2, "HEADER")
        AppendPair(9, "$ACADVER")
        AppendPair(1, "AC1009")
        AppendPair(0, "ENDSEC")

        AppendPair(0, "SECTION")
        AppendPair(2, "TABLES")
        WriteLinetypeTable()
        WriteLayerTable()
        AppendPair(0, "ENDSEC")

        AppendPair(0, "SECTION")
        AppendPair(2, "ENTITIES")
    End Sub

    Public Sub EndFile()
        AppendPair(0, "ENDSEC")
        AppendPair(0, "EOF")
    End Sub

    Private Sub WriteLinetypeTable()
        AppendPair(0, "TABLE")
        AppendPair(2, "LTYPE")
        AppendPair(70, "1")
        AppendPair(0, "LTYPE")
        AppendPair(2, "CONTINUOUS")
        AppendPair(70, "0")
        AppendPair(3, "Solid line")
        AppendPair(72, "65")
        AppendPair(73, "0")
        AppendPair(40, "0.0")
        AppendPair(0, "ENDTAB")
    End Sub

    Private Sub WriteLayerTable()
        AppendPair(0, "TABLE")
        AppendPair(2, "LAYER")
        AppendPair(70, "6")

        WriteLayer("0", 7)
        WriteLayer("AIR_G", 3)
        WriteLayer("SYM_W", 7)
        WriteLayer("TEXT_Y", 2)
        WriteLayer("TEXT_R", 1)
        WriteLayer("BORDER", 7)

        AppendPair(0, "ENDTAB")
    End Sub

    Private Sub WriteLayer(ByVal name As String, ByVal colorIndex As Integer)
        AppendPair(0, "LAYER")
        AppendPair(2, name)
        AppendPair(70, "0")
        AppendPair(62, colorIndex.ToString(CultureInfo.InvariantCulture))
        AppendPair(6, "CONTINUOUS")
    End Sub

    Public Sub Line(x1 As Double, y1 As Double, x2 As Double, y2 As Double, Optional layerName As String = "0")
        AppendPair(0, "LINE")
        AppendPair(8, layerName)
        AppendPair(10, F(x1))
        AppendPair(20, F(y1))
        AppendPair(11, F(x2))
        AppendPair(21, F(y2))
    End Sub

    Public Sub Circle(x As Double, y As Double, r As Double, Optional layerName As String = "0")
        AppendPair(0, "CIRCLE")
        AppendPair(8, layerName)
        AppendPair(10, F(x))
        AppendPair(20, F(y))
        AppendPair(40, F(r))
    End Sub

    Public Sub Text(x As Double, y As Double, txt As String, h As Double, Optional layerName As String = "0")
        AppendPair(0, "TEXT")
        AppendPair(8, layerName)
        AppendPair(10, F(x))
        AppendPair(20, F(y))
        AppendPair(40, F(h))
        AppendPair(1, txt)
    End Sub

    Private Function F(v As Double) As String
        Return v.ToString("0.###", CultureInfo.InvariantCulture)
    End Function

    Private Sub AppendPair(code As Integer, value As String)
        sb.AppendLine(code.ToString(CultureInfo.InvariantCulture))
        sb.AppendLine(value)
    End Sub

    Public Overrides Function ToString() As String
        Return sb.ToString()
    End Function
End Class