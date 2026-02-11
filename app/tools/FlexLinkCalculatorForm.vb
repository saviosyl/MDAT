Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Collections.Generic

'============================================================
' FlexLinkCalculatorForm
' - Base calculator form intended to be inherited by wrapper:
'   FlexLinkProjectConfiguratorForm : FlexLinkCalculatorForm
'============================================================
Public Class FlexLinkCalculatorForm
    Inherits Form

    '============================================================
    ' Models
    '============================================================
    Private Enum SegmentType
        Straight = 0
        Bend = 1
    End Enum

    Private Enum FamilyType
        Beam = 0
        Chain = 1
        Slide = 2
    End Enum

    Private Class Segment
        Public SegType As SegmentType
        Public Family As FamilyType
        Public Quantity As Integer

        ' Straight
        Public LengthMm As Double

        ' Bend
        Public BendAngleDeg As Double
        Public BendRadiusMm As Double

        Public Function EffectiveLengthMm() As Double
            If SegType = SegmentType.Straight Then
                Return LengthMm
            Else
                Dim thetaRad As Double = (Math.PI / 180.0R) * BendAngleDeg
                Return BendRadiusMm * thetaRad
            End If
        End Function
    End Class

    Private ReadOnly _segments As New List(Of Segment)()

    '============================================================
    ' UI Controls
    '============================================================
    Private headerPanel As Panel
    Private contentSplit As SplitContainer

    ' Left
    Private leftLayout As TableLayoutPanel

    Private grpSegmentEntry As GroupBox
    Private rbStraight As RadioButton
    Private rbBend As RadioButton
    Private cboFamily As ComboBox
    Private nudQty As NumericUpDown

    Private grpStraight As GroupBox
    Private txtStraightLenMm As TextBox

    Private grpBend As GroupBox
    Private cboBendAngle As ComboBox
    Private txtAngleCustom As TextBox
    Private cboBendRadius As ComboBox
    Private txtRadiusCustom As TextBox

    Private btnAddSegment As Button
    Private btnRemoveSelected As Button
    Private btnClearSegments As Button

    Private grpSegments As GroupBox
    Private dgvSegments As DataGridView

    Private grpSupports As GroupBox
    Private chkIncludeLegSupports As CheckBox
    Private txtLegSupportPitchMm As TextBox

    ' Right
    Private rightLayout As TableLayoutPanel

    Private grpTotals As GroupBox
    Private lvTotals As ListView

    Private grpBeamPack As GroupBox
    Private grpChainPack As GroupBox
    Private grpSlidePack As GroupBox

    Private nudBeamPackLenM As NumericUpDown
    Private nudChainPackLenM As NumericUpDown
    Private nudSlidePackLenM As NumericUpDown

    Private lblBeamSummary As Label
    Private lblChainSummary As Label
    Private lblSlideSummary As Label

    Private btnPrintPreview As Button
    Private btnExportPdf As Button

    ' Printing
    Private printDoc As PrintDocument
    Private printPreview As PrintPreviewDialog

    ' Totals
    Private Structure FamilyTotals
        Public TotalLengthMm As Double
        Public TotalBends As Integer
    End Structure

    Private _totBeam As FamilyTotals
    Private _totChain As FamilyTotals
    Private _totSlide As FamilyTotals

    '============================================================
    ' Constructor
    '============================================================
    Public Sub New()
        MyBase.New()

        Me.Text = "FlexLink Project Configurator"
        Me.StartPosition = FormStartPosition.CenterScreen

        Me.AutoScaleMode = AutoScaleMode.Dpi
        Me.WindowState = FormWindowState.Maximized
        Me.MinimumSize = New Size(1200, 750)

        BuildUi()
        WireEvents()

        ' Defaults
        rbStraight.Checked = True
        cboFamily.SelectedIndex = 0
        nudQty.Value = 1

        cboBendAngle.SelectedIndex = 3 ' 90
        cboBendRadius.SelectedIndex = 1 ' 300

        txtStraightLenMm.Text = "1000"
        txtAngleCustom.Text = "90"
        txtRadiusCustom.Text = "300"

        nudBeamPackLenM.Value = 3D
        nudChainPackLenM.Value = 5D
        nudSlidePackLenM.Value = 5D

        chkIncludeLegSupports.Checked = True
        txtLegSupportPitchMm.Text = "1000"

        RecalcAll()
    End Sub

    '============================================================
    ' UI Build
    '============================================================
    Private Sub BuildUi()
        Me.SuspendLayout()

        ' ===== Header =====
        headerPanel = New Panel()
        headerPanel.Dock = DockStyle.Top
        headerPanel.Height = 58
        headerPanel.Padding = New Padding(16, 10, 16, 10)
        headerPanel.BackColor = Color.FromArgb(16, 28, 51)

        Dim lblTitle As New Label()
        lblTitle.AutoSize = True
        lblTitle.ForeColor = Color.White
        lblTitle.Font = New Font("Segoe UI", 14.0F, FontStyle.Bold, GraphicsUnit.Point)
        lblTitle.Text = "FlexLink Project Configurator — Segments • Overall BOM Totals"
        lblTitle.Location = New Point(10, 13)

        headerPanel.Controls.Add(lblTitle)
        Me.Controls.Add(headerPanel)

        ' ===== Split =====
        contentSplit = New SplitContainer()
        contentSplit.Dock = DockStyle.Fill
        contentSplit.Orientation = Orientation.Vertical
        contentSplit.SplitterWidth = 8
        contentSplit.SplitterDistance = CInt(Me.ClientSize.Width * 0.62)

        Me.Controls.Add(contentSplit)

        ' =========================
        ' LEFT LAYOUT (3 rows)
        ' =========================
        leftLayout = New TableLayoutPanel()
        leftLayout.Dock = DockStyle.Fill
        leftLayout.ColumnCount = 1
        leftLayout.RowCount = 3
        leftLayout.Padding = New Padding(12)
        leftLayout.BackColor = Color.FromArgb(11, 18, 34)
        leftLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 280.0F)) ' segment entry
        leftLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))  ' segments grid
        leftLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 140.0F)) ' supports
        contentSplit.Panel1.Controls.Add(leftLayout)

        ' =========================
        ' RIGHT LAYOUT (3 rows)
        ' =========================
        rightLayout = New TableLayoutPanel()
        rightLayout.Dock = DockStyle.Fill
        rightLayout.ColumnCount = 1
        rightLayout.RowCount = 3
        rightLayout.Padding = New Padding(12)
        rightLayout.BackColor = Color.FromArgb(11, 18, 34)
        rightLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))  ' totals
        rightLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 420.0F)) ' packaging (3 boxes)
        rightLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 70.0F))  ' buttons
        contentSplit.Panel2.Controls.Add(rightLayout)

        '========================================================
        ' 1) Segment Entry (usable layout)
        '========================================================
        grpSegmentEntry = MakeGroupBox("1) Segment Entry", True)
        grpSegmentEntry.Dock = DockStyle.Fill
        leftLayout.Controls.Add(grpSegmentEntry, 0, 0)

        Dim segLayout As TableLayoutPanel = New TableLayoutPanel()
        segLayout.Dock = DockStyle.Fill
        segLayout.ColumnCount = 4
        segLayout.RowCount = 4
        segLayout.Padding = New Padding(10)
        segLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
        segLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
        segLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
        segLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
        segLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 34.0F)) ' straight/bend
        segLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F)) ' family/qty
        segLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 110.0F)) ' straight + bend boxes
        segLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 52.0F)) ' buttons
        grpSegmentEntry.Controls.Add(segLayout)

        ' Row 0: straight / bend
        rbStraight = New RadioButton()
        rbStraight.Text = "Straight"
        rbStraight.ForeColor = Color.White
        rbStraight.AutoSize = True
        rbStraight.Anchor = AnchorStyles.Left
        rbBend = New RadioButton()
        rbBend.Text = "Bend"
        rbBend.ForeColor = Color.White
        rbBend.AutoSize = True
        rbBend.Anchor = AnchorStyles.Left

        segLayout.Controls.Add(rbStraight, 0, 0)
        segLayout.Controls.Add(rbBend, 1, 0)

        ' Row 1: family + qty
        Dim lblFamily As New Label()
        lblFamily.Text = "Family:"
        lblFamily.ForeColor = Color.White
        lblFamily.AutoSize = True
        lblFamily.Anchor = AnchorStyles.Left

        cboFamily = New ComboBox()
        cboFamily.DropDownStyle = ComboBoxStyle.DropDownList
        cboFamily.Items.Add("Beam")
        cboFamily.Items.Add("Chain")
        cboFamily.Items.Add("Slide")
        cboFamily.Dock = DockStyle.Fill

        Dim lblQty As New Label()
        lblQty.Text = "Qty:"
        lblQty.ForeColor = Color.White
        lblQty.AutoSize = True
        lblQty.Anchor = AnchorStyles.Left

        nudQty = New NumericUpDown()
        nudQty.Minimum = 1
        nudQty.Maximum = 9999
        nudQty.Dock = DockStyle.Left
        nudQty.Width = 120

        segLayout.Controls.Add(lblFamily, 0, 1)
        segLayout.Controls.Add(cboFamily, 1, 1)
        segLayout.Controls.Add(lblQty, 2, 1)
        segLayout.Controls.Add(nudQty, 3, 1)

        ' Row 2: Straight + Bend GroupBoxes (2 columns each)
        grpStraight = MakeGroupBox("Straight", True)
        grpStraight.Dock = DockStyle.Fill
        segLayout.SetColumnSpan(grpStraight, 2)
        segLayout.Controls.Add(grpStraight, 0, 2)

        grpBend = MakeGroupBox("Bend", True)
        grpBend.Dock = DockStyle.Fill
        segLayout.SetColumnSpan(grpBend, 2)
        segLayout.Controls.Add(grpBend, 2, 2)

        ' Straight inner layout
        Dim straightLay As TableLayoutPanel = New TableLayoutPanel()
        straightLay.Dock = DockStyle.Fill
        straightLay.ColumnCount = 2
        straightLay.RowCount = 1
        straightLay.Padding = New Padding(10, 8, 10, 8)
        straightLay.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110.0F))
        straightLay.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        grpStraight.Controls.Add(straightLay)

        Dim lblLen As New Label()
        lblLen.Text = "Length (mm):"
        lblLen.ForeColor = Color.White
        lblLen.AutoSize = True
        lblLen.Anchor = AnchorStyles.Left

        txtStraightLenMm = New TextBox()
        txtStraightLenMm.Dock = DockStyle.Left
        txtStraightLenMm.Width = 140

        straightLay.Controls.Add(lblLen, 0, 0)
        straightLay.Controls.Add(txtStraightLenMm, 1, 0)

        ' Bend inner layout (no clipping)
        Dim bendLay As TableLayoutPanel = New TableLayoutPanel()
        bendLay.Dock = DockStyle.Fill
        bendLay.ColumnCount = 4
        bendLay.RowCount = 2
        bendLay.Padding = New Padding(10, 8, 10, 8)
        bendLay.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 95.0F))   ' Angle label
        bendLay.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120.0F)) ' Angle combo
        bendLay.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 95.0F))  ' Radius label
        bendLay.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))  ' Radius combo/custom
        bendLay.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0F))
        bendLay.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0F))
        grpBend.Controls.Add(bendLay)

        Dim lblAngle As New Label()
        lblAngle.Text = "Angle (deg):"
        lblAngle.ForeColor = Color.White
        lblAngle.AutoSize = True
        lblAngle.Anchor = AnchorStyles.Left

        cboBendAngle = New ComboBox()
        cboBendAngle.DropDownStyle = ComboBoxStyle.DropDownList
        cboBendAngle.Items.Add("30")
        cboBendAngle.Items.Add("45")
        cboBendAngle.Items.Add("60")
        cboBendAngle.Items.Add("90")
        cboBendAngle.Items.Add("Custom...")
        cboBendAngle.Dock = DockStyle.Fill

        txtAngleCustom = New TextBox()
        txtAngleCustom.Dock = DockStyle.Left
        txtAngleCustom.Width = 100
        txtAngleCustom.Enabled = False

        Dim lblRadius As New Label()
        lblRadius.Text = "Radius (mm):"
        lblRadius.ForeColor = Color.White
        lblRadius.AutoSize = True
        lblRadius.Anchor = AnchorStyles.Left

        cboBendRadius = New ComboBox()
        cboBendRadius.DropDownStyle = ComboBoxStyle.DropDownList
        cboBendRadius.Items.Add("150")
        cboBendRadius.Items.Add("300")
        cboBendRadius.Items.Add("500")
        cboBendRadius.Items.Add("Custom...")
        cboBendRadius.Dock = DockStyle.Left
        cboBendRadius.Width = 140

        txtRadiusCustom = New TextBox()
        txtRadiusCustom.Dock = DockStyle.Left
        txtRadiusCustom.Width = 140
        txtRadiusCustom.Enabled = False

        ' Row 0: Angle label + combo, Radius label + combo
        bendLay.Controls.Add(lblAngle, 0, 0)
        bendLay.Controls.Add(cboBendAngle, 1, 0)
        bendLay.Controls.Add(lblRadius, 2, 0)
        bendLay.Controls.Add(cboBendRadius, 3, 0)

        ' Row 1: custom inputs aligned under their combos
        bendLay.Controls.Add(New Label(), 0, 1)
        bendLay.Controls.Add(txtAngleCustom, 1, 1)
        bendLay.Controls.Add(New Label(), 2, 1)
        bendLay.Controls.Add(txtRadiusCustom, 3, 1)

        ' Row 3: Buttons
        Dim btnLay As FlowLayoutPanel = New FlowLayoutPanel()
        btnLay.Dock = DockStyle.Fill
        btnLay.FlowDirection = FlowDirection.LeftToRight
        btnLay.WrapContents = False
        btnLay.Padding = New Padding(0, 6, 0, 0)

        btnAddSegment = New Button()
        btnAddSegment.Text = "Add Segment"
        btnAddSegment.Width = 140
        btnAddSegment.Height = 34

        btnRemoveSelected = New Button()
        btnRemoveSelected.Text = "Remove Selected"
        btnRemoveSelected.Width = 150
        btnRemoveSelected.Height = 34

        btnClearSegments = New Button()
        btnClearSegments.Text = "Clear"
        btnClearSegments.Width = 90
        btnClearSegments.Height = 34

        btnLay.Controls.Add(btnAddSegment)
        btnLay.Controls.Add(btnRemoveSelected)
        btnLay.Controls.Add(btnClearSegments)

        segLayout.SetColumnSpan(btnLay, 4)
        segLayout.Controls.Add(btnLay, 0, 3)

        '========================================================
        ' 2) Segments grid
        '========================================================
        grpSegments = MakeGroupBox("2) Segments (for calculation)", True)
        grpSegments.Dock = DockStyle.Fill
        leftLayout.Controls.Add(grpSegments, 0, 1)

        dgvSegments = New DataGridView()
        dgvSegments.Dock = DockStyle.Fill
        dgvSegments.AllowUserToAddRows = False
        dgvSegments.AllowUserToDeleteRows = False
        dgvSegments.ReadOnly = True
        dgvSegments.RowHeadersVisible = False
        dgvSegments.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvSegments.MultiSelect = False
        dgvSegments.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dgvSegments.BackgroundColor = Color.FromArgb(16, 28, 51)
        dgvSegments.BorderStyle = BorderStyle.None
        dgvSegments.EnableHeadersVisualStyles = False
        dgvSegments.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(20, 40, 70)
        dgvSegments.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvSegments.DefaultCellStyle.BackColor = Color.FromArgb(16, 28, 51)
        dgvSegments.DefaultCellStyle.ForeColor = Color.White
        dgvSegments.DefaultCellStyle.SelectionBackColor = Color.FromArgb(22, 179, 184)
        dgvSegments.DefaultCellStyle.SelectionForeColor = Color.Black

        dgvSegments.Columns.Add("colType", "Type")
        dgvSegments.Columns.Add("colFamily", "Family")
        dgvSegments.Columns.Add("colQty", "Qty")
        dgvSegments.Columns.Add("colDetail", "Details")
        dgvSegments.Columns.Add("colEffLen", "Effective Length (mm)")

        grpSegments.Controls.Add(dgvSegments)

        '========================================================
        ' 3) Supports
        '========================================================
        grpSupports = MakeGroupBox("3) Leg Supports", True)
        grpSupports.Dock = DockStyle.Fill
        leftLayout.Controls.Add(grpSupports, 0, 2)

        Dim supLay As TableLayoutPanel = New TableLayoutPanel()
        supLay.Dock = DockStyle.Fill
        supLay.ColumnCount = 2
        supLay.RowCount = 2
        supLay.Padding = New Padding(10)
        supLay.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 170.0F))
        supLay.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        supLay.RowStyles.Add(New RowStyle(SizeType.Absolute, 34.0F))
        supLay.RowStyles.Add(New RowStyle(SizeType.Absolute, 34.0F))
        grpSupports.Controls.Add(supLay)

        chkIncludeLegSupports = New CheckBox()
        chkIncludeLegSupports.Text = "Include leg supports in totals"
        chkIncludeLegSupports.ForeColor = Color.White
        chkIncludeLegSupports.AutoSize = True
        chkIncludeLegSupports.Anchor = AnchorStyles.Left

        Dim lblPitch As New Label()
        lblPitch.Text = "Leg support pitch (mm):"
        lblPitch.ForeColor = Color.White
        lblPitch.AutoSize = True
        lblPitch.Anchor = AnchorStyles.Left

        ' Height fix: Multiline True + fixed Size
        txtLegSupportPitchMm = New TextBox()
        txtLegSupportPitchMm.Multiline = True
        txtLegSupportPitchMm.Size = New Size(160, 26)
        txtLegSupportPitchMm.Anchor = AnchorStyles.Left

        supLay.SetColumnSpan(chkIncludeLegSupports, 2)
        supLay.Controls.Add(chkIncludeLegSupports, 0, 0)
        supLay.Controls.Add(lblPitch, 0, 1)
        supLay.Controls.Add(txtLegSupportPitchMm, 1, 1)

        '========================================================
        ' RIGHT: Totals
        '========================================================
        grpTotals = MakeGroupBox("Overall Project BOM Totals (Only)", True)
        grpTotals.Dock = DockStyle.Fill
        rightLayout.Controls.Add(grpTotals, 0, 0)

        lvTotals = New ListView()
        lvTotals.Dock = DockStyle.Fill
        lvTotals.View = View.Details
        lvTotals.FullRowSelect = True
        lvTotals.GridLines = True
        lvTotals.HideSelection = False

        lvTotals.Columns.Add("Family", 100)
        lvTotals.Columns.Add("Item", 200)
        lvTotals.Columns.Add("Qty", 70)
        lvTotals.Columns.Add("Unit", 70)
        lvTotals.Columns.Add("Notes", 260)

        grpTotals.Controls.Add(lvTotals)

        '========================================================
        ' RIGHT: Packaging (3 visible boxes)
        '========================================================
        Dim packLayout As TableLayoutPanel = New TableLayoutPanel()
        packLayout.Dock = DockStyle.Fill
        packLayout.ColumnCount = 1
        packLayout.RowCount = 3
        packLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 33.33F))
        packLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 33.33F))
        packLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 33.33F))
        rightLayout.Controls.Add(packLayout, 0, 1)

        grpBeamPack = MakeGroupBox("Beam Packaging", True)
        grpChainPack = MakeGroupBox("Chain Packaging", True)
        grpSlidePack = MakeGroupBox("Slide Packaging", True)

        grpBeamPack.Dock = DockStyle.Fill
        grpChainPack.Dock = DockStyle.Fill
        grpSlidePack.Dock = DockStyle.Fill

        packLayout.Controls.Add(grpBeamPack, 0, 0)
        packLayout.Controls.Add(grpChainPack, 0, 1)
        packLayout.Controls.Add(grpSlidePack, 0, 2)

        BuildPackBox(grpBeamPack, nudBeamPackLenM, lblBeamSummary)
        BuildPackBox(grpChainPack, nudChainPackLenM, lblChainSummary)
        BuildPackBox(grpSlidePack, nudSlidePackLenM, lblSlideSummary)

        '========================================================
        ' RIGHT: Buttons bottom
        '========================================================
        Dim btnBottom As FlowLayoutPanel = New FlowLayoutPanel()
        btnBottom.Dock = DockStyle.Fill
        btnBottom.FlowDirection = FlowDirection.RightToLeft
        btnBottom.WrapContents = False
        btnBottom.Padding = New Padding(0, 12, 0, 0)

        btnExportPdf = New Button()
        btnExportPdf.Text = "Export PDF (Print)"
        btnExportPdf.Width = 170
        btnExportPdf.Height = 36

        btnPrintPreview = New Button()
        btnPrintPreview.Text = "Preview"
        btnPrintPreview.Width = 120
        btnPrintPreview.Height = 36

        btnBottom.Controls.Add(btnExportPdf)
        btnBottom.Controls.Add(btnPrintPreview)

        rightLayout.Controls.Add(btnBottom, 0, 2)

        ' Printing
        printDoc = New PrintDocument()
        printPreview = New PrintPreviewDialog()
        printPreview.Document = printDoc
        printPreview.Width = 1100
        printPreview.Height = 800

        Me.ResumeLayout()
    End Sub

    Private Sub BuildPackBox(ByVal grp As GroupBox, ByRef nud As NumericUpDown, ByRef lblSummary As Label)
        Dim lay As TableLayoutPanel = New TableLayoutPanel()
        lay.Dock = DockStyle.Fill
        lay.ColumnCount = 2
        lay.RowCount = 2
        lay.Padding = New Padding(10)
        lay.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 160.0F))
        lay.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        lay.RowStyles.Add(New RowStyle(SizeType.Absolute, 34.0F))
        lay.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        grp.Controls.Add(lay)

        Dim lbl As New Label()
        lbl.Text = "Pack length (m):"
        lbl.ForeColor = Color.White
        lbl.AutoSize = True
        lbl.Anchor = AnchorStyles.Left

        nud = New NumericUpDown()
        nud.DecimalPlaces = 1
        nud.Minimum = 0.1D
        nud.Maximum = 9999D
        nud.Increment = 0.5D
        nud.Width = 140
        nud.Anchor = AnchorStyles.Left

        lblSummary = New Label()
        lblSummary.ForeColor = Color.White
        lblSummary.Dock = DockStyle.Fill
        lblSummary.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)
        lblSummary.Text = "-"

        lay.Controls.Add(lbl, 0, 0)
        lay.Controls.Add(nud, 1, 0)
        lay.SetColumnSpan(lblSummary, 2)
        lay.Controls.Add(lblSummary, 0, 1)
    End Sub

    Private Function MakeGroupBox(ByVal title As String, ByVal bold As Boolean) As GroupBox
        Dim g As New GroupBox()
        g.Text = title
        g.ForeColor = Color.White
        g.BackColor = Color.FromArgb(16, 28, 51)
        g.Padding = New Padding(10)
        If bold Then
            g.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold, GraphicsUnit.Point)
        Else
            g.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)
        End If
        Return g
    End Function

    '============================================================
    ' Events
    '============================================================
    Private Sub WireEvents()
        AddHandler rbStraight.CheckedChanged, AddressOf OnSegTypeChanged
        AddHandler rbBend.CheckedChanged, AddressOf OnSegTypeChanged

        AddHandler cboBendAngle.SelectedIndexChanged, AddressOf OnAngleModeChanged
        AddHandler cboBendRadius.SelectedIndexChanged, AddressOf OnRadiusModeChanged

        AddHandler btnAddSegment.Click, AddressOf OnAddSegment
        AddHandler btnRemoveSelected.Click, AddressOf OnRemoveSelected
        AddHandler btnClearSegments.Click, AddressOf OnClearSegments

        AddHandler chkIncludeLegSupports.CheckedChanged, AddressOf OnRecalcAny
        AddHandler txtLegSupportPitchMm.TextChanged, AddressOf OnRecalcAny

        AddHandler nudBeamPackLenM.ValueChanged, AddressOf OnRecalcAny
        AddHandler nudChainPackLenM.ValueChanged, AddressOf OnRecalcAny
        AddHandler nudSlidePackLenM.ValueChanged, AddressOf OnRecalcAny

        AddHandler btnExportPdf.Click, AddressOf OnExportPdf
        AddHandler btnPrintPreview.Click, AddressOf OnPreview

        AddHandler printDoc.PrintPage, AddressOf OnPrintPage
    End Sub

    Private Sub OnRecalcAny(ByVal sender As Object, ByVal e As EventArgs)
        RecalcAll()
    End Sub

    Private Sub OnSegTypeChanged(ByVal sender As Object, ByVal e As EventArgs)
        grpStraight.Enabled = rbStraight.Checked
        grpBend.Enabled = rbBend.Checked
    End Sub

    Private Sub OnAngleModeChanged(ByVal sender As Object, ByVal e As EventArgs)
        If cboBendAngle.SelectedItem Is Nothing Then Exit Sub
        Dim s As String = cboBendAngle.SelectedItem.ToString()
        If String.Equals(s, "Custom...", StringComparison.OrdinalIgnoreCase) Then
            txtAngleCustom.Enabled = True
            txtAngleCustom.Focus()
        Else
            txtAngleCustom.Enabled = False
            txtAngleCustom.Text = s
        End If
    End Sub

    Private Sub OnRadiusModeChanged(ByVal sender As Object, ByVal e As EventArgs)
        If cboBendRadius.SelectedItem Is Nothing Then Exit Sub
        Dim s As String = cboBendRadius.SelectedItem.ToString()
        If String.Equals(s, "Custom...", StringComparison.OrdinalIgnoreCase) Then
            txtRadiusCustom.Enabled = True
            txtRadiusCustom.Focus()
        Else
            txtRadiusCustom.Enabled = False
            txtRadiusCustom.Text = s
        End If
    End Sub

    Private Sub OnAddSegment(ByVal sender As Object, ByVal e As EventArgs)
        Dim seg As New Segment()
        seg.Quantity = CInt(nudQty.Value)
        seg.Family = CType(cboFamily.SelectedIndex, FamilyType)

        If rbStraight.Checked Then
            seg.SegType = SegmentType.Straight
            Dim lenMm As Double
            If Not TryParseDouble(txtStraightLenMm.Text, lenMm) OrElse lenMm <= 0 Then
                MessageBox.Show("Enter a valid straight length in mm.", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            seg.LengthMm = lenMm
        Else
            seg.SegType = SegmentType.Bend
            Dim angleDeg As Double
            Dim radiusMm As Double

            If Not TryParseDouble(txtAngleCustom.Text, angleDeg) OrElse angleDeg <= 0 OrElse angleDeg > 360 Then
                MessageBox.Show("Enter a valid bend angle (1–360).", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            If Not TryParseDouble(txtRadiusCustom.Text, radiusMm) OrElse radiusMm <= 0 Then
                MessageBox.Show("Enter a valid bend radius in mm.", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            seg.BendAngleDeg = angleDeg
            seg.BendRadiusMm = radiusMm
        End If

        _segments.Add(seg)
        RefreshSegmentsGrid()
        RecalcAll()
    End Sub

    Private Sub OnRemoveSelected(ByVal sender As Object, ByVal e As EventArgs)
        If dgvSegments.SelectedRows Is Nothing OrElse dgvSegments.SelectedRows.Count = 0 Then Exit Sub
        Dim idx As Integer = dgvSegments.SelectedRows(0).Index
        If idx >= 0 AndAlso idx < _segments.Count Then
            _segments.RemoveAt(idx)
            RefreshSegmentsGrid()
            RecalcAll()
        End If
    End Sub

    Private Sub OnClearSegments(ByVal sender As Object, ByVal e As EventArgs)
        If _segments.Count = 0 Then Exit Sub
        If MessageBox.Show("Clear all segments?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            _segments.Clear()
            RefreshSegmentsGrid()
            RecalcAll()
        End If
    End Sub

    Private Sub OnPreview(ByVal sender As Object, ByVal e As EventArgs)
        Try
            printPreview.ShowDialog(Me)
        Catch ex As Exception
            MessageBox.Show("Preview failed: " & ex.Message, "Print Preview", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub OnExportPdf(ByVal sender As Object, ByVal e As EventArgs)
        Dim dlg As New PrintDialog()
        dlg.AllowSomePages = False
        dlg.AllowSelection = False
        dlg.UseEXDialog = True
        dlg.Document = printDoc

        If dlg.ShowDialog(Me) = DialogResult.OK Then
            Try
                printDoc.Print()
            Catch ex As Exception
                MessageBox.Show("Print/Export failed: " & ex.Message, "Export PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    '============================================================
    ' Grid + Totals
    '============================================================
    Private Sub RefreshSegmentsGrid()
        dgvSegments.Rows.Clear()

        For Each s As Segment In _segments
            Dim typeText As String = If(s.SegType = SegmentType.Straight, "Straight", "Bend")
            Dim famText As String = FamilyToText(s.Family)
            Dim detail As String

            If s.SegType = SegmentType.Straight Then
                detail = "L=" & FormatNumberNoTrail(s.LengthMm) & " mm"
            Else
                detail = "A=" & FormatNumberNoTrail(s.BendAngleDeg) & "°, R=" & FormatNumberNoTrail(s.BendRadiusMm) & " mm"
            End If

            Dim eff As Double = s.EffectiveLengthMm() * CDbl(s.Quantity)
            dgvSegments.Rows.Add(typeText, famText, s.Quantity.ToString(), detail, FormatNumberNoTrail(eff))
        Next
    End Sub

    Private Sub RecalcAll()
        _totBeam = New FamilyTotals()
        _totChain = New FamilyTotals()
        _totSlide = New FamilyTotals()

        For Each s As Segment In _segments
            Dim addLen As Double = s.EffectiveLengthMm() * CDbl(s.Quantity)
            Dim addBends As Integer = If(s.SegType = SegmentType.Bend, s.Quantity, 0)

            Select Case s.Family
                Case FamilyType.Beam
                    _totBeam.TotalLengthMm += addLen
                    _totBeam.TotalBends += addBends
                Case FamilyType.Chain
                    _totChain.TotalLengthMm += addLen
                    _totChain.TotalBends += addBends
                Case FamilyType.Slide
                    _totSlide.TotalLengthMm += addLen
                    _totSlide.TotalBends += addBends
            End Select
        Next

        Dim pitchMm As Double = 0
        Dim pitchOk As Boolean = TryParseDouble(txtLegSupportPitchMm.Text, pitchMm)

        Dim beamSupports As Integer = 0
        Dim chainSupports As Integer = 0
        Dim slideSupports As Integer = 0

        If chkIncludeLegSupports.Checked AndAlso pitchOk AndAlso pitchMm > 0 Then
            beamSupports = ComputeSupportCount(_totBeam.TotalLengthMm, pitchMm)
            chainSupports = ComputeSupportCount(_totChain.TotalLengthMm, pitchMm)
            slideSupports = ComputeSupportCount(_totSlide.TotalLengthMm, pitchMm)
        End If

        lvTotals.BeginUpdate()
        lvTotals.Items.Clear()

        AddTotalsRows("Beam", _totBeam, nudBeamPackLenM.Value, beamSupports)
        AddTotalsRows("Chain", _totChain, nudChainPackLenM.Value, chainSupports)
        AddTotalsRows("Slide", _totSlide, nudSlidePackLenM.Value, slideSupports)

        lvTotals.EndUpdate()

        lblBeamSummary.Text = MakePackSummary(_totBeam, nudBeamPackLenM.Value, "Beam")
        lblChainSummary.Text = MakePackSummary(_totChain, nudChainPackLenM.Value, "Chain")
        lblSlideSummary.Text = MakePackSummary(_totSlide, nudSlidePackLenM.Value, "Slide")
    End Sub

    Private Sub AddTotalsRows(ByVal familyName As String,
                             ByVal t As FamilyTotals,
                             ByVal packLenM As Decimal,
                             ByVal supportsCount As Integer)

        Dim totalM As Double = t.TotalLengthMm / 1000.0R

        Dim liLen As New ListViewItem(familyName)
        liLen.SubItems.Add("Total length (incl. bends)")
        liLen.SubItems.Add(FormatNumberNoTrail(totalM))
        liLen.SubItems.Add("m")
        liLen.SubItems.Add("Straight + bend arc length combined")
        lvTotals.Items.Add(liLen)

        Dim liBend As New ListViewItem(familyName) ' (NOT liB -> LIB keyword)
        liBend.SubItems.Add("Total bends")
        liBend.SubItems.Add(t.TotalBends.ToString())
        liBend.SubItems.Add("pcs")
        liBend.SubItems.Add("Bend segments included in totals")
        lvTotals.Items.Add(liBend)

        Dim packs As Integer = ComputePacks(totalM, CDbl(packLenM))
        Dim liPack As New ListViewItem(familyName)
        liPack.SubItems.Add("Packaging packs")
        liPack.SubItems.Add(packs.ToString())
        liPack.SubItems.Add("packs")
        liPack.SubItems.Add("Pack length = " & packLenM.ToString() & " m")
        lvTotals.Items.Add(liPack)

        Dim liSupport As New ListViewItem(familyName)
        liSupport.SubItems.Add("Leg supports")
        liSupport.SubItems.Add(supportsCount.ToString())
        liSupport.SubItems.Add("pcs")
        liSupport.SubItems.Add(If(chkIncludeLegSupports.Checked, "Pitch=" & txtLegSupportPitchMm.Text.Trim() & " mm", "Excluded"))
        lvTotals.Items.Add(liSupport)
    End Sub

    Private Function ComputeSupportCount(ByVal totalLenMm As Double, ByVal pitchMm As Double) As Integer
        If totalLenMm <= 0 Then Return 0
        Dim c As Integer = CInt(Math.Ceiling(totalLenMm / pitchMm))
        If c < 1 Then c = 1
        Return c
    End Function

    Private Function ComputePacks(ByVal totalM As Double, ByVal packLenM As Double) As Integer
        If totalM <= 0 Then Return 0
        If packLenM <= 0 Then Return 0
        Dim packs As Integer = CInt(Math.Ceiling(totalM / packLenM))
        If packs < 1 Then packs = 1
        Return packs
    End Function

    Private Function MakePackSummary(ByVal t As FamilyTotals, ByVal packLenM As Decimal, ByVal familyName As String) As String
        Dim totalM As Double = t.TotalLengthMm / 1000.0R
        Dim packs As Integer = ComputePacks(totalM, CDbl(packLenM))
        Return familyName & " totals:" & Environment.NewLine &
               "• Length: " & FormatNumberNoTrail(totalM) & " m" & Environment.NewLine &
               "• Bends: " & t.TotalBends.ToString() & " pcs" & Environment.NewLine &
               "• Packs: " & packs.ToString() & "  (@" & packLenM.ToString() & " m/pack)"
    End Function

    Private Function FamilyToText(ByVal f As FamilyType) As String
        Select Case f
            Case FamilyType.Beam : Return "Beam"
            Case FamilyType.Chain : Return "Chain"
            Case Else : Return "Slide"
        End Select
    End Function

    Private Function TryParseDouble(ByVal s As String, ByRef result As Double) As Boolean
        result = 0
        If s Is Nothing Then Return False
        Dim t As String = s.Trim()
        If t.Length = 0 Then Return False
        t = t.Replace(",", ".")
        Return Double.TryParse(t, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, result)
    End Function

    Private Function FormatNumberNoTrail(ByVal v As Double) As String
        Return v.ToString("0.##", Globalization.CultureInfo.InvariantCulture)
    End Function

    '============================================================
    ' Printing
    '============================================================
    Private Sub OnPrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        Dim g As Graphics = e.Graphics
        Dim margin As Rectangle = e.MarginBounds

        Dim y As Integer = margin.Top
        Dim x As Integer = margin.Left
        Dim lineH As Integer = 22

        Dim titleFont As New Font("Segoe UI", 16.0F, FontStyle.Bold, GraphicsUnit.Point)
        Dim hFont As New Font("Segoe UI", 11.0F, FontStyle.Bold, GraphicsUnit.Point)
        Dim bFont As New Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)

        g.DrawString("FlexLink Project — Overall BOM Totals", titleFont, Brushes.Black, x, y)
        y += 40

        g.DrawString("Totals", hFont, Brushes.Black, x, y)
        y += 26

        Dim colW0 As Integer = 90
        Dim colW1 As Integer = 220
        Dim colW2 As Integer = 70
        Dim colW3 As Integer = 60
        Dim colW4 As Integer = margin.Width - (colW0 + colW1 + colW2 + colW3) - 10

        DrawTableRow(g, hFont, x, y, colW0, colW1, colW2, colW3, colW4, True, "Family", "Item", "Qty", "Unit", "Notes")
        y += lineH + 6

        For Each it As ListViewItem In lvTotals.Items
            If y > margin.Bottom - 140 Then Exit For

            DrawTableRow(g, bFont, x, y, colW0, colW1, colW2, colW3, colW4, False,
                         it.Text, it.SubItems(1).Text, it.SubItems(2).Text, it.SubItems(3).Text, it.SubItems(4).Text)
            y += lineH + 2
        Next

        y += 16
        g.DrawString("Packaging", hFont, Brushes.Black, x, y)
        y += 26

        g.DrawString("Beam: " & lblBeamSummary.Text.Replace(Environment.NewLine, " | "), bFont, Brushes.Black, x, y)
        y += 22
        g.DrawString("Chain: " & lblChainSummary.Text.Replace(Environment.NewLine, " | "), bFont, Brushes.Black, x, y)
        y += 22
        g.DrawString("Slide: " & lblSlideSummary.Text.Replace(Environment.NewLine, " | "), bFont, Brushes.Black, x, y)

        e.HasMorePages = False
    End Sub

    Private Sub DrawTableRow(ByVal g As Graphics,
                             ByVal fnt As Font,
                             ByVal x As Integer,
                             ByVal y As Integer,
                             ByVal w0 As Integer, ByVal w1 As Integer, ByVal w2 As Integer, ByVal w3 As Integer, ByVal w4 As Integer,
                             ByVal isHeader As Boolean,
                             ByVal c0 As String, ByVal c1 As String, ByVal c2 As String, ByVal c3 As String, ByVal c4 As String)

        Dim bg As Brush = If(isHeader, Brushes.LightGray, Brushes.White)
        Dim pen As New Pen(Color.Black)
        Dim h As Integer = 22

        Dim r0 As New Rectangle(x, y, w0, h)
        Dim r1 As New Rectangle(x + w0, y, w1, h)
        Dim r2 As New Rectangle(x + w0 + w1, y, w2, h)
        Dim r3 As New Rectangle(x + w0 + w1 + w2, y, w3, h)
        Dim r4 As New Rectangle(x + w0 + w1 + w2 + w3, y, w4, h)

        g.FillRectangle(bg, r0) : g.FillRectangle(bg, r1) : g.FillRectangle(bg, r2) : g.FillRectangle(bg, r3) : g.FillRectangle(bg, r4)
        g.DrawRectangle(pen, r0) : g.DrawRectangle(pen, r1) : g.DrawRectangle(pen, r2) : g.DrawRectangle(pen, r3) : g.DrawRectangle(pen, r4)

        g.DrawString(TrimToFit(g, fnt, c0, w0 - 6), fnt, Brushes.Black, x + 3, y + 3)
        g.DrawString(TrimToFit(g, fnt, c1, w1 - 6), fnt, Brushes.Black, x + w0 + 3, y + 3)
        g.DrawString(TrimToFit(g, fnt, c2, w2 - 6), fnt, Brushes.Black, x + w0 + w1 + 3, y + 3)
        g.DrawString(TrimToFit(g, fnt, c3, w3 - 6), fnt, Brushes.Black, x + w0 + w1 + w2 + 3, y + 3)
        g.DrawString(TrimToFit(g, fnt, c4, w4 - 6), fnt, Brushes.Black, x + w0 + w1 + w2 + w3 + 3, y + 3)
    End Sub

    Private Function TrimToFit(ByVal g As Graphics, ByVal fnt As Font, ByVal text As String, ByVal maxWidth As Integer) As String
        If text Is Nothing Then Return ""
        Dim t As String = text
        If g.MeasureString(t, fnt).Width <= CSng(maxWidth) Then Return t
        While t.Length > 3 AndAlso g.MeasureString(t & "...", fnt).Width > CSng(maxWidth)
            t = t.Substring(0, t.Length - 1)
        End While
        Return t & "..."
    End Function

End Class
