Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Collections.Generic

Public Class AirConsumptionForm
    Inherits Form

    ' -----------------------------
    ' DATA
    ' -----------------------------
    Private dt As DataTable
    Private dv As DataView

    ' -----------------------------
    ' UI
    ' -----------------------------
    Private tip As ToolTip

    Private pnlTop As Panel
    Private pnlGrid As Panel
    Private pnlRight As Panel
    Private pnlBottom As Panel

    Private dgv As DataGridView

    ' Inputs
    Private txtName As TextBox
    Private cboAction As ComboBox
    Private nudBore As NumericUpDown
    Private nudRod As NumericUpDown
    Private nudStroke As NumericUpDown
    Private nudCycles As NumericUpDown
    Private nudQty As NumericUpDown
    Private nudPressure As NumericUpDown
    Private chkIncludeDefault As CheckBox

    ' Buttons
    Private btnAdd As Button
    Private btnPaste As Button
    Private btnSort As Button
    Private btnDelete As Button
    Private btnClear As Button
    Private btnExportPdf As Button
    Private btnCopyText As Button

    ' Summary + calculation log
    Private lblSummary As Label
    Private txtCalc As TextBox

    ' Constants
    Private ReadOnly Inv As CultureInfo = CultureInfo.InvariantCulture

    Public Sub New()
        Me.Text = "Air Consumption & Compressor Sizing"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Font = New Font("Segoe UI", 9.5F, FontStyle.Regular)
        Me.MinimumSize = New Size(1050, 730)
        Me.Size = New Size(1180, 800)

        FormFooter.AddPremiumFooter(Me)

        tip = New ToolTip()
        tip.AutoPopDelay = 12000
        tip.InitialDelay = 350
        tip.ReshowDelay = 150
        tip.ShowAlways = True

        BuildData()
        BuildUI()
        WireEvents()

        RecalcAll()
        UpdateSummaryAndLog()

        FormHeader.AddPremiumHeader(Me, "Air Consumption Calculator", "MetaMech Engineering Tools")
    End Sub

    ' =========================================================
    ' DATA TABLE
    ' =========================================================
    Private Sub BuildData()
        dt = New DataTable("Cylinders")

        ' IMPORTANT: Must exist to avoid "Column named Include cannot be found"
        dt.Columns.Add("Include", GetType(Boolean))
        dt.Columns.Add("Name", GetType(String))
        dt.Columns.Add("Action", GetType(String)) ' "DA" or "SA"
        dt.Columns.Add("Bore_mm", GetType(Double))
        dt.Columns.Add("Rod_mm", GetType(Double))
        dt.Columns.Add("Stroke_mm", GetType(Double))
        dt.Columns.Add("Cycles_per_min", GetType(Double))
        dt.Columns.Add("Qty", GetType(Integer))
        dt.Columns.Add("Pressure_bar_g", GetType(Double))

        dt.Columns.Add("NL_per_cycle", GetType(Double))
        dt.Columns.Add("NL_per_min", GetType(Double))

        dv = New DataView(dt)
    End Sub

    ' =========================================================
    ' UI
    ' =========================================================
    Private Sub BuildUI()
        pnlTop = New Panel() With {.Dock = DockStyle.Top, .Height = 150, .Padding = New Padding(10)}
        pnlGrid = New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(10)}
        pnlRight = New Panel() With {.Dock = DockStyle.Right, .Width = 360, .Padding = New Padding(10)}
        pnlBottom = New Panel() With {.Dock = DockStyle.Bottom, .Height = 48, .Padding = New Padding(10)}

        Me.Controls.Add(pnlGrid)
        Me.Controls.Add(pnlRight)
        Me.Controls.Add(pnlTop)
        Me.Controls.Add(pnlBottom)

        ' --- TOP
        Dim lblTitle As New Label() With {
            .Text = "Add / Edit Cylinder",
            .Font = New Font("Segoe UI", 11.5F, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(10, 10)
        }
        pnlTop.Controls.Add(lblTitle)

        Dim y0 As Integer = 44

        Dim lbl1 As New Label() With {.Text = "Name / Tag", .AutoSize = True, .Location = New Point(10, y0)}
        txtName = New TextBox() With {.Location = New Point(10, y0 + 22), .Width = 200}
        pnlTop.Controls.Add(lbl1) : pnlTop.Controls.Add(txtName)

        Dim lbl2 As New Label() With {.Text = "Action (DA / SA)", .AutoSize = True, .Location = New Point(225, y0)}
        cboAction = New ComboBox() With {.Location = New Point(225, y0 + 22), .Width = 110, .DropDownStyle = ComboBoxStyle.DropDownList}
        cboAction.Items.AddRange(New Object() {"DA", "SA"})
        cboAction.SelectedIndex = 0
        pnlTop.Controls.Add(lbl2) : pnlTop.Controls.Add(cboAction)

        Dim lbl3 As New Label() With {.Text = "Bore (mm)", .AutoSize = True, .Location = New Point(350, y0)}
        nudBore = New NumericUpDown() With {.Location = New Point(350, y0 + 22), .Width = 90, .Minimum = 1D, .Maximum = 1000D, .DecimalPlaces = 0, .Value = 32D}
        pnlTop.Controls.Add(lbl3) : pnlTop.Controls.Add(nudBore)

        Dim lbl4 As New Label() With {.Text = "Rod (mm)", .AutoSize = True, .Location = New Point(455, y0)}
        nudRod = New NumericUpDown() With {.Location = New Point(455, y0 + 22), .Width = 90, .Minimum = 0D, .Maximum = 1000D, .DecimalPlaces = 0, .Value = 12D}
        pnlTop.Controls.Add(lbl4) : pnlTop.Controls.Add(nudRod)

        Dim lbl5 As New Label() With {.Text = "Stroke (mm)", .AutoSize = True, .Location = New Point(560, y0)}
        nudStroke = New NumericUpDown() With {.Location = New Point(560, y0 + 22), .Width = 95, .Minimum = 1D, .Maximum = 5000D, .DecimalPlaces = 0, .Value = 100D}
        pnlTop.Controls.Add(lbl5) : pnlTop.Controls.Add(nudStroke)

        Dim lbl6 As New Label() With {.Text = "Cycles / min", .AutoSize = True, .Location = New Point(670, y0)}
        nudCycles = New NumericUpDown() With {.Location = New Point(670, y0 + 22), .Width = 95, .Minimum = 0D, .Maximum = 6000D, .DecimalPlaces = 1, .Increment = 0.5D, .Value = 10D}
        pnlTop.Controls.Add(lbl6) : pnlTop.Controls.Add(nudCycles)

        Dim lbl7 As New Label() With {.Text = "Qty", .AutoSize = True, .Location = New Point(780, y0)}
        nudQty = New NumericUpDown() With {.Location = New Point(780, y0 + 22), .Width = 70, .Minimum = 1D, .Maximum = 9999D, .DecimalPlaces = 0, .Value = 1D}
        pnlTop.Controls.Add(lbl7) : pnlTop.Controls.Add(nudQty)

        Dim lbl8 As New Label() With {.Text = "Supply Pressure (bar g)", .AutoSize = True, .Location = New Point(865, y0)}
        nudPressure = New NumericUpDown() With {.Location = New Point(865, y0 + 22), .Width = 120, .Minimum = 0D, .Maximum = 20D, .DecimalPlaces = 2, .Increment = 0.1D, .Value = 6D}
        pnlTop.Controls.Add(lbl8) : pnlTop.Controls.Add(nudPressure)

        chkIncludeDefault = New CheckBox() With {.Text = "Include", .Checked = True, .AutoSize = True, .Location = New Point(10, y0 + 60)}
        pnlTop.Controls.Add(chkIncludeDefault)

        btnAdd = New Button() With {.Text = "Add Cylinder", .Width = 140, .Height = 32, .Location = New Point(225, y0 + 56)}
        btnPaste = New Button() With {.Text = "Paste from Excel", .Width = 140, .Height = 32, .Location = New Point(370, y0 + 56)}
        btnSort = New Button() With {.Text = "Sort by NL/min", .Width = 140, .Height = 32, .Location = New Point(515, y0 + 56)}
        btnDelete = New Button() With {.Text = "Delete Selected", .Width = 140, .Height = 32, .Location = New Point(660, y0 + 56)}
        btnClear = New Button() With {.Text = "Clear All", .Width = 110, .Height = 32, .Location = New Point(805, y0 + 56)}

        pnlTop.Controls.Add(btnAdd)
        pnlTop.Controls.Add(btnPaste)
        pnlTop.Controls.Add(btnSort)
        pnlTop.Controls.Add(btnDelete)
        pnlTop.Controls.Add(btnClear)

        ' --- GRID
        dgv = New DataGridView() With {
            .Dock = DockStyle.Fill,
            .AutoGenerateColumns = False,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = True,
            .RowHeadersVisible = False
        }

        dgv.DataSource = dv
        pnlGrid.Controls.Add(dgv)

        AddGridColumns()

        ' --- RIGHT: Summary + Method
        Dim lblR As New Label() With {.Text = "Calculation / Method Log", .AutoSize = True, .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold)}
        lblR.Location = New Point(10, 10)
        pnlRight.Controls.Add(lblR)

        lblSummary = New Label() With {.Text = "", .AutoSize = False, .Location = New Point(10, 42), .Size = New Size(pnlRight.Width - 20, 90)}
        lblSummary.Font = New Font("Segoe UI", 9.2F, FontStyle.Regular)
        pnlRight.Controls.Add(lblSummary)

        txtCalc = New TextBox() With {.Multiline = True, .ReadOnly = True, .ScrollBars = ScrollBars.Vertical}
        txtCalc.Location = New Point(10, 140)
        txtCalc.Size = New Size(pnlRight.Width - 20, pnlRight.Height - 190)
        txtCalc.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        pnlRight.Controls.Add(txtCalc)

        ' --- BOTTOM: Export/Copy
        btnExportPdf = New Button() With {.Text = "Export PDF Report", .Width = 170, .Height = 30, .Location = New Point(10, 8)}
        btnCopyText = New Button() With {.Text = "Copy Report Text", .Width = 160, .Height = 30, .Location = New Point(190, 8)}
        pnlBottom.Controls.Add(btnExportPdf)
        pnlBottom.Controls.Add(btnCopyText)

        ApplyTooltips()
    End Sub

    Private Sub AddGridColumns()
        dgv.Columns.Clear()

        Dim cInclude As New DataGridViewCheckBoxColumn()
        cInclude.DataPropertyName = "Include"
        cInclude.HeaderText = "Include"
        cInclude.Width = 60
        dgv.Columns.Add(cInclude)

        dgv.Columns.Add(MkTextCol("Name", "Name", 160))
        dgv.Columns.Add(MkTextCol("Action", "Action", 60))
        dgv.Columns.Add(MkNumCol("Bore_mm", "Bore (mm)", 85))
        dgv.Columns.Add(MkNumCol("Rod_mm", "Rod (mm)", 85))
        dgv.Columns.Add(MkNumCol("Stroke_mm", "Stroke (mm)", 95))
        dgv.Columns.Add(MkNumCol("Cycles_per_min", "Cycles/min", 95))
        dgv.Columns.Add(MkNumCol("Qty", "Qty", 55))
        dgv.Columns.Add(MkNumCol("Pressure_bar_g", "P (bar g)", 85))

        dgv.Columns.Add(MkNumCol("NL_per_cycle", "NL/cycle", 90))
        dgv.Columns.Add(MkNumCol("NL_per_min", "NL/min", 90))
    End Sub

    Private Function MkTextCol(prop As String, header As String, w As Integer) As DataGridViewTextBoxColumn
        Dim c As New DataGridViewTextBoxColumn()
        c.DataPropertyName = prop
        c.HeaderText = header
        c.Width = w
        c.SortMode = DataGridViewColumnSortMode.Programmatic
        Return c
    End Function

    Private Function MkNumCol(prop As String, header As String, w As Integer) As DataGridViewTextBoxColumn
        Dim c As New DataGridViewTextBoxColumn()
        c.DataPropertyName = prop
        c.HeaderText = header
        c.Width = w
        c.SortMode = DataGridViewColumnSortMode.Programmatic
        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        c.DefaultCellStyle.Format = "0.###"
        Return c
    End Function

    Private Sub ApplyTooltips()
        tip.SetToolTip(txtName, "Optional tag for your cylinder (example: Clamp-01, Stopper-Left).")
        tip.SetToolTip(cboAction,
            "Action type:" & vbCrLf &
            "DA = Double Acting (air used for extend + retract)." & vbCrLf &
            "SA = Single Acting (air used in one direction, spring return).")

        tip.SetToolTip(nudBore, "Cylinder bore diameter in mm (example: 32, 40, 63).")
        tip.SetToolTip(nudRod, "Rod diameter in mm (used for DA retract volume). If unknown, leave typical value.")
        tip.SetToolTip(nudStroke, "Stroke length in mm (travel distance).")
        tip.SetToolTip(nudCycles, "Cycles per minute (complete cycles per minute). Example: 10 cycles/min.")
        tip.SetToolTip(nudQty, "How many identical cylinders of this type in the project.")
        tip.SetToolTip(nudPressure, "Supply pressure in bar gauge (bar g). Typical: 6.0 bar g.")
        tip.SetToolTip(chkIncludeDefault, "If unticked, cylinder is excluded from totals (useful for optional items).")

        tip.SetToolTip(btnAdd, "Adds cylinder to the list and calculates NL/cycle and NL/min.")
        tip.SetToolTip(btnPaste,
            "Paste multiple cylinders from Excel/Sheets." & vbCrLf &
            "Supports headers like: Name, Action, Bore, Rod, Stroke, Cycles/min, Qty, Pressure, Include." & vbCrLf &
            "Or no headers (fixed order).")
        tip.SetToolTip(btnSort, "Sort list by highest NL/min first (largest consumers at top).")
        tip.SetToolTip(btnDelete, "Deletes selected rows.")
        tip.SetToolTip(btnClear, "Clears the whole project list.")
        tip.SetToolTip(btnExportPdf, "Exports a multi-page PDF: summary + top consumers + tables + method appendix.")
        tip.SetToolTip(btnCopyText, "Copies a text report to clipboard (easy for emails).")

        tip.SetToolTip(dgv,
            "Edit cells directly. Include checkbox controls whether a row contributes to totals." & vbCrLf &
            "DA = extend + retract air consumption; SA = extend only (spring return).")
    End Sub

    Private Sub WireEvents()
        AddHandler btnAdd.Click, AddressOf OnAdd
        AddHandler btnPaste.Click, AddressOf OnPaste
        AddHandler btnSort.Click, AddressOf OnSort
        AddHandler btnDelete.Click, AddressOf OnDelete
        AddHandler btnClear.Click, AddressOf OnClear
        AddHandler btnExportPdf.Click, AddressOf OnExportPdf
        AddHandler btnCopyText.Click, AddressOf OnCopyText

        AddHandler dgv.CellEndEdit, AddressOf OnGridEdited
        AddHandler dgv.CurrentCellDirtyStateChanged, AddressOf OnDirtyChanged
        AddHandler dgv.SelectionChanged, AddressOf OnSelectionChanged
        AddHandler Me.Resize, AddressOf OnResized
    End Sub

    Private Sub OnResized(sender As Object, e As EventArgs)
        Try
            lblSummary.Size = New Size(pnlRight.Width - 20, 90)
            txtCalc.Size = New Size(pnlRight.Width - 20, pnlRight.Height - 190)
        Catch
        End Try
    End Sub

    ' =========================================================
    ' ADD / EDIT
    ' =========================================================
    Private Sub OnAdd(sender As Object, e As EventArgs)
        Dim name As String = txtName.Text.Trim()
        If name = "" Then name = "Cylinder"

        Dim action As String = "DA"
        Try
            action = cboAction.SelectedItem.ToString().Trim().ToUpperInvariant()
        Catch
            action = "DA"
        End Try
        If action <> "DA" AndAlso action <> "SA" Then action = "DA"

        Dim boreMm As Double = CDbl(nudBore.Value)
        Dim rodMm As Double = CDbl(nudRod.Value)
        Dim strokeMm As Double = CDbl(nudStroke.Value)
        Dim cycles As Double = CDbl(nudCycles.Value)
        Dim qty As Integer = CInt(nudQty.Value)
        Dim pBarG As Double = CDbl(nudPressure.Value)
        Dim inc As Boolean = chkIncludeDefault.Checked

        Dim r As DataRow = dt.NewRow()
        r("Include") = inc
        r("Name") = name
        r("Action") = action
        r("Bore_mm") = boreMm
        r("Rod_mm") = rodMm
        r("Stroke_mm") = strokeMm
        r("Cycles_per_min") = cycles
        r("Qty") = qty
        r("Pressure_bar_g") = pBarG

        dt.Rows.Add(r)

        RecalcRow(r)
        UpdateSummaryAndLog()
    End Sub

    Private Sub OnGridEdited(sender As Object, e As DataGridViewCellEventArgs)
        Try
            If e Is Nothing OrElse e.RowIndex < 0 Then Exit Sub
            Dim drv As DataRowView = TryCast(dv(e.RowIndex), DataRowView)
            If drv Is Nothing Then Exit Sub
            RecalcRow(drv.Row)
            UpdateSummaryAndLog()
        Catch
        End Try
    End Sub

    Private Sub OnDirtyChanged(sender As Object, e As EventArgs)
        Try
            If dgv.IsCurrentCellDirty Then
                dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            End If
        Catch
        End Try
    End Sub

    ' =========================================================
    ' EXCEL PASTE
    ' =========================================================
    Private Sub OnPaste(sender As Object, e As EventArgs)
        Try
            Dim clip As String = Clipboard.GetText()
            If clip Is Nothing OrElse clip.Trim() = "" Then
                MessageBox.Show("Clipboard is empty. Copy rows from Excel first.", "Paste from Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim lines() As String = clip.Replace(vbCrLf, vbLf).Split(New Char() {ChrW(10)}, StringSplitOptions.RemoveEmptyEntries)
            If lines.Length = 0 Then Return

            Dim firstCells() As String = SplitRow(lines(0))
            Dim hasHeader As Boolean = LooksLikeHeader(firstCells)

            Dim colMap As Dictionary(Of String, Integer) = Nothing
            Dim startIdx As Integer = 0

            If hasHeader Then
                colMap = BuildHeaderMap(firstCells)
                startIdx = 1
            End If

            Dim added As Integer = 0

            For i As Integer = startIdx To lines.Length - 1
                Dim cells() As String = SplitRow(lines(i))
                If cells.Length = 0 Then Continue For

                Dim r As DataRow = dt.NewRow()

                ' Defaults
                r("Include") = True
                r("Name") = "Cylinder"
                r("Action") = "DA"
                r("Bore_mm") = 32.0
                r("Rod_mm") = 12.0
                r("Stroke_mm") = 100.0
                r("Cycles_per_min") = 10.0
                r("Qty") = 1
                r("Pressure_bar_g") = 6.0

                If hasHeader AndAlso colMap IsNot Nothing Then
                    SetFromHeader(r, cells, colMap)
                Else
                    ' Fixed order:
                    ' Name, Action, Bore, Rod, Stroke, Cycles/min, Qty, Pressure, Include
                    SetFromFixedOrder(r, cells)
                End If

                dt.Rows.Add(r)
                RecalcRow(r)
                added += 1
            Next

            UpdateSummaryAndLog()

            MessageBox.Show("Pasted " & added.ToString() & " row(s).", "Paste from Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Paste failed: " & ex.Message, "Paste from Excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function SplitRow(line As String) As String()
        If line Is Nothing Then Return New String() {}
        Dim t As String = line.Trim()
        If t = "" Then Return New String() {}
        Return t.Split(New Char() {ChrW(9)}, StringSplitOptions.None) ' TAB
    End Function

    Private Function LooksLikeHeader(cells() As String) As Boolean
        If cells Is Nothing OrElse cells.Length = 0 Then Return False
        Dim joined As String = String.Join("|", cells).ToLowerInvariant()
        Return (joined.Contains("bore") OrElse joined.Contains("stroke") OrElse joined.Contains("cycle") OrElse joined.Contains("pressure") OrElse joined.Contains("action") OrElse joined.Contains("include"))
    End Function

    Private Function BuildHeaderMap(cells() As String) As Dictionary(Of String, Integer)
        Dim m As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For i As Integer = 0 To cells.Length - 1
            Dim h As String = cells(i).Trim().ToLowerInvariant()
            If h = "" Then Continue For

            If h.Contains("name") OrElse h.Contains("tag") Then m("name") = i
            If h = "da/sa" OrElse h.Contains("action") Then m("action") = i
            If h.Contains("bore") Then m("bore") = i
            If h.Contains("rod") Then m("rod") = i
            If h.Contains("stroke") Then m("stroke") = i
            If h.Contains("cycle") Then m("cycles") = i
            If h = "qty" OrElse h.Contains("quantity") Then m("qty") = i
            If h.Contains("press") Then m("pressure") = i
            If h.Contains("include") OrElse h.Contains("use") Then m("include") = i
        Next
        Return m
    End Function

    Private Sub SetFromHeader(r As DataRow, cells() As String, m As Dictionary(Of String, Integer))
        Try
            If m.ContainsKey("name") Then r("Name") = SafeCell(cells, m("name"), "Cylinder")
            If m.ContainsKey("action") Then r("Action") = NormalizeAction(SafeCell(cells, m("action"), "DA"))
            If m.ContainsKey("bore") Then r("Bore_mm") = SafeDbl(SafeCell(cells, m("bore"), "32"))
            If m.ContainsKey("rod") Then r("Rod_mm") = SafeDbl(SafeCell(cells, m("rod"), "12"))
            If m.ContainsKey("stroke") Then r("Stroke_mm") = SafeDbl(SafeCell(cells, m("stroke"), "100"))
            If m.ContainsKey("cycles") Then r("Cycles_per_min") = SafeDbl(SafeCell(cells, m("cycles"), "10"))
            If m.ContainsKey("qty") Then r("Qty") = SafeInt(SafeCell(cells, m("qty"), "1"))
            If m.ContainsKey("pressure") Then r("Pressure_bar_g") = SafeDbl(SafeCell(cells, m("pressure"), "6"))
            If m.ContainsKey("include") Then r("Include") = SafeBool(SafeCell(cells, m("include"), "true"))
        Catch
        End Try
    End Sub

    Private Sub SetFromFixedOrder(r As DataRow, cells() As String)
        Try
            If cells.Length > 0 Then r("Name") = If(cells(0).Trim() = "", "Cylinder", cells(0).Trim())
            If cells.Length > 1 Then r("Action") = NormalizeAction(cells(1))
            If cells.Length > 2 Then r("Bore_mm") = SafeDbl(cells(2))
            If cells.Length > 3 Then r("Rod_mm") = SafeDbl(cells(3))
            If cells.Length > 4 Then r("Stroke_mm") = SafeDbl(cells(4))
            If cells.Length > 5 Then r("Cycles_per_min") = SafeDbl(cells(5))
            If cells.Length > 6 Then r("Qty") = SafeInt(cells(6))
            If cells.Length > 7 Then r("Pressure_bar_g") = SafeDbl(cells(7))
            If cells.Length > 8 Then r("Include") = SafeBool(cells(8))
        Catch
        End Try
    End Sub

    Private Function SafeCell(cells() As String, idx As Integer, def As String) As String
        If cells Is Nothing Then Return def
        If idx < 0 OrElse idx >= cells.Length Then Return def
        Dim s As String = cells(idx)
        If s Is Nothing Then Return def
        s = s.Trim()
        If s = "" Then Return def
        Return s
    End Function

    Private Function SafeDbl(s As String) As Double
        If s Is Nothing Then Return 0
        Dim t As String = s.Trim().Replace(",", ".")
        Dim v As Double = 0
        Double.TryParse(t, NumberStyles.Any, Inv, v)
        Return v
    End Function

    Private Function SafeInt(s As String) As Integer
        If s Is Nothing Then Return 0
        Dim t As String = s.Trim()
        Dim v As Integer = 0
        Integer.TryParse(t, v)
        Return v
    End Function

    Private Function SafeBool(s As String) As Boolean
        If s Is Nothing Then Return True
        Dim t As String = s.Trim().ToLowerInvariant()
        If t = "0" OrElse t = "false" OrElse t = "no" OrElse t = "off" Then Return False
        If t = "1" OrElse t = "true" OrElse t = "yes" OrElse t = "on" Then Return True
        Return True
    End Function

    Private Function NormalizeAction(s As String) As String
        If s Is Nothing Then Return "DA"
        Dim t As String = s.Trim().ToUpperInvariant()
        If t.Contains("DOUBLE") Then Return "DA"
        If t.Contains("SINGLE") Then Return "SA"
        If t = "DA" OrElse t = "SA" Then Return t
        If t = "D" Then Return "DA"
        If t = "S" Then Return "SA"
        Return "DA"
    End Function

    ' =========================================================
    ' SORT / DELETE / CLEAR
    ' =========================================================
    Private Sub OnSort(sender As Object, e As EventArgs)
        Try
            dv.Sort = "NL_per_min DESC"
        Catch
        End Try
        UpdateSummaryAndLog()
    End Sub

    Private Sub OnDelete(sender As Object, e As EventArgs)
        Try
            If dgv.SelectedRows Is Nothing OrElse dgv.SelectedRows.Count = 0 Then Return

            Dim idxs As New List(Of Integer)()
            For Each r As DataGridViewRow In dgv.SelectedRows
                idxs.Add(r.Index)
            Next

            idxs.Sort()
            idxs.Reverse()

            For Each idx As Integer In idxs
                If idx >= 0 AndAlso idx < dv.Count Then
                    dv.Delete(idx)
                End If
            Next

            RecalcAll()
            UpdateSummaryAndLog()
        Catch
        End Try
    End Sub

    Private Sub OnClear(sender As Object, e As EventArgs)
        If MessageBox.Show("Clear all cylinders?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        Try
            dt.Rows.Clear()
            UpdateSummaryAndLog()
        Catch
        End Try
    End Sub

    ' =========================================================
    ' CALCULATIONS
    ' =========================================================
    Private Sub RecalcAll()
        Try
            For Each r As DataRow In dt.Rows
                RecalcRow(r)
            Next
        Catch
        End Try
    End Sub

    Private Sub RecalcRow(r As DataRow)
        Try
            Dim action As String = Convert.ToString(r("Action")).Trim().ToUpperInvariant()
            If action <> "DA" AndAlso action <> "SA" Then action = "DA"

            Dim boreMm As Double = SafeDbl(Convert.ToString(r("Bore_mm")))
            Dim rodMm As Double = SafeDbl(Convert.ToString(r("Rod_mm")))
            Dim strokeMm As Double = SafeDbl(Convert.ToString(r("Stroke_mm")))
            Dim cycles As Double = SafeDbl(Convert.ToString(r("Cycles_per_min")))

            Dim qty As Integer = 1
            Try
                qty = CInt(r("Qty"))
            Catch
                qty = SafeInt(Convert.ToString(r("Qty")))
            End Try
            If qty <= 0 Then qty = 1

            Dim pBarG As Double = SafeDbl(Convert.ToString(r("Pressure_bar_g")))
            If pBarG < 0 Then pBarG = 0

            ' mm -> m
            Dim boreM As Double = boreMm / 1000.0
            Dim rodM As Double = rodMm / 1000.0
            Dim strokeM As Double = strokeMm / 1000.0

            Dim areaBore As Double = Math.PI * (boreM * boreM) / 4.0
            Dim areaRod As Double = Math.PI * (rodM * rodM) / 4.0
            Dim areaAnnulus As Double = areaBore - areaRod
            If areaAnnulus < 0 Then areaAnnulus = 0

            Dim volExtend_m3 As Double = areaBore * strokeM
            Dim volRetract_m3 As Double = areaAnnulus * strokeM

            ' m3 -> L
            Dim volExtend_L As Double = volExtend_m3 * 1000.0
            Dim volRetract_L As Double = volRetract_m3 * 1000.0

            ' NL approx using absolute pressure ratio
            Dim pAbs As Double = pBarG + 1.0

            Dim nlPerCycleOne As Double
            If action = "DA" Then
                nlPerCycleOne = (volExtend_L + volRetract_L) * pAbs
            Else
                nlPerCycleOne = volExtend_L * pAbs
            End If

            Dim nlPerCycleAll As Double = nlPerCycleOne * CDbl(qty)
            Dim nlPerMinAll As Double = nlPerCycleAll * cycles

            r("NL_per_cycle") = Math.Max(0.0, nlPerCycleAll)
            r("NL_per_min") = Math.Max(0.0, nlPerMinAll)

        Catch
            Try
                r("NL_per_cycle") = 0.0
                r("NL_per_min") = 0.0
            Catch
            End Try
        End Try
    End Sub

    ' =========================================================
    ' SUMMARY + METHOD LOG
    ' =========================================================
    Private Sub OnSelectionChanged(sender As Object, e As EventArgs)
        UpdateSummaryAndLog()
    End Sub

    Private Sub UpdateSummaryAndLog()
        Dim totalNlMin As Double = 0.0
        Dim includedCount As Integer = 0
        Dim totalRows As Integer = dt.Rows.Count

        For Each r As DataRow In dt.Rows
            Dim inc As Boolean = True
            Try
                inc = CBool(r("Include"))
            Catch
                inc = SafeBool(Convert.ToString(r("Include")))
            End Try

            If inc Then
                includedCount += 1
                totalNlMin += SafeDbl(Convert.ToString(r("NL_per_min")))
            End If
        Next

        Dim safetyFactor As Double = 1.25
        Dim suggestedNlMin As Double = totalNlMin * safetyFactor
        Dim suggested_m3h As Double = (suggestedNlMin / 1000.0) * 60.0

        lblSummary.Text =
            "Rows: " & totalRows.ToString() & "  (Included: " & includedCount.ToString() & ")" & vbCrLf &
            "Total (Included) = " & totalNlMin.ToString("0.##", Inv) & " NL/min" & vbCrLf &
            "Suggested compressor (x" & safetyFactor.ToString("0.##", Inv) & ") = " & suggestedNlMin.ToString("0.##", Inv) & " NL/min" & vbCrLf &
            "Approx = " & suggested_m3h.ToString("0.##", Inv) & " m^3/h (free air)"

        txtCalc.Text = BuildMethodLog(totalNlMin, safetyFactor)
    End Sub

    Private Function BuildMethodLog(totalNlMin As Double, safetyFactor As Double) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("DA / SA meaning:")
        sb.AppendLine("  DA = Double Acting: air used for EXTEND + RETRACT.")
        sb.AppendLine("  SA = Single Acting: air used one direction (typically EXTEND), spring return.")
        sb.AppendLine()

        sb.AppendLine("Core method (per cylinder):")
        sb.AppendLine("  1) Abore = pi * (Bore^2) / 4")
        sb.AppendLine("  2) Arod  = pi * (Rod^2)  / 4")
        sb.AppendLine("  3) Vext  = Abore * Stroke")
        sb.AppendLine("  4) Vret  = (Abore - Arod) * Stroke  (DA only)")
        sb.AppendLine("  5) Convert to litres: L = m^3 * 1000")
        sb.AppendLine("  6) Convert to Normal Litres: NL = L * Pabs   where Pabs = Pbar(g) + 1")
        sb.AppendLine("  7) Multiply by Qty and Cycles/min")
        sb.AppendLine()

        sb.AppendLine("Project totals (Included only):")
        sb.AppendLine("  Total NL/min = " & totalNlMin.ToString("0.###", Inv))
        sb.AppendLine("  Suggested compressor = Total * " & safetyFactor.ToString("0.##", Inv))
        sb.AppendLine("  Suggested NL/min = " & (totalNlMin * safetyFactor).ToString("0.###", Inv))
        sb.AppendLine()

        sb.AppendLine("Notes:")
        sb.AppendLine("  - Fast sizing method (good estimate).")
        sb.AppendLine("  - Add extra margin for leaks, blow-off, long hoses, high duty tools.")

        Return sb.ToString()
    End Function

    ' =========================================================
    ' REPORT BUILDERS
    ' =========================================================
    Private Class RowInfo
        Public Include As Boolean
        Public Name As String
        Public Action As String
        Public Bore As Double
        Public Rod As Double
        Public Stroke As Double
        Public Cycles As Double
        Public Qty As Integer
        Public Pressure As Double
        Public NLcycle As Double
        Public NLmin As Double
    End Class

    Private Function GetRowsSnapshot() As List(Of RowInfo)
        Dim list As New List(Of RowInfo)()
        For Each r As DataRow In dt.Rows
            Dim ri As New RowInfo()
            Try
                ri.Include = CBool(r("Include"))
            Catch
                ri.Include = SafeBool(Convert.ToString(r("Include")))
            End Try
            ri.Name = Convert.ToString(r("Name"))
            ri.Action = Convert.ToString(r("Action")).Trim().ToUpperInvariant()
            ri.Bore = SafeDbl(Convert.ToString(r("Bore_mm")))
            ri.Rod = SafeDbl(Convert.ToString(r("Rod_mm")))
            ri.Stroke = SafeDbl(Convert.ToString(r("Stroke_mm")))
            ri.Cycles = SafeDbl(Convert.ToString(r("Cycles_per_min")))
            ri.Qty = SafeInt(Convert.ToString(r("Qty")))
            If ri.Qty <= 0 Then ri.Qty = 1
            ri.Pressure = SafeDbl(Convert.ToString(r("Pressure_bar_g")))
            ri.NLcycle = SafeDbl(Convert.ToString(r("NL_per_cycle")))
            ri.NLmin = SafeDbl(Convert.ToString(r("NL_per_min")))
            list.Add(ri)
        Next
        Return list
    End Function

    Private Function BuildReportLines() As List(Of String)
        Dim lines As New List(Of String)()

        Dim rows As List(Of RowInfo) = GetRowsSnapshot()
        Dim included As New List(Of RowInfo)()
        Dim totalNlMin As Double = 0.0

        For Each r As RowInfo In rows
            If r.Include Then
                included.Add(r)
                totalNlMin += r.NLmin
            End If
        Next

        Dim safety As Double = 1.25
        Dim recNlMin As Double = totalNlMin * safety
        Dim rec_m3h As Double = (recNlMin / 1000.0) * 60.0

        ' ---- HEADER
        lines.Add("REPORT: Air Consumption & Compressor Sizing")
        lines.Add("Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        lines.Add("")

        ' ---- DECISION FIRST
        lines.Add("RECOMMENDED COMPRESSOR (Free Air Delivery)")
        lines.Add("----------------------------------------")
        lines.Add("Total included demand: " & totalNlMin.ToString("0.##", Inv) & " NL/min")
        lines.Add("Recommended with margin (x" & safety.ToString("0.##", Inv) & "): " & recNlMin.ToString("0.##", Inv) & " NL/min")
        lines.Add("Approx: " & rec_m3h.ToString("0.##", Inv) & " m^3/h (free air)")
        lines.Add("")

        ' ---- GLOSSARY
        lines.Add("GLOSSARY")
        lines.Add("--------")
        lines.Add("DA = Double Acting (air used for extend + retract)")
        lines.Add("SA = Single Acting (air used one direction, spring return)")
        lines.Add("NL/min = Normal Litres per minute (standard air volume per minute)")
        lines.Add("bar g  = gauge pressure (Pabs = bar g + 1)")
        lines.Add("")

        ' ---- TOP CONSUMERS
        lines.Add("TOP AIR CONSUMERS (Included only)")
        lines.Add("--------------------------------")
        If included.Count = 0 Then
            lines.Add("No included rows.")
        Else
            included.Sort(Function(a, b) b.NLmin.CompareTo(a.NLmin))
            Dim topN As Integer = Math.Min(5, included.Count)
            For i As Integer = 0 To topN - 1
                Dim rr As RowInfo = included(i)
                Dim pct As Double = 0.0
                If totalNlMin > 0 Then pct = (rr.NLmin / totalNlMin) * 100.0
                lines.Add((i + 1).ToString() & ") " & rr.Name & "  |  " & rr.NLmin.ToString("0.##", Inv) & " NL/min  |  " & pct.ToString("0.#", Inv) & "% of total")
            Next
        End If
        lines.Add("")

        ' ---- TABLE A: INPUTS
        lines.Add("TABLE A: CYLINDER INPUTS")
        lines.Add("------------------------")
        lines.Add("Inc | Name                 | Act | Bore | Rod | Stroke | Cyc/min | Qty | P(bar g)")
        lines.Add("----+----------------------+-----+------+------+--------+---------+-----+--------")
        For Each r As RowInfo In rows
            lines.Add( _
                (If(r.Include, "Y", "N")).PadRight(3) & " | " &
                TruncPad(r.Name, 20) & " | " &
                TruncPad(r.Action, 3) & " | " &
                r.Bore.ToString("0.#", Inv).PadLeft(4) & " | " &
                r.Rod.ToString("0.#", Inv).PadLeft(4) & " | " &
                r.Stroke.ToString("0.#", Inv).PadLeft(6) & " | " &
                r.Cycles.ToString("0.#", Inv).PadLeft(7) & " | " &
                r.Qty.ToString().PadLeft(3) & " | " &
                r.Pressure.ToString("0.##", Inv).PadLeft(6) _
            )
        Next
        lines.Add("")

        ' ---- TABLE B: RESULTS
        lines.Add("TABLE B: RESULTS")
        lines.Add("----------------")
        lines.Add("Name                 | NL/cycle | NL/min  | Included")
        lines.Add("---------------------+----------+---------+---------")
        For Each r As RowInfo In rows
            lines.Add( _
                TruncPad(r.Name, 20) & " | " &
                r.NLcycle.ToString("0.##", Inv).PadLeft(8) & " | " &
                r.NLmin.ToString("0.##", Inv).PadLeft(7) & " | " &
                If(r.Include, "YES", "NO") _
            )
        Next
        lines.Add("")

        ' ---- APPENDIX: METHOD
        lines.Add("APPENDIX: CALCULATION METHOD (Summary)")
        lines.Add("--------------------------------------")
        lines.Add("1) Abore = pi * (Bore^2) / 4")
        lines.Add("2) Arod  = pi * (Rod^2)  / 4")
        lines.Add("3) Vext  = Abore * Stroke")
        lines.Add("4) Vret  = (Abore - Arod) * Stroke (DA only)")
        lines.Add("5) Convert to litres: L = m^3 * 1000")
        lines.Add("6) Pabs = P(bar g) + 1")
        lines.Add("7) NL = L * Pabs")
        lines.Add("8) DA: NL/cycle(one) = (Lext + Lret) * Pabs")
        lines.Add("   SA: NL/cycle(one) =  Lext        * Pabs")
        lines.Add("9) Multiply by Qty and Cycles/min")
        lines.Add("")

        ' ---- NEW: BREAKDOWN FOR ALL INCLUDED CYLINDERS
        lines.Add("APPENDIX: PER-CYLINDER BREAKDOWN (Included only)")
        lines.Add("------------------------------------------------")
        If included.Count = 0 Then
            lines.Add("No included rows.")
        Else
            ' Keep the same order as Top Consumers (largest first)
            included.Sort(Function(a, b) b.NLmin.CompareTo(a.NLmin))

            For Each rr As RowInfo In included
                AppendPerCylinderBreakdown(lines, rr)
                lines.Add("") ' spacing between cylinders
            Next
        End If
        lines.Add("")

        Return lines
    End Function

    Private Sub AppendPerCylinderBreakdown(lines As List(Of String), rr As RowInfo)
        If lines Is Nothing OrElse rr Is Nothing Then Exit Sub

        Dim name As String = If(rr.Name, "")
        Dim action As String = If(rr.Action, "DA").Trim().ToUpperInvariant()
        If action <> "DA" AndAlso action <> "SA" Then action = "DA"

        Dim boreMm As Double = rr.Bore
        Dim rodMm As Double = rr.Rod
        Dim strokeMm As Double = rr.Stroke
        Dim cycles As Double = rr.Cycles
        Dim qty As Integer = rr.Qty
        If qty <= 0 Then qty = 1

        Dim pBarG As Double = rr.Pressure
        If pBarG < 0 Then pBarG = 0
        Dim pAbs As Double = pBarG + 1.0

        Dim boreM As Double = boreMm / 1000.0
        Dim rodM As Double = rodMm / 1000.0
        Dim strokeM As Double = strokeMm / 1000.0

        Dim areaBore As Double = Math.PI * (boreM * boreM) / 4.0
        Dim areaRod As Double = Math.PI * (rodM * rodM) / 4.0
        Dim areaAnn As Double = areaBore - areaRod
        If areaAnn < 0 Then areaAnn = 0

        Dim vExt_m3 As Double = areaBore * strokeM
        Dim vRet_m3 As Double = areaAnn * strokeM
        Dim vExt_L As Double = vExt_m3 * 1000.0
        Dim vRet_L As Double = vRet_m3 * 1000.0

        Dim nlCycleOne As Double
        If action = "DA" Then
            nlCycleOne = (vExt_L + vRet_L) * pAbs
        Else
            nlCycleOne = vExt_L * pAbs
        End If

        Dim nlCycleAll As Double = nlCycleOne * CDbl(qty)
        Dim nlMinAll As Double = nlCycleAll * cycles

        lines.Add("Cylinder: " & name)
        lines.Add("  Action: " & action & "   Pabs = (Pbar(g)+1) = " & pAbs.ToString("0.###", Inv))
        lines.Add("  Bore=" & boreMm.ToString("0.###", Inv) & "mm  Rod=" & rodMm.ToString("0.###", Inv) & "mm  Stroke=" & strokeMm.ToString("0.###", Inv) & "mm")
        lines.Add("  Abore = pi*(Bore^2)/4 = " & areaBore.ToString("0.000000", Inv) & " m^2")
        lines.Add("  Arod  = pi*(Rod^2)/4  = " & areaRod.ToString("0.000000", Inv) & " m^2")
        lines.Add("  Vext  = Abore*Stroke  = " & vExt_m3.ToString("0.000000", Inv) & " m^3 = " & vExt_L.ToString("0.###", Inv) & " L")
        If action = "DA" Then
            lines.Add("  Vret  = (Abore-Arod)*Stroke = " & vRet_m3.ToString("0.000000", Inv) & " m^3 = " & vRet_L.ToString("0.###", Inv) & " L")
        End If
        lines.Add("  NL/cycle (one cyl) = Volume(L) * Pabs = " & nlCycleOne.ToString("0.###", Inv))
        lines.Add("  Qty=" & qty.ToString() & " => NL/cycle(all) = " & nlCycleAll.ToString("0.###", Inv))
        lines.Add("  Cycles/min=" & cycles.ToString("0.###", Inv) & " => NL/min = " & nlMinAll.ToString("0.###", Inv))
    End Sub

    Private Function TruncPad(s As String, width As Integer) As String
        Dim t As String = If(s, "")
        t = t.Replace(vbTab, " ").Trim()
        If t.Length > width Then
            t = t.Substring(0, width - 1) & "…"
        End If
        Return t.PadRight(width)
    End Function

    ' =========================================================
    ' EXPORT / COPY
    ' =========================================================
    Private Sub OnCopyText(sender As Object, e As EventArgs)
        Try
            Dim report As String = String.Join(Environment.NewLine, BuildReportLines().ToArray())
            Clipboard.SetText(report)
            MessageBox.Show("Report copied to clipboard.", "Copy", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Copy failed: " & ex.Message, "Copy", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub OnExportPdf(sender As Object, e As EventArgs)
        Try
            If dt.Rows.Count = 0 Then
                MessageBox.Show("No rows to export.", "Export PDF", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Using sfd As New SaveFileDialog()
                sfd.Title = "Export PDF Report"
                sfd.Filter = "PDF (*.pdf)|*.pdf"
                sfd.FileName = "AirConsumption_Report_" & DateTime.Now.ToString("yyyyMMdd_HHmm") & ".pdf"
                If sfd.ShowDialog() <> DialogResult.OK Then Return

                Dim lines As List(Of String) = BuildReportLines()
                SimplePdfWriter.WriteTextPdfMultiPage(sfd.FileName, "Air Consumption & Compressor Sizing", lines)

                MessageBox.Show("PDF exported successfully:" & vbCrLf & sfd.FileName, "Export PDF", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using

        Catch ex As Exception
            MessageBox.Show("PDF export failed: " & ex.Message, "Export PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' =========================================================
    ' SIMPLE MULTI-PAGE PDF WRITER (no dependencies)
    ' =========================================================
    Private NotInheritable Class SimplePdfWriter
        Private Sub New()
        End Sub

        Public Shared Sub WriteTextPdfMultiPage(path As String, title As String, lines As List(Of String))
            If lines Is Nothing Then lines = New List(Of String)()

            ' PDF page settings (A4 portrait)
            Dim pageW As Integer = 595
            Dim pageH As Integer = 842

            Dim marginL As Integer = 50
            Dim marginT As Integer = 34
            Dim marginB As Integer = 50

            Dim titleY As Integer = pageH - marginT
            Dim bodyStartY As Integer = titleY - 24

            Dim fontTitle As Integer = 14
            Dim fontBody As Integer = 9
            Dim lineStep As Integer = 11

            ' Lines per page
            Dim usableH As Integer = (bodyStartY - marginB)
            Dim maxLinesPerPage As Integer = Math.Max(10, usableH \ lineStep)

            ' Wrap long lines
            Dim wrapped As List(Of String) = WrapLines(lines, 105)

            ' Split into pages
            Dim pages As New List(Of List(Of String))()
            Dim i As Integer = 0
            While i < wrapped.Count
                Dim p As New List(Of String)()
                Dim takeN As Integer = Math.Min(maxLinesPerPage, wrapped.Count - i)
                For k As Integer = 0 To takeN - 1
                    p.Add(wrapped(i + k))
                Next
                pages.Add(p)
                i += takeN
            End While

            If pages.Count = 0 Then
                pages.Add(New List(Of String)() From {""})
            End If

            Dim pageCount As Integer = pages.Count
            Dim fontObj As Integer = 3 + 2 * pageCount

            Dim sb As New StringBuilder()
            sb.AppendLine("%PDF-1.4")

            Dim xref As New List(Of Integer)()

            ' 1) Catalog
            xref.Add(sb.Length)
            sb.AppendLine("1 0 obj")
            sb.AppendLine("<< /Type /Catalog /Pages 2 0 R >>")
            sb.AppendLine("endobj")

            ' 2) Pages root
            xref.Add(sb.Length)
            sb.AppendLine("2 0 obj")
            Dim kids As New StringBuilder()
            kids.Append("[")
            For p As Integer = 0 To pageCount - 1
                Dim pageObj As Integer = 3 + 2 * p
                kids.Append(pageObj.ToString() & " 0 R ")
            Next
            kids.Append("]")
            sb.AppendLine("<< /Type /Pages /Kids " & kids.ToString() & " /Count " & pageCount.ToString() & " >>")
            sb.AppendLine("endobj")

            ' Pages + contents
            For p As Integer = 0 To pageCount - 1
                Dim pageObj As Integer = 3 + 2 * p
                Dim contentObj As Integer = 4 + 2 * p

                xref.Add(sb.Length)
                sb.AppendLine(pageObj.ToString() & " 0 obj")
                sb.AppendLine("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " & pageW.ToString() & " " & pageH.ToString() & "] " &
                              "/Contents " & contentObj.ToString() & " 0 R " &
                              "/Resources << /Font << /F1 " & fontObj.ToString() & " 0 R >> >> >>")
                sb.AppendLine("endobj")

                Dim content As String = BuildPageContent(title, pages(p), p + 1, pageCount, marginL, titleY, bodyStartY, fontTitle, fontBody, lineStep)
                Dim contentBytes() As Byte = Encoding.ASCII.GetBytes(content)

                xref.Add(sb.Length)
                sb.AppendLine(contentObj.ToString() & " 0 obj")
                sb.AppendLine("<< /Length " & contentBytes.Length.ToString() & " >>")
                sb.AppendLine("stream")
                sb.Append(content)
                sb.AppendLine()
                sb.AppendLine("endstream")
                sb.AppendLine("endobj")
            Next

            ' Font object
            xref.Add(sb.Length)
            sb.AppendLine(fontObj.ToString() & " 0 obj")
            sb.AppendLine("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
            sb.AppendLine("endobj")

            ' XRef
            Dim xrefStart As Integer = sb.Length
            sb.AppendLine("xref")
            sb.AppendLine("0 " & (xref.Count + 1).ToString())
            sb.AppendLine("0000000000 65535 f ")
            For Each off As Integer In xref
                sb.AppendLine(off.ToString("0000000000") & " 00000 n ")
            Next

            sb.AppendLine("trailer")
            sb.AppendLine("<< /Size " & (xref.Count + 1).ToString() & " /Root 1 0 R >>")
            sb.AppendLine("startxref")
            sb.AppendLine(xrefStart.ToString())
            sb.AppendLine("%%EOF")

            File.WriteAllText(path, sb.ToString(), Encoding.ASCII)
        End Sub

        Private Shared Function BuildPageContent(title As String, pageLines As List(Of String), pageNum As Integer, pageCount As Integer,
                                                x As Integer, yTitle As Integer, yBody As Integer,
                                                fTitle As Integer, fBody As Integer, stepY As Integer) As String
            Dim sb As New StringBuilder()

            Dim safeTitle As String = PdfEscape(SanitizeAscii(title))
            Dim header As String = safeTitle & "  (Page " & pageNum.ToString() & " of " & pageCount.ToString() & ")"
            header = PdfEscape(SanitizeAscii(header))

            sb.AppendLine("BT")
            sb.AppendLine("/F1 " & fTitle.ToString() & " Tf")
            sb.AppendLine(x.ToString() & " " & yTitle.ToString() & " Td")
            sb.AppendLine("(" & header & ") Tj")
            sb.AppendLine("ET")

            sb.AppendLine("BT")
            sb.AppendLine("/F1 " & fBody.ToString() & " Tf")
            sb.AppendLine(x.ToString() & " " & yBody.ToString() & " Td")

            If pageLines Is Nothing Then pageLines = New List(Of String)()
            For Each raw As String In pageLines
                Dim t As String = If(raw, "")
                t = SanitizeAscii(t)
                sb.AppendLine("(" & PdfEscape(t) & ") Tj")
                sb.AppendLine("0 -" & stepY.ToString() & " Td")
            Next

            sb.AppendLine("ET")
            Return sb.ToString()
        End Function

        Private Shared Function SanitizeAscii(s As String) As String
            If s Is Nothing Then Return ""
            Dim sb As New StringBuilder()
            For i As Integer = 0 To s.Length - 1
                Dim ch As Char = s(i)
                Dim code As Integer = AscW(ch)
                If code = 9 Then
                    sb.Append("    ")
                ElseIf code >= 32 AndAlso code <= 126 Then
                    sb.Append(ch)
                Else
                    sb.Append("-")
                End If
            Next
            Return sb.ToString()
        End Function

        Private Shared Function PdfEscape(s As String) As String
            If s Is Nothing Then Return ""
            Dim t As String = s
            t = t.Replace("\", "\\")
            t = t.Replace("(", "\(")
            t = t.Replace(")", "\)")
            t = t.Replace(ChrW(13), " ")
            t = t.Replace(ChrW(10), " ")
            Return t
        End Function

        Private Shared Function WrapLines(lines As List(Of String), maxLen As Integer) As List(Of String)
            Dim out As New List(Of String)()
            If lines Is Nothing Then Return out
            For Each l As String In lines
                Dim t As String = If(l, "")
                If t = "" Then
                    out.Add("")
                Else
                    While t.Length > maxLen
                        out.Add(t.Substring(0, maxLen))
                        t = t.Substring(maxLen)
                    End While
                    out.Add(t)
                End If
            Next
            Return out
        End Function
    End Class

End Class
