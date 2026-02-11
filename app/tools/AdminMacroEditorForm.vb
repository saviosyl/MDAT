Imports System
Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Xml

Public Class AdminMacroEditorForm
    Inherits Form

    Private dgv As DataGridView
    Private btnSave As Button
    Private btnClose As Button

    Private Const CONFIG_FILE As String = "admin_macros.xml"

    Public Sub New()

        Me.Text = "Admin Macro Editor"
        Me.Size = New Size(1020, 420)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.FromArgb(24, 30, 40)
        Me.ForeColor = Color.White
        Me.Font = New Font("Segoe UI", 9)

        dgv = New DataGridView()
        dgv.Dock = DockStyle.Top
        dgv.Height = 300
        dgv.AllowUserToAddRows = False
        dgv.RowHeadersVisible = False
        dgv.SelectionMode = DataGridViewSelectionMode.CellSelect
        dgv.EnableHeadersVisualStyles = False

        dgv.BackgroundColor = Color.FromArgb(30, 36, 48)
        dgv.GridColor = Color.DarkGray

        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 170, 210)
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)

        dgv.DefaultCellStyle.BackColor = Color.White
        dgv.DefaultCellStyle.ForeColor = Color.Black
        dgv.DefaultCellStyle.SelectionBackColor = Color.LightSteelBlue
        dgv.DefaultCellStyle.SelectionForeColor = Color.Black

        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245)

        dgv.Columns.Add("Slot", "Slot")
        dgv.Columns.Add("Name", "Button Name")
        dgv.Columns.Add("Tooltip", "Tooltip")
        dgv.Columns.Add("File", "Macro File (.swp)")
        dgv.Columns.Add("Module", "Module")
        dgv.Columns.Add("Method", "Method")

        Dim browseCol As New DataGridViewButtonColumn()
        browseCol.HeaderText = ""
        browseCol.Text = "Browse..."
        browseCol.UseColumnTextForButtonValue = True
        dgv.Columns.Add(browseCol)

        dgv.Columns(0).Width = 50
        dgv.Columns(0).ReadOnly = True
        dgv.Columns(0).DefaultCellStyle.BackColor = Color.Gainsboro
        dgv.Columns(0).DefaultCellStyle.ForeColor = Color.Black

        dgv.Columns(1).Width = 150
        dgv.Columns(2).Width = 200
        dgv.Columns(3).Width = 220
        dgv.Columns(4).Width = 100
        dgv.Columns(5).Width = 100
        dgv.Columns(6).Width = 90

        AddHandler dgv.CellContentClick, AddressOf BrowseMacroFile

        Me.Controls.Add(dgv)

        btnSave = New Button()
        btnSave.Text = "Save"
        btnSave.Size = New Size(120, 34)
        btnSave.Location = New Point(620, 320)
        AddHandler btnSave.Click, AddressOf SaveConfig
        Me.Controls.Add(btnSave)

        btnClose = New Button()
        btnClose.Text = "Close"
        btnClose.Size = New Size(120, 34)
        btnClose.Location = New Point(760, 320)
        AddHandler btnClose.Click, Sub() Me.Close()
        Me.Controls.Add(btnClose)

        LoadConfig()
    End Sub

    '====================================================
    ' LOAD XML
    '====================================================
    Private Sub LoadConfig()

        dgv.Rows.Clear()

        If Not File.Exists(CONFIG_FILE) Then
            For i As Integer = 1 To 10
                dgv.Rows.Add(i, "Macro " & i, "", "", "", "main", "Browse...")
            Next
            Exit Sub
        End If

        Dim xd As New XmlDocument()
        xd.Load(CONFIG_FILE)

        For Each n As XmlNode In xd.SelectNodes("//Macro")
            dgv.Rows.Add(
                n.Attributes("slot").Value,
                n("Name").InnerText,
                n("Tooltip").InnerText,
                n("File").InnerText,
                If(n("Module") Is Nothing, "", n("Module").InnerText),
                If(n("Method") Is Nothing, "main", n("Method").InnerText),
                "Browse..."
            )
        Next

    End Sub

    '====================================================
    ' BROWSE MACRO
    '====================================================
    Private Sub BrowseMacroFile(sender As Object, e As DataGridViewCellEventArgs)

        If e.RowIndex < 0 OrElse e.ColumnIndex <> 6 Then Exit Sub

        Dim ofd As New OpenFileDialog()
        ofd.Filter = "SolidWorks Macro (*.swp)|*.swp"

        If ofd.ShowDialog() = DialogResult.OK Then
            dgv.Rows(e.RowIndex).Cells(3).Value = ofd.FileName
        End If

    End Sub

    '====================================================
    ' SAVE XML (FIXED)
    '====================================================
    Private Sub SaveConfig(sender As Object, e As EventArgs)

        Dim xd As New XmlDocument()
        Dim root As XmlElement = xd.CreateElement("Macros")
        xd.AppendChild(root)

        For Each r As DataGridViewRow In dgv.Rows

            Dim nameVal As String = SafeCell(r, 1)
            Dim tooltipVal As String = SafeCell(r, 2)
            Dim fileVal As String = SafeCell(r, 3)
            Dim moduleVal As String = SafeCell(r, 4)   ' MAY BE EMPTY
            Dim methodVal As String = SafeCell(r, 5)   ' REQUIRED

If nameVal = "" OrElse methodVal = "" Then
    MessageBox.Show("Name and Method cannot be empty.", "Validation Error")
    Exit Sub
End If



            Dim macroElem As XmlElement = xd.CreateElement("Macro")
            macroElem.SetAttribute("slot", r.Cells(0).Value.ToString())

            macroElem.AppendChild(MakeElem(xd, "Name", nameVal))
            macroElem.AppendChild(MakeElem(xd, "Tooltip", tooltipVal))
            macroElem.AppendChild(MakeElem(xd, "File", fileVal))
            macroElem.AppendChild(MakeElem(xd, "Module", moduleVal))
            macroElem.AppendChild(MakeElem(xd, "Method", methodVal))

            root.AppendChild(macroElem)

        Next

        xd.Save(CONFIG_FILE)
        MessageBox.Show("Macro configuration saved successfully.", "Saved")

    End Sub

    Private Function SafeCell(r As DataGridViewRow, idx As Integer) As String
        If r.Cells(idx).Value Is Nothing Then Return ""
        Return r.Cells(idx).Value.ToString().Trim()
    End Function

    Private Function MakeElem(xd As XmlDocument, name As String, value As String) As XmlElement
        Dim e As XmlElement = xd.CreateElement(name)
        e.InnerText = value
        Return e
    End Function

End Class
