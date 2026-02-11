Option Strict On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class UnitConverterForm
    Inherits Form

    Private cmbCategory As ComboBox
    Private cmbFrom As ComboBox
    Private cmbTo As ComboBox
    Private txtInput As TextBox
    Private txtOutput As TextBox
    Private lblStatus As Label

    Public Sub New()
        Me.Text = "Engineering Unit Converter"
        Me.Size = New Size(480, 420)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.FromArgb(20, 24, 32)
        FormFooter.AddPremiumFooter(Me)
        BuildUI()
        FormHeader.AddPremiumHeader(Me, "Unit Converter", "MetaMech Engineering Tools")
    End Sub

    Private Sub BuildUI()

        ' ===== CATEGORY =====
        Dim lblCat As New Label()
        lblCat.Text = "Category"
        lblCat.ForeColor = Color.LightGray
        lblCat.Location = New Point(20, 10)
        Me.Controls.Add(lblCat)

        cmbCategory = New ComboBox()
        cmbCategory.Items.AddRange(New String() {"Length", "Mass", "Force", "Power"})
        cmbCategory.DropDownStyle = ComboBoxStyle.DropDownList
        cmbCategory.Location = New Point(20, 30)
        cmbCategory.Width = 200
        AddHandler cmbCategory.SelectedIndexChanged, AddressOf UpdateUnits
        Me.Controls.Add(cmbCategory)

        ' ===== FROM =====
        Dim lblFrom As New Label()
        lblFrom.Text = "From"
        lblFrom.ForeColor = Color.LightGray
        lblFrom.Location = New Point(20, 70)
        Me.Controls.Add(lblFrom)

        cmbFrom = New ComboBox()
        cmbFrom.DropDownStyle = ComboBoxStyle.DropDownList
        cmbFrom.Location = New Point(20, 90)
        cmbFrom.Width = 200
        Me.Controls.Add(cmbFrom)

        ' ===== TO =====
        Dim lblTo As New Label()
        lblTo.Text = "To"
        lblTo.ForeColor = Color.LightGray
        lblTo.Location = New Point(240, 70)
        Me.Controls.Add(lblTo)

        cmbTo = New ComboBox()
        cmbTo.DropDownStyle = ComboBoxStyle.DropDownList
        cmbTo.Location = New Point(240, 90)
        cmbTo.Width = 200
        Me.Controls.Add(cmbTo)

        ' ===== INPUT =====
        Dim lblIn As New Label()
        lblIn.Text = "Input"
        lblIn.ForeColor = Color.LightGray
        lblIn.Location = New Point(20, 130)
        Me.Controls.Add(lblIn)

        txtInput = New TextBox()
        txtInput.Location = New Point(20, 150)
        txtInput.Width = 200
        Me.Controls.Add(txtInput)

        ' ===== OUTPUT =====
        Dim lblOut As New Label()
        lblOut.Text = "Result"
        lblOut.ForeColor = Color.LightGray
        lblOut.Location = New Point(240, 130)
        Me.Controls.Add(lblOut)

        txtOutput = New TextBox()
        txtOutput.Location = New Point(240, 150)
        txtOutput.Width = 200
        txtOutput.ReadOnly = True
        Me.Controls.Add(txtOutput)

        ' ===== BUTTON =====
        Dim btnConvert As New Button()
        btnConvert.Text = "Convert"
        btnConvert.Location = New Point(20, 190)
        btnConvert.Width = 420
        btnConvert.Height = 32
        btnConvert.BackColor = Color.FromArgb(0, 122, 204)
        btnConvert.ForeColor = Color.White
        btnConvert.FlatStyle = FlatStyle.Flat
        AddHandler btnConvert.Click, AddressOf ConvertValue
        Me.Controls.Add(btnConvert)

        ' ===== STATUS =====
        lblStatus = New Label()
        lblStatus.Text = "Ready"
        lblStatus.ForeColor = Color.LightGray
        lblStatus.Location = New Point(20, 230)
        lblStatus.AutoSize = True
        Me.Controls.Add(lblStatus)

        cmbCategory.SelectedIndex = 0

    End Sub

    ' ================= UPDATE UNIT LIST =================
    Private Sub UpdateUnits(sender As Object, e As EventArgs)

        cmbFrom.Items.Clear()
        cmbTo.Items.Clear()

        Select Case cmbCategory.Text
            Case "Length"
                cmbFrom.Items.AddRange(New String() {"mm", "inch", "m"})
            Case "Mass"
                cmbFrom.Items.AddRange(New String() {"kg", "lb"})
            Case "Force"
                cmbFrom.Items.AddRange(New String() {"N", "kgf"})
            Case "Power"
                cmbFrom.Items.AddRange(New String() {"kW", "HP"})
        End Select

        For Each item In cmbFrom.Items
            cmbTo.Items.Add(item)
        Next

        If cmbFrom.Items.Count > 0 Then
            cmbFrom.SelectedIndex = 0
            cmbTo.SelectedIndex = 1
        End If

    End Sub

    ' ================= CONVERT =================
    Private Sub ConvertValue(sender As Object, e As EventArgs)

        Dim inputVal As Double

        If Not Double.TryParse(txtInput.Text, inputVal) Then
            MessageBox.Show("Invalid number", "Unit Converter")
            Exit Sub
        End If

        Dim result As Double = inputVal
        Dim f As String = cmbFrom.Text
        Dim t As String = cmbTo.Text

        ' --- Length ---
        If f = "mm" And t = "inch" Then result = inputVal / 25.4
        If f = "inch" And t = "mm" Then result = inputVal * 25.4
        If f = "m" And t = "mm" Then result = inputVal * 1000
        If f = "mm" And t = "m" Then result = inputVal / 1000

        ' --- Mass ---
        If f = "kg" And t = "lb" Then result = inputVal * 2.20462
        If f = "lb" And t = "kg" Then result = inputVal / 2.20462

        ' --- Force ---
        If f = "N" And t = "kgf" Then result = inputVal / 9.80665
        If f = "kgf" And t = "N" Then result = inputVal * 9.80665

        ' --- Power ---
        If f = "kW" And t = "HP" Then result = inputVal * 1.34102
        If f = "HP" And t = "kW" Then result = inputVal / 1.34102

        txtOutput.Text = Math.Round(result, 4).ToString()
        lblStatus.Text = "Converted"

    End Sub

End Class
