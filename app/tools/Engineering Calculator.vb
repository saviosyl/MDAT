Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class CalcForm
    Inherits Form

    Private txtA As TextBox
    Private txtB As TextBox
    Private lblResult As Label

    Public Sub New()

        Me.Text = "Engineering Calculator"
        Me.Size = New Size(300, 330)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent

        FormFooter.AddPremiumFooter(Me)
        FormHeader.AddPremiumHeader(Me, "Engineering Calculator", "MetaMech Engineering Tools")

        txtA = New TextBox()
        txtA.Location = New Point(30, 115)
        Me.Controls.Add(txtA)

        txtB = New TextBox()
        txtB.Location = New Point(30, 145)
        Me.Controls.Add(txtB)

        Dim btnAdd As New Button()
        btnAdd.Text = "+"
        btnAdd.Location = New Point(180, 115)
        AddHandler btnAdd.Click, Sub() Calc(Function(a, b) a + b)
        Me.Controls.Add(btnAdd)

        Dim btnMul As New Button()
        btnMul.Text = "×"
        btnMul.Location = New Point(180, 145)
        AddHandler btnMul.Click, Sub() Calc(Function(a, b) a * b)
        Me.Controls.Add(btnMul)

        lblResult = New Label()
        lblResult.Location = New Point(30, 195)
        lblResult.AutoSize = True
        Me.Controls.Add(lblResult)

    End Sub

    Private Sub Calc(f As Func(Of Double, Double, Double))
        Dim a, b As Double
        If Double.TryParse(txtA.Text, a) AndAlso Double.TryParse(txtB.Text, b) Then
            lblResult.Text = "Result: " & f(a, b).ToString()
        End If
    End Sub

End Class
