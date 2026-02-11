Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class TorqueCalculatorForm
    Inherits Form

    Public Sub New()
        Me.Text = "Torque Calculator"
        Me.Size = New Size(420, 260)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.FromArgb(24, 30, 40)
        Me.ForeColor = Color.White
        Me.Font = New Font("Segoe UI", 9)

        Dim lbl As New Label()
        lbl.Text = "Torque Requirement Calculator" & vbCrLf &
                   "Calculates torque based on load and radius." & vbCrLf & vbCrLf &
                   "Status: Coming soon"
        lbl.AutoSize = True
        lbl.Location = New Point(20, 20)
        Me.Controls.Add(lbl)

        Dim btn As New Button()
        btn.Text = "Close"
        btn.Size = New Size(100, 30)
        btn.Location = New Point(150, 170)
        AddHandler btn.Click, Sub() Me.Close()
        Me.Controls.Add(btn)
    End Sub

End Class
