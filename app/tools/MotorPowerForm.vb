Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class MotorPowerForm
    Inherits Form

    Public Sub New()
        Me.Text = "Motor Power Calculator"
        Me.Size = New Size(420, 490)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.FromArgb(24, 30, 40)
        Me.ForeColor = Color.White
        Me.Font = New Font("Segoe UI", 9)

        FormHeader.AddPremiumHeader(Me, "Motor Power Calculator", "MetaMech Engineering Tools")

        Dim lbl As New Label()
        lbl.Text = "Motor Power (kW) Calculator" & vbCrLf &
                   "Estimates required motor power." & vbCrLf & vbCrLf &
                   "Status: Coming soon"
        lbl.AutoSize = True
        lbl.Location = New Point(20, 105)
        Me.Controls.Add(lbl)

        Dim btn As New Button()
        btn.Text = "Close"
        btn.Size = New Size(100, 30)
        btn.Location = New Point(150, 255)
        AddHandler btn.Click, Sub() Me.Close()
        Me.Controls.Add(btn)
    End Sub

End Class
