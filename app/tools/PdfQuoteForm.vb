Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class PdfQuoteForm
    Inherits Form

    Public Sub New()
        Me.Text = "Export Engineering Report (PDF)"
        Me.Size = New Size(420, 370)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.FromArgb(24, 30, 40)
        Me.ForeColor = Color.White
        Me.Font = New Font("Segoe UI", 9)

        FormFooter.AddPremiumFooter(Me)
        FormHeader.AddPremiumHeader(Me, "PDF Quote Generator", "MetaMech Engineering Tools")

        Dim lbl As New Label()
        lbl.Text = "Engineering Report Export" & vbCrLf &
                   "This will export calculations to PDF." & vbCrLf & vbCrLf &
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
