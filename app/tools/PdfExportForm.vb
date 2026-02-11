Imports System
Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms

Public Class PdfExportForm
    Inherits Form

    Private txtContent As TextBox

    Public Sub New()
        Me.Text = "Export Report"
        Me.Size = New Size(420, 520)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent

        FormHeader.AddPremiumHeader(Me, "PDF Export", "MetaMech Engineering Tools")

        txtContent = New TextBox()
        txtContent.Multiline = True
        txtContent.Dock = DockStyle.Top
        txtContent.Height = 180
        txtContent.Text = "MetaMech Report" & vbCrLf & DateTime.Now.ToString()
        Me.Controls.Add(txtContent)

        Dim btn As New Button()
        btn.Text = "Export"
        btn.Dock = DockStyle.Bottom
        AddHandler btn.Click, AddressOf Export
        Me.Controls.Add(btn)
    End Sub

    Private Sub Export(sender As Object, e As EventArgs)
        Dim filePath As String = IO.Path.Combine(
            Application.StartupPath,
            "Reports",
            "Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
        )

        Directory.CreateDirectory(IO.Path.GetDirectoryName(filePath))
        File.WriteAllText(filePath, txtContent.Text)

        MessageBox.Show(
            "Report exported:" & vbCrLf & filePath,
            "Export Complete",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
    End Sub

End Class
