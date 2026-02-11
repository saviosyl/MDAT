Public Class UnitConverterForm
    Inherits Form

    Private txtMm As TextBox
    Private txtIn As TextBox

    Public Sub New()

        Me.Text = "Unit Converter"
        Me.Size = New Size(320, 200)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent

        txtMm = New TextBox()
        txtMm.Location = New Point(30, 30)
        Me.Controls.Add(txtMm)

        Dim btnToIn As New Button()
        btnToIn.Text = "mm → inch"
        btnToIn.Location = New Point(180, 30)
        AddHandler btnToIn.Click, AddressOf MmToIn
        Me.Controls.Add(btnToIn)

        txtIn = New TextBox()
        txtIn.Location = New Point(30, 70)
        Me.Controls.Add(txtIn)

        Dim btnToMm As New Button()
        btnToMm.Text = "inch → mm"
        btnToMm.Location = New Point(180, 70)
        AddHandler btnToMm.Click, AddressOf InToMm
        Me.Controls.Add(btnToMm)

    End Sub

    Private Sub MmToIn(sender As Object, e As EventArgs)
        Dim mm As Double
        If Double.TryParse(txtMm.Text, mm) Then
            txtIn.Text = (mm / 25.4).ToString("0.###")
        End If
    End Sub

    Private Sub InToMm(sender As Object, e As EventArgs)
        Dim inch As Double
        If Double.TryParse(txtIn.Text, inch) Then
            txtMm.Text = (inch * 25.4).ToString("0.###")
        End If
    End Sub

End Class
