Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class AdminLoginForm
    Inherits Form

    Public Property LoginSuccessful As Boolean = False

    Private txtUser As TextBox
    Private txtPass As TextBox

    Public Sub New()

        Me.Text = "Admin Login"
        Me.Size = New Size(360, 220)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(24, 30, 40)
        Me.ForeColor = Color.White
        Me.Font = New Font("Segoe UI", 10)

        Dim lblTitle As New Label()
        lblTitle.Text = "Administrator Login"
        lblTitle.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(80, 15)
        Me.Controls.Add(lblTitle)

        Dim lblUser As New Label()
        lblUser.Text = "Username"
        lblUser.Location = New Point(40, 60)
        lblUser.AutoSize = True
        Me.Controls.Add(lblUser)

        txtUser = New TextBox()
        txtUser.Location = New Point(140, 56)
        txtUser.Width = 160
        Me.Controls.Add(txtUser)

        Dim lblPass As New Label()
        lblPass.Text = "Password"
        lblPass.Location = New Point(40, 95)
        lblPass.AutoSize = True
        Me.Controls.Add(lblPass)

        txtPass = New TextBox()
        txtPass.Location = New Point(140, 91)
        txtPass.Width = 160
        txtPass.PasswordChar = "*"c
        Me.Controls.Add(txtPass)

        Dim btnLogin As New Button()
        btnLogin.Text = "Login"
        btnLogin.Size = New Size(90, 32)
        btnLogin.Location = New Point(140, 135)
        AddHandler btnLogin.Click, AddressOf DoLogin
        Me.Controls.Add(btnLogin)

        Dim btnCancel As New Button()
        btnCancel.Text = "Cancel"
        btnCancel.Size = New Size(90, 32)
        btnCancel.Location = New Point(240, 135)
        AddHandler btnCancel.Click, Sub() Me.Close()
        Me.Controls.Add(btnCancel)

        Me.AcceptButton = btnLogin
    End Sub

    Private Sub DoLogin(sender As Object, e As EventArgs)

        If txtUser.Text = "admin" AndAlso txtPass.Text = "MetaMech@2026" Then
            LoginSuccessful = True
            Me.Close()
        Else
            MessageBox.Show(
                "Invalid admin credentials.",
                "Access Denied",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End If

    End Sub

End Class
