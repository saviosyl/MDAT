Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Windows.Forms

Module Program

    Private Const LICENSE_FILE As String = "license.key"
    Private Const CONFIG_FILE As String = "config.txt"

    <STAThread()>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        ' Load config (for server URL and client token)
        LoadConfig()

        ' If no license.key exists, try to activate a trial automatically
        Dim licPath As String = Path.Combine(Application.StartupPath, LICENSE_FILE)
        If Not File.Exists(licPath) Then
            TryActivateTrial(licPath)
        End If

        Dim splash As New SplashForm()
        splash.Show()
        Application.DoEvents()

        Dim mainForm As New MainForm()
        splash.Close()
        splash.Dispose()

        Application.Run(mainForm)
    End Sub

    Private Sub TryActivateTrial(ByVal licPath As String)
        Try
            Dim machineId As String = MachineFingerprint.GetMachineId()
            Dim machineName As String = Environment.MachineName
            Dim userName As String = Environment.UserName

            Dim result As TrialActivateResult = SeatServerClient.TrialActivate(machineId, machineName, userName)

            If result.Ok AndAlso result.License IsNot Nothing AndAlso result.License.Trim().Length > 0 Then
                ' Save the signed license
                File.WriteAllText(licPath, result.License.Trim())

                MessageBox.Show(
                    "Welcome to MetaMech!" & vbCrLf & vbCrLf &
                    "Your 3-day free trial has been activated." & vbCrLf &
                    "License ID: " & result.LicenseId & vbCrLf & vbCrLf &
                    "To continue using MetaMech after the trial," & vbCrLf &
                    "visit https://metamechsolutions.com/pricing",
                    "MetaMech - Trial Activated",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)
            Else
                Dim code As String = If(result.Code, "").Trim().ToUpperInvariant()
                Dim msg As String = If(result.Message, "").Trim()

                If code = "TRIAL_USED" Then
                    MessageBox.Show(
                        "Free trial has already been used on this machine." & vbCrLf & vbCrLf &
                        "To continue using MetaMech, please purchase a license:" & vbCrLf &
                        "https://metamechsolutions.com/pricing" & vbCrLf & vbCrLf &
                        "If you believe this is an error, contact support:" & vbCrLf &
                        "https://metamechsolutions.com/contact",
                        "MetaMech - Trial Expired",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
                Else
                    ' Some other error — let them know but don't block startup
                    If msg.Length > 0 Then
                        MessageBox.Show(
                            "Could not activate trial: " & msg & vbCrLf & vbCrLf &
                            "You can still use MetaMech with a purchased license.",
                            "MetaMech",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
                    End If
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(
                "Trial activation error:" & vbCrLf & vbCrLf &
                ex.ToString() & vbCrLf & vbCrLf &
                "Server: " & SeatServerClient.ServerBaseUrl & vbCrLf &
                "Token set: " & (SeatServerClient.ClientToken.Length > 0).ToString(),
                "MetaMech - Debug",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub LoadConfig()
        Try
            Dim cfgPath As String = Path.Combine(Application.StartupPath, CONFIG_FILE)
            If Not File.Exists(cfgPath) Then Return

            Dim lines() As String = File.ReadAllLines(cfgPath)
            For Each line As String In lines
                If line Is Nothing Then Continue For
                Dim t As String = line.Trim()
                If t.Length = 0 OrElse t.StartsWith("#") Then Continue For

                Dim eq As Integer = t.IndexOf("="c)
                If eq <= 0 Then Continue For

                Dim key As String = t.Substring(0, eq).Trim().ToUpperInvariant()
                Dim val As String = t.Substring(eq + 1).Trim()

                Select Case key
                    Case "SERVER_URL", "SEAT_SERVER_URL", "SEAT_SERVER"
                        If val.Length > 0 Then SeatServerClient.ServerBaseUrl = val
                    Case "CLIENT_TOKEN"
                        If val.Length > 0 Then SeatServerClient.ClientToken = val
                End Select
            Next
        Catch
        End Try
    End Sub

End Module
