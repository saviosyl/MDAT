Option Strict On
Option Explicit On

Imports System
Imports System.Text
Imports Microsoft.Win32
Imports System.Security.Cryptography

Public NotInheritable Class MachineFingerprint

    Private Sub New()
    End Sub

    Public Shared Function GetMachineId() As String
        Dim machineGuid As String = ReadMachineGuid()
        Dim name As String = Environment.MachineName
        Dim os As String = Environment.OSVersion.VersionString
        Dim cpuCount As String = Environment.ProcessorCount.ToString()

        Dim raw As String = machineGuid & "|" & name & "|" & os & "|" & cpuCount
        Return "MM-" & Sha256Hex(raw)
    End Function

    Private Shared Function ReadMachineGuid() As String
        Try
            Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Cryptography", False)
            If key Is Nothing Then Return "NO_MACHINEGUID"
            Dim v As Object = key.GetValue("MachineGuid", Nothing)
            If v Is Nothing Then Return "NO_MACHINEGUID"
            Dim s As String = Convert.ToString(v)
            If s Is Nothing OrElse s.Trim().Length = 0 Then Return "NO_MACHINEGUID"
            Return s.Trim()
        Catch
            Return "NO_MACHINEGUID"
        End Try
    End Function

    Private Shared Function Sha256Hex(ByVal input As String) As String
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(input)
        Using sha As SHA256 = SHA256.Create()
            Dim hash As Byte() = sha.ComputeHash(bytes)
            Dim sb As New StringBuilder(hash.Length * 2)
            Dim i As Integer
            For i = 0 To hash.Length - 1
                sb.Append(hash(i).ToString("x2"))
            Next
            Return sb.ToString()
        End Using
    End Function

End Class
