Option Strict On
Option Explicit On

Imports System
Imports System.Text
Imports System.Security.Cryptography
Imports Microsoft.Win32

' ============================================================
' MACHINE ID (STABLE + PERMANENT)
'
' Uses Windows MachineGuid
' Fallback: MachineName
'
' Output:
'   SHA256 hex string (uppercase)
' ============================================================

Public Module MachineId

    ' ================= PUBLIC =================
    Public Function GetMachineFingerprint() As String
        Dim source As String = ""

        Try
            source = ReadMachineGuid()
        Catch
            source = ""
        End Try

        If String.IsNullOrEmpty(source) Then
            source = Environment.MachineName
        End If

        Return Sha256Hex(source)
    End Function

    ' ================= PRIVATE =================
    Private Function ReadMachineGuid() As String
        Dim rk As RegistryKey = Nothing

        Try
            rk = Registry.LocalMachine.OpenSubKey(
                "SOFTWARE\Microsoft\Cryptography", False)

            If rk IsNot Nothing Then
                Dim v As Object = rk.GetValue("MachineGuid")
                If v IsNot Nothing Then
                    Return v.ToString()
                End If
            End If

        Finally
            If rk IsNot Nothing Then
                rk.Close()
            End If
        End Try

        Return ""
    End Function

    Private Function Sha256Hex(input As String) As String
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(input)

        Dim sha As SHA256 = SHA256.Create()
        Dim hash As Byte() = sha.ComputeHash(bytes)

        sha.Dispose()

        Dim sb As New StringBuilder(hash.Length * 2)
        For Each b As Byte In hash
            sb.Append(b.ToString("X2"))
        Next

        Return sb.ToString()
    End Function

End Module
