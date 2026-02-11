Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

Public NotInheritable Class SeatCache

    Private Sub New()
    End Sub

    Private Shared Function CacheFolder() As String
        Dim basePath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim folder As String = Path.Combine(basePath, "MetaMech", "MDAT")
        If Not Directory.Exists(folder) Then
            Directory.CreateDirectory(folder)
        End If
        Return folder
    End Function

    Private Shared Function CacheFilePath() As String
        Return Path.Combine(CacheFolder(), "seat.cache")
    End Function

    Public Shared Sub Save(ByVal cache As SeatCacheData)
        Dim plain As String = Serialize(cache)
        Dim plainBytes As Byte() = Encoding.UTF8.GetBytes(plain)
        Dim enc As Byte() = ProtectedData.Protect(plainBytes, Nothing, DataProtectionScope.CurrentUser)
        File.WriteAllBytes(CacheFilePath(), enc)
    End Sub

    Public Shared Function Load() As SeatCacheData
        Dim path As String = CacheFilePath()
        If Not File.Exists(path) Then Return Nothing

        Try
            Dim enc As Byte() = File.ReadAllBytes(path)
            Dim plainBytes As Byte() = ProtectedData.Unprotect(enc, Nothing, DataProtectionScope.CurrentUser)
            Dim plain As String = Encoding.UTF8.GetString(plainBytes)
            Return Deserialize(plain)
        Catch
            Return Nothing
        End Try
    End Function

    Public Shared Sub Clear()
        Dim path As String = CacheFilePath()
        If File.Exists(path) Then
            Try
                File.Delete(path)
            Catch
            End Try
        End If
    End Sub

    Private Shared Function Serialize(ByVal d As SeatCacheData) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("v=1")
        sb.AppendLine("license_id=" & Safe(d.LicenseId))
        sb.AppendLine("machine_id=" & Safe(d.MachineId))
        sb.AppendLine("last_ok_utc=" & d.LastOkUtc.ToString("o"))
        sb.AppendLine("grace_days=" & d.GraceDays.ToString())
        Return sb.ToString()
    End Function

    Private Shared Function Deserialize(ByVal text As String) As SeatCacheData
        Dim d As New SeatCacheData()
        Dim lines As String() = text.Replace(vbCr, "").Split(New Char() {ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)

        Dim i As Integer
        For i = 0 To lines.Length - 1
            Dim line As String = lines(i)
            Dim eq As Integer = line.IndexOf("="c)
            If eq <= 0 Then Continue For

            Dim k As String = line.Substring(0, eq).Trim().ToLowerInvariant()
            Dim v As String = line.Substring(eq + 1)

            Select Case k
                Case "license_id"
                    d.LicenseId = UnSafe(v)
                Case "machine_id"
                    d.MachineId = UnSafe(v)
                Case "last_ok_utc"
                    Dim dt As DateTime
                    If DateTime.TryParse(v, Nothing, Globalization.DateTimeStyles.RoundtripKind, dt) Then
                        d.LastOkUtc = dt
                    Else
                        d.LastOkUtc = DateTime.MinValue
                    End If
                Case "grace_days"
                    Dim n As Integer
                    If Integer.TryParse(v, n) Then d.GraceDays = n
            End Select
        Next

        If d.LicenseId Is Nothing Then d.LicenseId = ""
        If d.MachineId Is Nothing Then d.MachineId = ""

        Return d
    End Function

    Private Shared Function Safe(ByVal s As String) As String
        If s Is Nothing Then Return ""
        Dim t As String = s.Replace("\", "\\").Replace(vbCrLf, "\n").Replace(vbLf, "\n").Replace("=", "\=")
        Return t
    End Function

    Private Shared Function UnSafe(ByVal s As String) As String
        If s Is Nothing Then Return ""
        Dim t As String = s.Replace("\=", "=").Replace("\n", vbLf).Replace("\\", "\")
        Return t
    End Function

End Class

Public Class SeatCacheData
    Public LicenseId As String = ""
    Public MachineId As String = ""
    Public LastOkUtc As DateTime = DateTime.MinValue
    Public GraceDays As Integer = 0
End Class
