Option Strict On
Option Explicit On

Imports System
Imports System.Net

Public NotInheritable Class SeatEnforcer

    Private Sub New()
    End Sub

    Public Shared Function GetDefaultGraceDays(ByVal tier As Integer) As Integer
        Select Case tier
            Case 0 : Return 1     ' trial
            Case 1 : Return 3     ' standard
            Case 2 : Return 7     ' premium
            Case 3 : Return 14    ' premium plus
            Case Else : Return 3
        End Select
    End Function

    ' Call AFTER RSA license verification. Throws user-friendly Exception if blocked.
    Public Shared Function EnsureSeatOrThrow(ByVal licenseId As String, ByVal tier As Integer, ByVal seatsMax As Integer) As Boolean

        If licenseId Is Nothing OrElse licenseId.Trim().Length = 0 Then
            Throw New Exception("License ID missing. Please activate a valid license.")
        End If

        Dim machineId As String = MachineFingerprint.GetMachineId()
        Dim machineName As String = Environment.MachineName
        Dim userName As String = Environment.UserName

        ' Online attempt: /claim (your Worker updates lastSeenUtc)
        Try
            If seatsMax <= 0 Then seatsMax = 1

            Dim res As SeatServerResult = SeatServerClient.Claim(licenseId.Trim(), machineId, seatsMax, machineName, userName)

            If res IsNot Nothing AndAlso res.Ok Then
                Dim grace As Integer = GetDefaultGraceDays(tier)

                Dim cache As New SeatCacheData()
                cache.LicenseId = licenseId.Trim()
                cache.MachineId = machineId
                cache.LastOkUtc = DateTime.UtcNow
                cache.GraceDays = grace

                SeatCache.Save(cache)
                Return True
            End If

            ' If online returns NO_SEATS, we still fall back to offline cache
            ' (Only works if this machine was previously activated and still within grace)

        Catch ex As WebException
            ' network down -> offline fallback
        Catch ex As Exception
            ' any error -> offline fallback
        End Try

        ' Offline fallback
        Dim c As SeatCacheData = SeatCache.Load()
        If c Is Nothing Then
            Throw New Exception("No internet connection and no offline seat cache found. Please connect to the internet to activate your seat.")
        End If

        If c.LicenseId <> licenseId.Trim() OrElse c.MachineId <> machineId Then
            Throw New Exception("Offline cache does not match this machine. Please connect to the internet to renew/activate your seat.")
        End If

        Dim graceDays As Integer = c.GraceDays
        If graceDays <= 0 Then graceDays = GetDefaultGraceDays(tier)

        Dim deadlineUtc As DateTime = c.LastOkUtc.AddDays(CDbl(graceDays))
        If DateTime.UtcNow <= deadlineUtc Then
            Return True
        End If

        Throw New Exception("Offline grace period has expired. Please connect to the internet to renew your seat.")
    End Function

End Class
