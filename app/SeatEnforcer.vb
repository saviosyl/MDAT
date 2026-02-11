Option Strict On
Option Explicit On

Imports System
Imports System.Net

Public NotInheritable Class SeatEnforcer

    Private Sub New()
    End Sub

    Public Shared Function GetDefaultGraceDays(ByVal tier As Integer) As Integer
        Select Case tier
            Case 0 : Return 1
            Case 1 : Return 3
            Case 2 : Return 7
            Case 3 : Return 14
            Case Else : Return 3
        End Select
    End Function

    ' Returns True if allowed. Throws Exception with a clear reason if blocked.
    Public Shared Function EnsureSeatOrThrow(ByVal licenseId As String, ByVal tier As Integer, ByVal seatsMax As Integer) As Boolean

        If licenseId Is Nothing OrElse licenseId.Trim().Length = 0 Then
            Throw New Exception("License ID missing. Please activate a valid license.")
        End If

        Dim machineId As String = MachineFingerprint.GetMachineId()
        Dim machineName As String = Environment.MachineName
        Dim userName As String = Environment.UserName

        If seatsMax <= 0 Then seatsMax = 1

        ' -------------------------
        ' ONLINE ATTEMPT (CLAIM)
        ' -------------------------
        Dim onlineAttempted As Boolean = False
        Try
            onlineAttempted = True

            Dim res As SeatServerResult = SeatServerClient.Claim(licenseId.Trim(), machineId, seatsMax, machineName, userName)

            If res IsNot Nothing AndAlso res.Ok Then
                ' Save offline grace cache
                Dim grace As Integer = GetDefaultGraceDays(tier)

                Dim cache As New SeatCacheData()
                cache.LicenseId = licenseId.Trim()
                cache.MachineId = machineId
                cache.LastOkUtc = DateTime.UtcNow
                cache.GraceDays = grace

                SeatCache.Save(cache)
                Return True
            End If

            ' If we got a response, it is NOT an "internet down" case.
            ' Surface the real reason to the user.
            If res IsNot Nothing Then
                Dim code As String = If(res.Code, "").Trim().ToUpperInvariant()

                If code = "UNAUTH" Then
                    Throw New Exception("Seat server rejected this request (UNAUTH). " &
                                        "Check output\config.txt has CLIENT_TOKEN=... and it matches Cloudflare Worker CLIENT_TOKEN exactly.")
                ElseIf code = "NO_SEATS" Then
                    Throw New Exception("No seats available for this license right now. " &
                                        "Seats used: " & res.SeatsUsed.ToString() & " / " & res.SeatsMax.ToString() & ".")
                ElseIf code = "NOT_FOUND" Then
                    Throw New Exception("Seat server endpoint not found (NOT_FOUND). " &
                                        "Your Worker must support POST /claim exactly.")
                ElseIf code = "METHOD" Then
                    Throw New Exception("Seat server rejected method (METHOD). Expected POST.")
                ElseIf code <> "" Then
                    Throw New Exception("Seat server rejected this request (" & code & "). Raw: " & Trunc(res.RawJson, 220))
                Else
                    ' Unknown but still a server response
                    Throw New Exception("Seat server rejected this request. Raw: " & Trunc(res.RawJson, 220))
                End If
            End If

        Catch ex As WebException
            ' Network/TLS/proxy issues fall through to offline cache.
            ' If no cache exists, we will show the true network error.
            If SeatCache.Load() Is Nothing Then
                Dim msg As String = "Seat server network error. " & ex.Message
                Throw New Exception(msg)
            End If
            ' else: allow offline fallback below
        Catch ex As Exception
            ' If we already got a meaningful failure (UNAUTH/NO_SEATS/etc.), do NOT mask it as "no internet".
            Throw
        End Try

        ' -------------------------
        ' OFFLINE FALLBACK
        ' -------------------------
        Dim c As SeatCacheData = SeatCache.Load()
        If c Is Nothing Then
            If onlineAttempted Then
                Throw New Exception("Seat activation failed and no offline seat cache found. Please connect to the internet to activate your seat.")
            Else
                Throw New Exception("No offline seat cache found. Please connect to the internet to activate your seat.")
            End If
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

    Private Shared Function Trunc(ByVal s As String, ByVal maxLen As Integer) As String
        If s Is Nothing Then Return ""
        Dim t As String = s.Trim()
        If t.Length <= maxLen Then Return t
        Return t.Substring(0, maxLen) & "..."
    End Function

End Class
