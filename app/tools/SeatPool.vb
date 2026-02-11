Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Text
Imports System.Collections.Generic

Public Module SeatPool

    Private ReadOnly SeatPoolPath As String =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "MetaMech", "Seats")

    Private _claimedLicenseId As String = ""
    Private _claimedMachineId As String = ""

    ' ============================================================
    ' CLAIM SEAT
    ' ============================================================
    Public Function ClaimSeat(lic As LicenseInfo, ByRef err As String) As Boolean

        err = ""

        Try
            If lic Is Nothing OrElse Not lic.IsValid Then
                err = "Invalid license."
                Return False
            End If

            If lic.Seats <= 0 Then
                err = "No seats available."
                Return False
            End If

            _claimedLicenseId = lic.LicenseId

            Try
                _claimedMachineId = MachineId.GetMachineFingerprint()
            Catch ex As Exception
                err = "Machine ID error: " & ex.Message
                Return False
            End Try

            Directory.CreateDirectory(SeatPoolPath)

            CleanupStale(lic.LicenseId)

            Dim seatFile As String =
                Path.Combine(
                    SeatPoolPath,
                    SafeName(lic.LicenseId) & "__" & _claimedMachineId & ".seat")

            If File.Exists(seatFile) Then
                Return True ' already claimed on this machine
            End If

            Dim active As List(Of String) = GetActiveSeatFiles(lic.LicenseId)

            If active.Count >= lic.Seats Then
                err = "All seats are in use (" &
                      active.Count & "/" & lic.Seats & ")."
                Return False
            End If

            File.WriteAllText(
                seatFile,
                "LICENSE=" & lic.LicenseId & vbCrLf &
                "MACHINE=" & _claimedMachineId & vbCrLf &
                "UTC=" & DateTime.UtcNow.ToString("o"))

            Return True

        Catch ex As Exception
            err = "Seat claim error: " & ex.Message
            Return False
        End Try

    End Function

    ' ============================================================
    ' HELPERS
    ' ============================================================
    Private Function GetActiveSeatFiles(licenseId As String) As List(Of String)
        Dim list As New List(Of String)

        If Not Directory.Exists(SeatPoolPath) Then Return list

        For Each f As String In Directory.GetFiles(SeatPoolPath, SafeName(licenseId) & "__*.seat")
            list.Add(f)
        Next

        Return list
    End Function

    Private Sub CleanupStale(licenseId As String)
        If Not Directory.Exists(SeatPoolPath) Then Return

        For Each f As String In Directory.GetFiles(SeatPoolPath, SafeName(licenseId) & "__*.seat")
            Try
                Dim age As TimeSpan = DateTime.UtcNow - File.GetLastWriteTimeUtc(f)
                If age.TotalDays > 7 Then
                    File.Delete(f)
                End If
            Catch
            End Try
        Next
    End Sub

    Private Function SafeName(s As String) As String
        For Each c As Char In Path.GetInvalidFileNameChars()
            s = s.Replace(c, "_"c)
        Next
        Return s
    End Function

End Module
