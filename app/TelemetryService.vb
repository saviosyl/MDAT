Option Strict On
Option Explicit On

Imports System

Public NotInheritable Class TelemetryService
    Private Sub New()
    End Sub

    Public Shared Sub SendEvent(syncUrl As String,
                                token As String,
                                status As String,
                                actionSlot As Integer,
                                actionName As String,
                                exeVersion As String,
                                licenseId As String,
                                machineId As String,
                                durationMs As Integer,
                                logText As String)

        Try
            If syncUrl Is Nothing OrElse syncUrl.Trim() = "" Then Exit Sub
            If token Is Nothing OrElse token.Trim() = "" Then Exit Sub

            Dim ts As String = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")

            Dim st As String = If(status, "").Trim().ToUpperInvariant()
            If st = "" Then st = "INFO"

            Dim json As String = BuildJson(ts, st, actionSlot, If(actionName, ""), If(exeVersion, ""), If(licenseId, ""), If(machineId, ""), durationMs, If(logText, ""))

            ' try immediate send
            Dim resp As String = ""
            Dim ok As Boolean = TelemetryClient.PostJson(syncUrl, token, json, 8000, resp)

            If Not ok Then
                ' queue if fail
                TelemetryQueue.Enqueue(json)
            End If

            ' best-effort flush backlog
            TelemetryQueue.Flush(syncUrl, token, 8000)

        Catch
        End Try
    End Sub

    Private Shared Function BuildJson(timestampUtc As String,
                                      status As String,
                                      actionSlot As Integer,
                                      actionName As String,
                                      exeVersion As String,
                                      licenseId As String,
                                      machineId As String,
                                      durationMs As Integer,
                                      logText As String) As String

        Dim s As String = "{"
        s &= """timestamp_utc"":" & TelemetryJson.Quote(timestampUtc) & ","
        s &= """status"":" & TelemetryJson.Quote(status) & ","
        s &= """action_slot"":" & actionSlot.ToString() & ","
        s &= """action_name"":" & TelemetryJson.Quote(actionName) & ","
        s &= """exe_version"":" & TelemetryJson.Quote(exeVersion) & ","
        s &= """license_id"":" & TelemetryJson.Quote(licenseId) & ","
        s &= """machine_id"":" & TelemetryJson.Quote(machineId) & ","
        s &= """duration_ms"":" & durationMs.ToString() & ","
        s &= """log_text"":" & TelemetryJson.Quote(logText)
        s &= "}"
        Return s
    End Function
End Class
