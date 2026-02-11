Option Strict On
Option Explicit On

Imports System

Public NotInheritable Class TelemetryService

    Private Sub New()
    End Sub

    ' Call this for START/END/FAIL.
    ' It will:
    '  1) try to flush old queued events
    '  2) try to send this event
    '  3) if send fails, queue it
    Public Shared Sub SendEvent(ByVal syncUrl As String,
                                ByVal token As String,
                                ByVal status As String,
                                ByVal actionSlot As Integer,
                                ByVal actionName As String,
                                ByVal exeVersion As String,
                                ByVal licenseId As String,
                                ByVal machineId As String,
                                ByVal durationMs As Integer,
                                ByVal logText As String)

        If String.IsNullOrEmpty(syncUrl) Then Exit Sub
        If String.IsNullOrEmpty(token) Then Exit Sub

        Dim ts As String = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
        Dim json As String = TelemetryJson.BuildEventJson(ts, status, actionSlot, actionName, exeVersion, licenseId, machineId, durationMs, logText)

        ' Always try flush first (best effort)
        Try
            TelemetryQueue.Flush(syncUrl, token, 8000)
        Catch
        End Try

        ' Send current
        Try
            TelemetryClient.PostJson(syncUrl, token, json, 8000)
        Catch
            TelemetryQueue.Enqueue(json)
        End Try
    End Sub

End Class
