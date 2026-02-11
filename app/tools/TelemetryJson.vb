Option Strict On
Option Explicit On

Imports System
Imports System.Text

Public NotInheritable Class TelemetryJson

    Private Sub New()
    End Sub

    Public Shared Function Escape(ByVal s As String) As String
        If s Is Nothing Then Return ""
        Dim t As String = s
        t = t.Replace("\", "\\")
        t = t.Replace("""", "\""")
        t = t.Replace(vbCrLf, "\n").Replace(vbCr, "\n").Replace(vbLf, "\n")
        Return t
    End Function

    Public Shared Function BuildEventJson(ByVal timestampUtc As String,
                                          ByVal status As String,
                                          ByVal actionSlot As Integer,
                                          ByVal actionName As String,
                                          ByVal exeVersion As String,
                                          ByVal licenseId As String,
                                          ByVal machineId As String,
                                          ByVal durationMs As Integer,
                                          ByVal logText As String) As String

        Dim sb As New StringBuilder(512)
        sb.Append("{")
        sb.Append("""timestamp_utc"":""").Append(Escape(timestampUtc)).Append(""",")
        sb.Append("""status"":""").Append(Escape(status)).Append(""",")
        sb.Append("""action_slot"":").Append(actionSlot.ToString()).Append(",")
        sb.Append("""action_name"":""").Append(Escape(actionName)).Append(""",")
        sb.Append("""exe_version"":""").Append(Escape(exeVersion)).Append(""",")
        sb.Append("""license_id"":""").Append(Escape(licenseId)).Append(""",")
        sb.Append("""machine_id"":""").Append(Escape(machineId)).Append(""",")
        sb.Append("""duration_ms"":").Append(durationMs.ToString()).Append(",")
        sb.Append("""log_text"":""").Append(Escape(logText)).Append("""")
        sb.Append("}")
        Return sb.ToString()
    End Function

End Class
