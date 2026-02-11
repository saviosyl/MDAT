Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Text
Imports System.Net

Public NotInheritable Class TelemetryQueue

    Private Sub New()
    End Sub

    ' Where queued JSON lines are stored (one JSON per line)
    Public Shared Function QueueFilePath() As String
        Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sync_queue.jsonl")
    End Function

    Public Shared Sub Enqueue(ByVal json As String)
        Try
            If json Is Nothing Then Exit Sub
            File.AppendAllText(QueueFilePath(), json & Environment.NewLine, Encoding.UTF8)
        Catch
            ' ignore - never crash EXE
        End Try
    End Sub

    ' Try to send queued events. Leaves any failed ones in the file.
    Public Shared Sub Flush(ByVal syncUrl As String, ByVal token As String, ByVal timeoutMs As Integer)
        Dim path As String = QueueFilePath()
        If Not File.Exists(path) Then Exit Sub

        Dim lines As String() = Nothing
        Try
            lines = File.ReadAllLines(path, Encoding.UTF8)
        Catch
            Exit Sub
        End Try

        If lines Is Nothing OrElse lines.Length = 0 Then
            Try
                File.Delete(path)
            Catch
            End Try
            Exit Sub
        End If

        Dim remaining As New System.Collections.Generic.List(Of String)()

        Dim i As Integer
        For i = 0 To lines.Length - 1
            Dim ln As String = lines(i)
            If String.IsNullOrEmpty(ln) Then
                Continue For
            End If

            Try
                ' If this succeeds, do not keep the line
                TelemetryClient.PostJson(syncUrl, token, ln, timeoutMs)
            Catch
                ' Keep it for later
                remaining.Add(ln)
            End Try
        Next

        Try
            If remaining.Count = 0 Then
                File.Delete(path)
            Else
                File.WriteAllLines(path, remaining.ToArray(), Encoding.UTF8)
            End If
        Catch
            ' ignore
        End Try
    End Sub

End Class
