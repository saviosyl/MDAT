Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Collections.Generic

Public NotInheritable Class TelemetryQueue
    Private Sub New()
    End Sub

    Private Shared Function QueuePath() As String
        Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "telemetry_queue.ndjson")
    End Function

    Public Shared Sub Enqueue(jsonLine As String)
        Try
            If jsonLine Is Nothing Then Exit Sub
            Dim t As String = jsonLine.Trim()
            If t = "" Then Exit Sub

            Dim p As String = QueuePath()
            Using sw As New StreamWriter(p, True)
                sw.WriteLine(t)
            End Using
        Catch
        End Try
    End Sub

    Public Shared Sub Flush(syncUrl As String, token As String, timeoutMs As Integer)
        Try
            Dim p As String = QueuePath()
            If Not File.Exists(p) Then Exit Sub

            Dim raw() As String = Nothing
            Try
                raw = File.ReadAllLines(p)
            Catch
                Exit Sub
            End Try

            If raw Is Nothing OrElse raw.Length = 0 Then
                Try
                    File.Delete(p)
                Catch
                End Try
                Exit Sub
            End If

            Dim remaining As New List(Of String)()

            Dim i As Integer
            For i = 0 To raw.Length - 1
                Dim line As String = raw(i)
                If line Is Nothing Then Continue For
                line = line.Trim()
                If line = "" Then Continue For

                Dim resp As String = ""
                Dim ok As Boolean = TelemetryClient.PostJson(syncUrl, token, line, timeoutMs, resp)
                If Not ok Then
                    remaining.Add(line)
                End If
            Next

            If remaining.Count = 0 Then
                Try
                    File.Delete(p)
                Catch
                End Try
            Else
                Try
                    File.WriteAllLines(p, remaining.ToArray())
                Catch
                End Try
            End If

        Catch
        End Try
    End Sub
End Class
