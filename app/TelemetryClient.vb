Option Strict On
Option Explicit On

Imports System
Imports System.Net
Imports System.Text
Imports System.IO

Public NotInheritable Class TelemetryClient
    Private Sub New()
    End Sub

    Public Shared Function PostJson(url As String, bearerToken As String, jsonBody As String, timeoutMs As Integer, ByRef responseText As String) As Boolean
        responseText = ""
        If url Is Nothing OrElse url.Trim() = "" Then Return False
        If jsonBody Is Nothing Then jsonBody = ""

        Try
            ' Best-effort TLS enabling (safe for .NET 4.0)
            Try
                Dim proto As Integer = 0
                Try
                    proto = CInt(ServicePointManager.SecurityProtocol)
                Catch
                    proto = 0
                End Try
                proto = proto Or 192 Or 768 Or 3072 ' TLS1.0 + TLS1.1 + TLS1.2
                ServicePointManager.SecurityProtocol = CType(proto, SecurityProtocolType)
            Catch
            End Try

            ServicePointManager.Expect100Continue = False

            Dim req As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
            req.Method = "POST"
            req.ContentType = "application/json"
            req.Accept = "application/json"
            req.Timeout = Math.Max(1000, timeoutMs)
            req.ReadWriteTimeout = Math.Max(1000, timeoutMs)
            req.UserAgent = "MDAT-Telemetry/1.0"

            If bearerToken IsNot Nothing AndAlso bearerToken.Trim() <> "" Then
                req.Headers(HttpRequestHeader.Authorization) = "Bearer " & bearerToken.Trim()
            End If

            Dim bytes() As Byte = Encoding.UTF8.GetBytes(jsonBody)
            req.ContentLength = bytes.Length

            Using s As Stream = req.GetRequestStream()
                s.Write(bytes, 0, bytes.Length)
            End Using

            Using resp As HttpWebResponse = CType(req.GetResponse(), HttpWebResponse)
                Using rs As Stream = resp.GetResponseStream()
                    If rs IsNot Nothing Then
                        Using sr As New StreamReader(rs, Encoding.UTF8)
                            responseText = sr.ReadToEnd()
                        End Using
                    End If
                End Using

                Dim code As Integer = CInt(resp.StatusCode)
                Return (code >= 200 AndAlso code < 300)
            End Using

        Catch ex As WebException
            Try
                If ex.Response IsNot Nothing Then
                    Using r As HttpWebResponse = CType(ex.Response, HttpWebResponse)
                        Using rs As Stream = r.GetResponseStream()
                            If rs IsNot Nothing Then
                                Using sr As New StreamReader(rs, Encoding.UTF8)
                                    responseText = sr.ReadToEnd()
                                End Using
                            End If
                        End Using
                    End Using
                End If
            Catch
            End Try
            Return False

        Catch ex As Exception
            responseText = ex.Message
            Return False
        End Try
    End Function
End Class
