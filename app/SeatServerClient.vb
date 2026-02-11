Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Net
Imports System.Text

Public NotInheritable Class SeatServerClient

    Private Sub New()
    End Sub

    Public Shared ServerBaseUrl As String = "https://metamech-license-server.saviosyl.workers.dev"
    Public Shared ClientToken As String = ""

    Private Const CLAIM_PATH As String = "/claim"

    Public Shared Function Claim(ByVal licenseId As String,
                                 ByVal machineId As String,
                                 ByVal seatsMax As Integer,
                                 ByVal machineName As String,
                                 ByVal userName As String) As SeatServerResult

        Dim url As String = CombineUrl(ServerBaseUrl, CLAIM_PATH)

        Dim body As String = _
            "{" & _
            """licenseId"":""" & JsonEsc(licenseId) & """," & _
            """machineId"":""" & JsonEsc(machineId) & """," & _
            """seatsMax"":" & seatsMax.ToString() & "," & _
            """machineName"":""" & JsonEsc(machineName) & """," & _
            """userName"":""" & JsonEsc(userName) & """" & _
            "}"

        Dim json As String = PostJson(url, body)
        Return ParseResult(json)
    End Function

    '===============================
    ' TLS FIX for .NET 4.0 + Cloudflare
    '===============================
    Private Shared Sub EnsureModernTls()
        Try
            ' Turn off 100-Continue delays (some proxies/servers behave better)
            ServicePointManager.Expect100Continue = False
        Catch
        End Try

        Try
            ' .NET 4.0 may not expose Tls11/Tls12 enums, but we can still OR the numeric values:
            ' TLS 1.1 = 768, TLS 1.2 = 3072
            Dim tls11 As SecurityProtocolType = CType(768, SecurityProtocolType)
            Dim tls12 As SecurityProtocolType = CType(3072, SecurityProtocolType)

            ServicePointManager.SecurityProtocol = _
                ServicePointManager.SecurityProtocol Or SecurityProtocolType.Tls Or tls11 Or tls12
        Catch
            ' If OS/framework doesn't support, it will fail silently and keep defaults
        End Try
    End Sub

    Private Shared Function PostJson(ByVal url As String, ByVal jsonBody As String) As String

        EnsureModernTls()

        Dim req As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        req.Method = "POST"
        req.ContentType = "application/json"
        req.Timeout = 12000
        req.ReadWriteTimeout = 12000
        req.KeepAlive = True

        If ClientToken IsNot Nothing AndAlso ClientToken.Trim().Length > 0 Then
            req.Headers("x-mm-token") = ClientToken.Trim()
        End If

        Dim bodyBytes As Byte() = Encoding.UTF8.GetBytes(jsonBody)
        req.ContentLength = bodyBytes.Length

        Using reqStream As Stream = req.GetRequestStream()
            reqStream.Write(bodyBytes, 0, bodyBytes.Length)
        End Using

        Try
            Using resp As HttpWebResponse = CType(req.GetResponse(), HttpWebResponse)
                Using sr As New StreamReader(resp.GetResponseStream(), Encoding.UTF8)
                    Return sr.ReadToEnd()
                End Using
            End Using
        Catch ex As WebException
            ' Try to read server JSON body (helps debug UNAUTH/NO_SEATS/etc.)
            Try
                Dim wr As HttpWebResponse = TryCast(ex.Response, HttpWebResponse)
                If wr IsNot Nothing AndAlso wr.GetResponseStream() IsNot Nothing Then
                    Using sr As New StreamReader(wr.GetResponseStream(), Encoding.UTF8)
                        Dim txt As String = sr.ReadToEnd()
                        If txt IsNot Nothing AndAlso txt.Trim().Length > 0 Then
                            Return txt
                        End If
                    End Using
                End If
            Catch
            End Try

            ' If no response body, rethrow so SeatEnforcer treats it as offline and uses cache
            Throw
        End Try

    End Function

    Private Shared Function ParseResult(ByVal json As String) As SeatServerResult
        Dim r As New SeatServerResult()
        r.RawJson = If(json, "")

        r.Ok = GetJsonBool(json, "ok")
        r.Code = GetJsonString(json, "code")
        r.SeatsUsed = GetJsonInt(json, "seatsUsed")
        r.SeatsMax = GetJsonInt(json, "seatsMax")

        Return r
    End Function

    Private Shared Function GetJsonString(ByVal json As String, ByVal key As String) As String
        If json Is Nothing Then Return ""
        Dim pat As String = """" & key & """"
        Dim i As Integer = json.IndexOf(pat, StringComparison.OrdinalIgnoreCase)
        If i < 0 Then Return ""

        Dim colon As Integer = json.IndexOf(":"c, i + pat.Length)
        If colon < 0 Then Return ""

        Dim q1 As Integer = json.IndexOf(""""c, colon + 1)
        If q1 < 0 Then Return ""

        Dim q2 As Integer = q1 + 1
        Do While q2 < json.Length
            q2 = json.IndexOf(""""c, q2)
            If q2 < 0 Then Return ""
            If json(q2 - 1) <> "\"c Then Exit Do
            q2 += 1
        Loop

        If q2 <= q1 Then Return ""
        Dim val As String = json.Substring(q1 + 1, q2 - q1 - 1)
        Return val.Replace("\""", """").Replace("\\", "\")
    End Function

    Private Shared Function GetJsonInt(ByVal json As String, ByVal key As String) As Integer
        If json Is Nothing Then Return 0
        Dim pat As String = """" & key & """"
        Dim i As Integer = json.IndexOf(pat, StringComparison.OrdinalIgnoreCase)
        If i < 0 Then Return 0

        Dim colon As Integer = json.IndexOf(":"c, i + pat.Length)
        If colon < 0 Then Return 0

        Dim j As Integer = colon + 1
        Do While j < json.Length AndAlso Char.IsWhiteSpace(json(j))
            j += 1
        Loop

        Dim k As Integer = j
        Do While k < json.Length AndAlso (Char.IsDigit(json(k)) OrElse json(k) = "-"c)
            k += 1
        Loop

        Dim s As String = json.Substring(j, k - j).Trim()
        Dim n As Integer
        If Integer.TryParse(s, n) Then Return n
        Return 0
    End Function

    Private Shared Function GetJsonBool(ByVal json As String, ByVal key As String) As Boolean
        If json Is Nothing Then Return False
        Dim pat As String = """" & key & """"
        Dim i As Integer = json.IndexOf(pat, StringComparison.OrdinalIgnoreCase)
        If i < 0 Then Return False

        Dim colon As Integer = json.IndexOf(":"c, i + pat.Length)
        If colon < 0 Then Return False

        Dim rest As String = json.Substring(colon + 1).Trim().ToLowerInvariant()
        If rest.StartsWith("true") Then Return True
        Return False
    End Function

    Private Shared Function JsonEsc(ByVal s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("\", "\\").Replace("""", "\""").Replace(vbCrLf, "\n").Replace(vbLf, "\n")
    End Function

    Private Shared Function CombineUrl(ByVal baseUrl As String, ByVal path As String) As String
        If baseUrl Is Nothing Then baseUrl = ""
        If path Is Nothing Then path = ""
        If baseUrl.EndsWith("/") Then baseUrl = baseUrl.Substring(0, baseUrl.Length - 1)
        If Not path.StartsWith("/") Then path = "/" & path
        Return baseUrl & path
    End Function

End Class

Public Class SeatServerResult
    Public Ok As Boolean
    Public Code As String
    Public SeatsUsed As Integer
    Public SeatsMax As Integer
    Public RawJson As String

    Public Sub New()
        Ok = False
        Code = ""
        SeatsUsed = 0
        SeatsMax = 0
        RawJson = ""
    End Sub
End Class
