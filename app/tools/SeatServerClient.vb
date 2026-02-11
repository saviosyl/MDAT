Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Net
Imports System.Text

Public NotInheritable Class SeatServerClient

    Private Sub New()
    End Sub

    ' Your Cloudflare Worker:
    Public Shared ServerBaseUrl As String = "https://metamech-license-server.saviosyl.workers.dev"

    ' Must match env.CLIENT_TOKEN in Worker
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

    Private Shared Function PostJson(ByVal url As String, ByVal jsonBody As String) As String
        Dim req As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        req.Method = "POST"
        req.ContentType = "application/json"
        req.Timeout = 8000
        req.ReadWriteTimeout = 8000

        If ClientToken IsNot Nothing AndAlso ClientToken.Trim().Length > 0 Then
            req.Headers("x-mm-token") = ClientToken.Trim()
        End If

        Dim bodyBytes As Byte() = Encoding.UTF8.GetBytes(jsonBody)
        req.ContentLength = bodyBytes.Length

        Using reqStream As Stream = req.GetRequestStream()
            reqStream.Write(bodyBytes, 0, bodyBytes.Length)
        End Using

        Using resp As HttpWebResponse = CType(req.GetResponse(), HttpWebResponse)
            Using sr As New StreamReader(resp.GetResponseStream(), Encoding.UTF8)
                Return sr.ReadToEnd()
            End Using
        End Using
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
