Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports System.Globalization
Imports System.Windows.Forms

'========================================================
' Licensing.vb
' - Provides Licensing.GetLicenseInfo() used by MainForm
' - Reads:  license.key  (payloadBase64.signatureBase64)
' - Uses:   MetaMech_RSA_PUBLIC.xml
' - Payload format: LICENSEID|TIER|EXPIRY_UTC|SEATS
'========================================================
Public Module Licensing

    Private Const LICENSE_FILE As String = "license.key"
    Private Const PUBLIC_KEY_FILE As String = "MetaMech_RSA_PUBLIC.xml"

    '========================================================
    ' MAIN ENTRY POINT (MainForm calls this)
    '========================================================
    Public Function GetLicenseInfo() As LicenseInfo
        Try
            Dim basePath As String = Application.StartupPath

            Dim licPath As String = Path.Combine(basePath, LICENSE_FILE)
            Dim pubPath As String = Path.Combine(basePath, PUBLIC_KEY_FILE)

            If Not File.Exists(licPath) Then
                Return CreateInvalid("License file not found.")
            End If

            If Not File.Exists(pubPath) Then
                Return CreateInvalid("Public key file not found.")
            End If

            Dim licLine As String = ""
            Using sr As New StreamReader(licPath, Encoding.UTF8)
                licLine = sr.ReadLine()
            End Using

            If licLine Is Nothing Then licLine = ""
            licLine = licLine.Trim()

            If licLine = "" OrElse licLine.IndexOf("."c) <= 0 Then
                Return CreateInvalid("License content invalid.")
            End If

            Dim parts() As String = licLine.Split("."c)
            If parts.Length <> 2 Then
                Return CreateInvalid("License format invalid.")
            End If

            Dim payloadBase64 As String = parts(0).Trim()
            Dim sigBase64 As String = parts(1).Trim()

            Dim payloadBytes() As Byte = Nothing
            Dim sigBytes() As Byte = Nothing

            Try
                payloadBytes = Convert.FromBase64String(payloadBase64)
                sigBytes = Convert.FromBase64String(sigBase64)
            Catch
                Return CreateInvalid("Base64 decode failed.")
            End Try

            Dim publicKeyXml As String = File.ReadAllText(pubPath, Encoding.UTF8)

            ' Verify signature (RSA SHA-256)
            Dim okSig As Boolean = VerifySignature(payloadBytes, sigBytes, publicKeyXml)
            If Not okSig Then
                Return CreateInvalid("Signature invalid.")
            End If

            Dim payload As String = Encoding.UTF8.GetString(payloadBytes)

            ' Parse payload -> LicenseInfo
            Dim info As LicenseInfo = ParsePayload(payload)
            If info Is Nothing Then
                Return CreateInvalid("Payload parse failed.")
            End If

            info.IsValid = True
            Return info

        Catch ex As Exception
            Return CreateInvalid("Licensing exception: " & ex.Message)
        End Try
    End Function

    '========================================================
    ' VERIFY SIGNATURE
    '========================================================
    Private Function VerifySignature(payloadBytes() As Byte, sigBytes() As Byte, publicKeyXml As String) As Boolean
        Dim ok As Boolean = False
        Try
            Using rsa As New RSACryptoServiceProvider()
                rsa.FromXmlString(publicKeyXml)
                Using sha As SHA256 = SHA256.Create()
                    ok = rsa.VerifyData(payloadBytes, sha, sigBytes)
                End Using
            End Using
        Catch
            ok = False
        End Try
        Return ok
    End Function

    '========================================================
    ' PARSE PAYLOAD: LICENSEID|TIER|EXPIRY_UTC|SEATS
    '========================================================
    Private Function ParsePayload(payload As String) As LicenseInfo
        If payload Is Nothing Then Return Nothing
        payload = payload.Trim()
        If payload = "" Then Return Nothing

        Dim p() As String = payload.Split("|"c)
        If p.Length < 4 Then
            ' Backward compat: if older format "TIER|EXPIRY" etc.
            Return ParsePayloadLegacy(payload)
        End If

        Dim licenseId As String = SafeStr(p(0))
        Dim tier As Integer = SafeInt(p(1), 0)
        Dim expiryUtc As DateTime = SafeUtcDate(p(2), DateTime.MinValue)
        Dim seats As Integer = SafeInt(p(3), 0)

        Dim li As LicenseInfo = New LicenseInfo()

        ' Required members used by your MainForm
        li.Tier = tier
        li.ExpiryUtc = expiryUtc
        li.Seats = seats

        ' Optional member (only set if it exists in your LicenseInfo class)
        Try
            li.LicenseId = licenseId
        Catch
            ' ignore if property does not exist
        End Try

        Return li
    End Function

    '========================================================
    ' LEGACY PARSE (if old payloads exist)
    ' Examples supported:
    '   "TIER|EXPIRY_UTC"
    '   "TIER|EXPIRY_UTC|SEATS"
    '========================================================
    Private Function ParsePayloadLegacy(payload As String) As LicenseInfo
        Dim p() As String = payload.Split("|"c)
        If p.Length < 2 Then Return Nothing

        Dim tier As Integer = SafeInt(p(0), 0)
        Dim expiryUtc As DateTime = SafeUtcDate(p(1), DateTime.MinValue)
        Dim seats As Integer = 0
        If p.Length >= 3 Then seats = SafeInt(p(2), 0)

        Dim li As LicenseInfo = New LicenseInfo()
        li.Tier = tier
        li.ExpiryUtc = expiryUtc
        li.Seats = seats
        Return li
    End Function

    '========================================================
    ' HELPERS
    '========================================================
    Private Function CreateInvalid(reason As String) As LicenseInfo
        Dim li As LicenseInfo = New LicenseInfo()
        li.IsValid = False
        li.Tier = 0
        li.ExpiryUtc = DateTime.MinValue
        li.Seats = 0
        Try
            li.ErrorMessage = reason
        Catch
            ' ignore if not present
        End Try
        Return li
    End Function

    Private Function SafeStr(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Trim()
    End Function

    Private Function SafeInt(s As String, defaultVal As Integer) As Integer
        If s Is Nothing Then Return defaultVal
        s = s.Trim()
        If s = "" Then Return defaultVal
        Dim v As Integer = defaultVal
        If Integer.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, v) Then
            Return v
        End If
        Return defaultVal
    End Function

    ' Accepts common UTC formats:
    '  - 2026-02-03
    '  - 2026-02-03T12:34:56Z
    '  - 2026-02-03 12:34:56
    Private Function SafeUtcDate(s As String, defaultVal As DateTime) As DateTime
        If s Is Nothing Then Return defaultVal
        s = s.Trim()
        If s = "" Then Return defaultVal

        Dim dt As DateTime = defaultVal

        Dim formats() As String = New String() {
            "yyyy-MM-dd",
            "yyyy-MM-ddTHH:mm:ssZ",
            "yyyy-MM-ddTHH:mm:ss",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss.fffZ",
            "yyyy-MM-ddTHH:mm:ss.fff"
        }

        If DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture,
                                  DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal, dt) Then
            Return DateTime.SpecifyKind(dt, DateTimeKind.Utc)
        End If

        ' fallback parse
        If DateTime.TryParse(s, CultureInfo.InvariantCulture,
                             DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal, dt) Then
            Return DateTime.SpecifyKind(dt, DateTimeKind.Utc)
        End If

        Return defaultVal
    End Function

End Module
