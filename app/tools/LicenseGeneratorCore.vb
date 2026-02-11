Option Strict On
Option Explicit On

Imports System
Imports System.Text
Imports System.Security.Cryptography
Imports System.Globalization

Public Module LicenseGeneratorCore

    ' ============================================================
    ' REQUIRED PAYLOAD FORMAT (matches your MDAT licensing):
    '   LICENSEID|TIER|EXPIRY_UTC|SEATS
    '
    ' License string format:
    '   Base64(payloadBytes) . Base64(signatureBytes)
    ' ============================================================

    Public Const TRIAL_DAYS As Integer = 3

    Public Function GenerateLicense(ByVal licenseId As String,
                                    ByVal tier As Integer,
                                    ByVal expiryUtc As DateTime,
                                    ByVal seats As Integer) As String

        If licenseId Is Nothing OrElse licenseId.Trim() = "" Then
            Throw New ApplicationException("License ID is required.")
        End If

        If tier < 0 OrElse tier > 3 Then
            Throw New ApplicationException("Tier must be 0..3 (0=Trial, 1..3 Paid).")
        End If

        If seats <= 0 OrElse seats > 999 Then
            Throw New ApplicationException("Seats must be between 1 and 999.")
        End If

        ' Expiry must be UTC
        Dim expUtc As DateTime = expiryUtc
        If expUtc.Kind <> DateTimeKind.Utc Then
            expUtc = DateTime.SpecifyKind(expUtc, DateTimeKind.Utc)
        End If

        ' Use a stable UTC format
        Dim expText As String = expUtc.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)

        Dim payload As String = licenseId.Trim() & "|" & tier.ToString(CultureInfo.InvariantCulture) & "|" &
                                expText & "|" & seats.ToString(CultureInfo.InvariantCulture)

        Dim payloadBytes() As Byte = Encoding.UTF8.GetBytes(payload)
        Dim payloadBase64 As String = Convert.ToBase64String(payloadBytes)

        Dim signatureBase64 As String = SignPayload(payloadBytes)

        Return payloadBase64 & "." & signatureBase64
    End Function

    Private Function SignPayload(ByVal payloadBytes() As Byte) As String

        ' ==================================================
        ' PRIVATE RSA KEY (KEEP SECRET – DO NOT SHIP)
        ' ==================================================
        Dim privateKeyXml As String = _
            "<RSAKeyValue>" & _
            "<Modulus>oekFBZi1rppt77E+9ND1pzXW9HlQaO4taHwqg9lZDKZoJ8dfghA0g5K67SjAKiIi7L/+98XcsABzkxJSMlp0+IpDXgrnVuL8e7Jn6frSUOFaixjSBINY+Ez3TB/StbjIARIuw+L8y8ofx9VTWdy96Lq+tieBUgECN8hEfrA/BI3SFGNqe5Ksk/2tj9pGmxiSOwEN+pHGYkNkaQs8hThUtKXxWcrJhbBzzPrawiK87ekSo+Y44oNUirBqJ0mfvP7rdI1qhuOu1gu+gGq5ZAgJLGK0l3xJQU6BZwm5bIUbfNPvTgCluq2ziSVKyWSHFW+H2xStyCAxq5LvNc684FLwYQ==</Modulus>" & _
            "<Exponent>AQAB</Exponent>" & _
            "<P>xH9IrzWzkmeRJAPdhoax6dkg+l/JpzShZYdRGtmbSPiYIEDjyfmueLc1Xm2bdoVPl2zsICmWjA+tYNZXZcN1dpXnusNeNOdHyRQFj444Hmo0qAlXR5kRxpuUb6i4DuFkcVDyUJAQPxl/klb/DNvKM1vd0n7o7R6c1de+Y9eOpAc=</P>" & _
            "<Q>0vCCXzmnuyawmtptrLZGIPtKiIz6tu7uO10r/GIpMiOOA/4tdZRriU1sxCrACctFl49M9BLohQ5uFeGWAYfI6Ay5CdbqIvY9iEl1cz5i7Qizt6fyKUV8I8YO/OZzeeF0RtsERaY5edpv72ohH36UGbr++4YZRTOIP8x0BAcEvlc=</Q>" & _
            "<DP>Xf+hzqc64wOGTBtJQsx2ma6T9xIRjlpAByinZNfKUCsT4wIHthwqGXdTIXv/jcASJhcmEfCzIwdw4k1G+9h3/aWSeCZzj41AKvHYAyd+sxYNYIEvboHMHh1Y8d1dR0kNWqPldyKjkvvoqiHR2t3dqZn90G+Dj8NANZACdRKuGss=</DP>" & _
            "<DQ>toxFhztSGPimpZyahXlIv4o5OmsnHeEwcldzlXstw6JZaMMzfCnx1mUW171XbCJqG3t8UU17xIp0YqNTOgfUql04VXeUMKWBIszPw+gdnJyHS00gmO71O9BPcDXPgY7HHfq0e8Iaw4VykXL7L1JPwOS/fdTTUbwDEZNSY5nfVQk=</DQ>" & _
            "<InverseQ>RggLqdL4YZm7Y3jgXHUwWMmbgQpHnEZmRFpK7c2WTtUDyMUyrHQuHdmxVR3RHSj1AHq/+Bqnst2NmvFHeZdr1u9X9TQPxhZzpKBeNwDEAI/Ye5WUoPvdWOwEGgenvhk1AKCIA/PKhsh8x5y319FvjGjy9lNhVMIc71X6v1HhNIg=</InverseQ>" & _
            "<D>cwwIrXldX7wL8g2YFo2EgFQZcY3iPS1AxsWz0AxY4kw8Jkfc7aZmKjfQ60PRiB7JgkDLA3Rs5ALuHMsf7PepthFI3UISMAMKNTsH65J2b1Ix7DSuxtYuGgFWl5jlOIscUuaApGBeENCG1JAYsfnQV9aaPQTFN2fQE6MSSJMjtC+QW/SHM4cE1iY77Z8ku/WRQKtyK1SRhjxl6YK3rAHN1Da7nwnHp04Xe6zO76DKyWAam8LED9NntXgBold9JBcV+PSW0hFvMxkR44lBzUd84payTQBEEaeXqBaQ95f2UoDXJDDR5djE7dFSD7XCYrpTb6nkLdja4MjVmEnlfWCHoQ==</D>" & _
            "</RSAKeyValue>"

        Using rsa As New RSACryptoServiceProvider()
            rsa.FromXmlString(privateKeyXml)

            Using sha As SHA256 = SHA256.Create()
                Dim sigBytes() As Byte = rsa.SignData(payloadBytes, sha)
                Return Convert.ToBase64String(sigBytes)
            End Using
        End Using

    End Function

End Module
