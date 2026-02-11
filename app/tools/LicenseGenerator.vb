Imports System
Imports System.Text
Imports System.Security.Cryptography

Module LicenseGenerator

    Sub Main()

        Console.WriteLine("============================================")
        Console.WriteLine(" MetaMech License Generator (INTERNAL USE)")
        Console.WriteLine("============================================")
        Console.WriteLine()

        Console.Write("Enter Tier (1=Standard, 2=Premium, 3=Premium Plus): ")
        Dim tierInput As String = Console.ReadLine()

        Dim tier As Integer
        If Not Integer.TryParse(tierInput, tier) OrElse tier < 1 OrElse tier > 3 Then
            Console.WriteLine("ERROR: Invalid tier.")
            Console.ReadLine()
            Return
        End If

        Console.Write("Enter Expiry Date (yyyy-MM-dd): ")
        Dim expiryInput As String = Console.ReadLine()

        Dim expiry As DateTime
        If Not DateTime.TryParseExact(expiryInput, "yyyy-MM-dd",
                                      Globalization.CultureInfo.InvariantCulture,
                                      Globalization.DateTimeStyles.None,
                                      expiry) Then
            Console.WriteLine("ERROR: Invalid expiry date format.")
            Console.ReadLine()
            Return
        End If

        Dim payload As String = tier.ToString() & "|" & expiry.ToString("yyyy-MM-dd")
        Dim payloadBytes() As Byte = Encoding.UTF8.GetBytes(payload)
        Dim payloadBase64 As String = Convert.ToBase64String(payloadBytes)

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


        Dim signatureBase64 As String

        Using rsa As New RSACryptoServiceProvider()
            rsa.FromXmlString(privateKeyXml)
            Using sha As SHA256 = SHA256.Create()
                Dim sigBytes() As Byte = rsa.SignData(payloadBytes, sha)
                signatureBase64 = Convert.ToBase64String(sigBytes)
            End Using
        End Using

        Dim license As String = payloadBase64 & "." & signatureBase64

        Console.WriteLine()
        Console.WriteLine("LICENSE.KEY (copy the line below):")
        Console.WriteLine("--------------------------------------------")
        Console.WriteLine(license)
        Console.WriteLine("--------------------------------------------")
        Console.WriteLine()
        Console.WriteLine("Done.")
        Console.ReadLine()

    End Sub

End Module
