Option Strict On
Option Explicit On

Imports System
Imports System.Globalization

Public Class LicenseInfo

    Public Property IsValid As Boolean
    Public Property ErrorMessage As String

    Public Property LicenseId As String
    Public Property Tier As Integer          ' 0=Trial, 1=Standard, 2=Premium, 3=Premium Plus
    Public Property Seats As Integer
    Public Property ExpiryUtc As DateTime

    Public ReadOnly Property IsExpired As Boolean
        Get
            If ExpiryUtc = DateTime.MinValue Then Return True
            Return DateTime.UtcNow > ExpiryUtc
        End Get
    End Property

    Public ReadOnly Property DaysRemaining As Integer
        Get
            If ExpiryUtc = DateTime.MinValue Then Return 0
            Dim ts As TimeSpan = ExpiryUtc.Subtract(DateTime.UtcNow)
            Dim days As Integer = CInt(Math.Floor(ts.TotalDays))
            If days < 0 Then days = 0
            Return days
        End Get
    End Property

    Public ReadOnly Property TierName As String
        Get
            Select Case Tier
                Case 0 : Return "TRIAL"
                Case 1 : Return "STANDARD"
                Case 2 : Return "PREMIUM"
                Case 3 : Return "PREMIUM PLUS"
                Case Else : Return "UNKNOWN"
            End Select
        End Get
    End Property

    ' Convenience: UI-friendly one-line status (no layout changes needed)
    Public ReadOnly Property HeaderLine As String
        Get
            If Not IsValid Then
                Return "LICENCE: INVALID"
            End If

            Dim activeText As String = If(IsExpired, "EXPIRED", "ACTIVE")
            Dim seatText As String = Seats.ToString(CultureInfo.InvariantCulture)

            Return "LICENCE: " & activeText & "  |  TIER: " & TierName & "  |  SEATS: " & seatText
        End Get
    End Property

    Public ReadOnly Property ExpiryLine As String
        Get
            If Not IsValid Then
                If ErrorMessage Is Nothing OrElse ErrorMessage.Trim() = "" Then
                    Return "ERROR: License not valid"
                End If
                Return "ERROR: " & ErrorMessage.Trim()
            End If

            If IsExpired Then
                Return "❌ EXPIRED"
            End If

            Dim d As Integer = DaysRemaining
            If Tier = 0 Then
                Return "⏳ TRIAL — " & d.ToString(CultureInfo.InvariantCulture) & " days remaining"
            End If

            If d <= 3 Then
                Return "⛔ EXPIRING IN " & d.ToString(CultureInfo.InvariantCulture) & " DAYS"
            ElseIf d <= 30 Then
                Return "⚠ EXPIRING IN " & d.ToString(CultureInfo.InvariantCulture) & " DAYS"
            End If

            Return "✅ " & d.ToString(CultureInfo.InvariantCulture) & " days remaining"
        End Get
    End Property

End Class
