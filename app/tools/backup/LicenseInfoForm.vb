Option Strict On
Option Explicit On

Public Class LicenseInfo
    Public Property IsValid As Boolean = False
    Public Property LicenseId As String = ""
    Public Property Tier As Integer = 0
    Public Property Seats As Integer = 0
    Public Property ExpiryUtc As DateTime
    Public Property ErrorMessage As String = ""
End Class
