Option Strict On
Option Explicit On

Public Module TierLocks

    ' -------- DESIGN TOOLS (1–10) --------
    Public Function CanRunDesignTool(slot As Integer, tier As Integer) As Boolean
        Select Case tier
            Case 0 ' Free Trial → ALL ENABLED
                Return True
            Case 1 ' Standard
                Return slot >= 1 AndAlso slot <= 3
            Case 2 ' Pro
                Return slot >= 1 AndAlso slot <= 5
            Case 3 ' Enterprise
                Return True
            Case Else
                Return False
        End Select
    End Function

    ' -------- ENGINEERING TOOLS (11–15) --------
    Public Function CanRunEngineeringTool(slot As Integer, tier As Integer) As Boolean
        Select Case tier
            Case 0 ' Free Trial → ALL ENABLED
                Return True
            Case 1 ' Standard
                Return slot = 11
            Case 2 ' Pro
                Return slot = 11 OrElse slot = 12
            Case 3 ' Enterprise
                Return True
            Case Else
                Return False
        End Select
    End Function

End Module
