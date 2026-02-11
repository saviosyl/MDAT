Public Module UnitSystem
    Public Enum Units
        Metric
        Imperial
    End Enum

    Public CurrentUnit As Units = Units.Metric

    Public Function MmToIn(mm As Double) As Double
        Return mm / 25.4
    End Function

    Public Function InToMm(inch As Double) As Double
        Return inch * 25.4
    End Function
End Module
