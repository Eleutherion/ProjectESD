Module ModInterpolate
    Public Function Interpolate(CeilingMark As Decimal, FloorMark As Decimal, CeilingValue As Decimal, FloorValue As Decimal, RequiredMark As Decimal) As Decimal
        Dim reqCorVal As Decimal

        reqCorVal = ((CeilingValue - FloorValue) * (RequiredMark - FloorMark) / (CeilingMark - FloorMark)) + FloorValue

        Return reqCorVal
    End Function
End Module
