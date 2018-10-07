Module ModMisc
    Public Function Interpolate(CeilingMark As Decimal, FloorMark As Decimal, CeilingValue As Decimal, FloorValue As Decimal, RequiredMark As Decimal) As Decimal
        Dim reqCorVal As Decimal

        reqCorVal = ((CeilingValue - FloorValue) * (RequiredMark - FloorMark) / (CeilingMark - FloorMark)) + FloorValue

        Return reqCorVal
    End Function

    Public Function FractiontoDecimal(ByVal Fraction As String) As Decimal
        Dim decimalVal As Decimal = 0
        Dim upper As Decimal = 0
        Dim lower As Decimal = 0
        Dim remain As Decimal = 0
        If Fraction.IndexOf("/") <> -1 Then

            If Fraction.IndexOf(" ") <> -1 Then
                remain = CType(Fraction.Substring(0, Fraction.IndexOf(" ")), Decimal)
                Fraction = Fraction.Substring(Fraction.IndexOf(" "))
            End If
            upper = CType(Fraction.Substring(0, Fraction.IndexOf("/")), Decimal)
            lower = CType(Fraction.Substring(Fraction.IndexOf("/") + 1), Decimal)
            decimalVal = remain + (upper / lower)
        End If
        Return decimalVal
        Return 0
    End Function
End Module
