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

    Public Function VoltageDrop(ByVal LineCurrent As Decimal, ByVal Length As Decimal, Conduit As String, ByVal WireSize As Decimal, Conductor As String, ThreePhase As Boolean, ByVal Voltage As Integer) As Decimal
        Dim NominalXL(20), NominalR(20), TrueR, TrueXL, VD As Decimal
        Dim x As Integer

        Dim Size(20) As Decimal
        Size = {2.0, 3.5, 5.5, 8.0, 14, 22, 30, 38, 50, 60, 80, 100, 125, 150, 175, 200, 250, 325, 375, 400, 500}

        For i As Integer = 0 To 20
            If Size(i) = WireSize Then
                x = i
            End If
        Next i

        Select Case Conduit
            Case "Rigid PVC 80", "Rigid PVC 40", "HDPE"
                NominalXL = {0.058, 0.054, 0.05, 0.052, 0.051, 0.048, 0.045, 0.046, 0.044, 0.043, 0.042, 0.041, 0.041, 0.041, 0.04, 0.04, 0.039, 0.039, 0.038, 0.038, 0.037}
                If Conductor = "Copper" Then
                    NominalR = {3.1, 2.0, 1.2, 0.78, 0.49, 0.31, 0.19, 0.15, 0.12, 0.1, 0.077, 0.062, 0.052, 0.044, 0.038, 0.033, 0.027, 0.023, 0.019, 0.019, 0.015}
                Else
                    NominalR = {0, 3.2, 2.0, 1.3, 0.81, 0.51, 0.32, 0.25, 0.2, 0.16, 0.13, 0.1, 0.085, 0.071, 0.061, 0.054, 0.043, 0.036, 0.029, 0.029, 0.023}
                End If
            Case Else
                NominalXL = {0.073, 0.068, 0.063, 0.065, 0.064, 0.06, 0.057, 0.057, 0.055, 0.054, 0.052, 0.051, 0.052, 0.051, 0.05, 0.049, 0.048, 0.048, 0.048, 0.048, 0.046}
                If Conductor = "Copper" Then
                    NominalR = {3.1, 2.0, 1.2, 0.78, 0.49, 0.31, 0.2, 0.16, 0.12, 0.1, 0.079, 0.063, 0.054, 0.045, 0.039, 0.035, 0.029, 0.025, 0.021, 0.021, 0.018}
                Else
                    NominalR = {0, 3.2, 2.0, 1.3, 0.81, 0.51, 0.32, 0.25, 0.2, 0.16, 0.13, 0.1, 0.086, 0.072, 0.063, 0.055, 0.045, 0.038, 0.031, 0.031, 0.025}
                End If
        End Select

        TrueR = NominalR(x) * Length / 305
        TrueXL = NominalXL(x) * Length / 305

        If ThreePhase = True Then
            VD = Math.Sqrt(3) * LineCurrent * ((TrueR * 0.85) + (TrueXL * Math.Sin(Math.Acos(0.85)))) * 100 / Voltage
        Else
            VD = 2 * LineCurrent * ((TrueR * 0.85) + (TrueXL * Math.Sin(Math.Acos(0.85)))) * 100 / Voltage
        End If

        Return VD
    End Function

    Public Function TransformerRating(ByVal TotalLoad As Decimal) As Decimal
        Dim x As Integer
        Dim rating = {0.01, 0.016, 0.025, 0.063, 0.1, 0.16, 0.2, 0.25, 0.315, 0.4, 0.5, 0.63, 1, 1.25, 1.6, 2, 2.5,
            3, 5, 10, 12.5, 20, 31.5, 40, 50, 63, 85, 100}

        x = 0

        For i As Integer = 1 To 27 Step 1
            If rating(i - 1) > TotalLoad / 1000000 And rating(i) < TotalLoad / 1000000 Then
                x = i
                GoTo ReturnLine
            End If
        Next i

        If x = 0 Then
            rating(x) = 0.01
        End If

ReturnLine: Return rating(x)
    End Function
End Module
