Module ModIllu
    Public EDD As Decimal
    Public Function LLD(LampType As String) As Decimal
        Dim value

        If LampType = "General Service Incandescent" Or "Metal Halide" Then
            value = 0.8
        ElseIf LampType = "Halogen Incandescent" Or "High Pressure Sodium" Then
            value = 0.9
        ElseIf LampType = "Light Loading Fluorescent" Then
            value = 0.84
        ElseIf LampType = "Medium Loading Fluorescent" Then
            value = 0.79
        ElseIf LampType = "Heavy Loading Fluorescent" Then
            value = 0.72
        Else
            value = 0.83
        End If
        Return value
    End Function
    Public Function LDDRating(LumCat As Integer, Atmosphere As String, Frequency As Decimal) As Decimal
        Dim LDD, a, b, t As Decimal
        t = Frequency / 12
        If LumCat = 1 Then
            b = 0.69
            If Atmosphere = "Very Clean" Then
                a = 0.038
            ElseIf Atmosphere = "Clean" Then
                a = 0.071
            ElseIf Atmosphere = "Medium" Then
                a = 0.111
            ElseIf Atmosphere = "Dirty" Then
                a = 0.162
            ElseIf Atmosphere = "Very Dirty" Then
                a = 0.301
            End If
        ElseIf LumCat = 2 Then
            b = 0.62
            If Atmosphere = "Very Clean" Then
                a = 0.033
            ElseIf Atmosphere = "Clean" Then
                a = 0.068
            ElseIf Atmosphere = "Medium" Then
                a = 0.102
            ElseIf Atmosphere = "Dirty" Then
                a = 0.147
            ElseIf Atmosphere = "Very Dirty" Then
                a = 0.188
            End If
        ElseIf LumCat = 3 Then
            b = 0.7
            If Atmosphere = "Very Clean" Then
                a = 0.079
            ElseIf Atmosphere = "Clean" Then
                a = 0.106
            ElseIf Atmosphere = "Medium" Then
                a = 0.143
            ElseIf Atmosphere = "Dirty" Then
                a = 0.184
            ElseIf Atmosphere = "Very Dirty" Then
                a = 0.236
            End If
        ElseIf LumCat = 4 Then
            b = 0.72
            If Atmosphere = "Very Clean" Then
                a = 0.07
            ElseIf Atmosphere = "Clean" Then
                a = 0.131
            ElseIf Atmosphere = "Medium" Then
                a = 0.216
            ElseIf Atmosphere = "Dirty" Then
                a = 0.314
            ElseIf Atmosphere = "Very Dirty" Then
                a = 0.452
            End If
        ElseIf LumCat = 5 Then
            b = 0.53
            If Atmosphere = "Very Clean" Then
                a = 0.078
            ElseIf Atmosphere = "Clean" Then
                a = 0.128
            ElseIf Atmosphere = "Medium" Then
                a = 0.19
            ElseIf Atmosphere = "Dirty" Then
                a = 0.249
            ElseIf Atmosphere = "Very Dirty" Then
                a = 0.321
            End If
        ElseIf LumCat = 6 Then
            b = 0.88
            If Atmosphere = "Very Clean" Then
                a = 0.076
            ElseIf Atmosphere = "Clean" Then
                a = 0.145
            ElseIf Atmosphere = "Medium" Then
                a = 0.218
            ElseIf Atmosphere = "Dirty" Then
                a = 0.284
            ElseIf Atmosphere = "Very Dirty" Then
                a = 0.396
            End If
        End If

        LDD = Math.E ^ (-a * (t ^ b))
        EDD = 100 * (1 - LDD)

        Return LDD
    End Function
    'Public Function RSDD(RCR As Decimal, EDD As Decimal, LDT As String) As Decimal
    'Dim value, nomRSDD As Decimal, nomRCR As Integer
    'If LDT = "Direct" Then
    'Select Case EDD <= 10
    'Case nomRCR = 1 Or 2 Or 3
    '               nomRSDD = 0.98
    'Case nomRCR = 4 Or 5 Or 6 Or 7
    '               nomRSDD = 0.97
    'Case nomRCR = 8 Or 9 Or 10
    '               nomRSDD = 0.96
    '
    'End Select
    'End If
    'Return value
    'End Function
End Module
