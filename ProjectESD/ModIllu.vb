Module ModIllu
    Public EDD As Decimal
    Public Function LLD(LampType As String) As Decimal
        'Gaining values for LLD
        Dim value

        If LampType = "General Service Incandescent" Or LampType = "Metal Halide" Then
            value = 80
        ElseIf LampType = "Halogen Incandescent" Or LampType = "High Pressure Sodium" Then
            value = 90
        ElseIf LampType = "Light Loading Fluorescent" Then
            value = 84
        ElseIf LampType = "Medium Loading Fluorescent" Then
            value = 79
        ElseIf LampType = "Heavy Loading Fluorescent" Then
            value = 72
        Else
            value = 83
        End If
        Return value
    End Function
    Public Function LDDRating(LumCat As Integer, Atmosphere As String, Frequency As Decimal) As Decimal
        'Gaining values for LDD
        Dim LDD, a, b, t As Decimal
        t = Frequency / 12

        'Assigning values for constants a and b as in the table
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

        'Basic calculation for LDD and %EDD
        LDD = Math.Round(100 * Math.E ^ (-a * (t ^ b)), 4)

        Return LDD
    End Function
    Public Function TrueRSDD(TrueRCR As Decimal, LDD As Decimal, LDT As String) As Decimal
        'Gaining values for RSDD
        Dim RCRUp, RCRDown, RSDD, EDD As Decimal
        Dim i, n As Integer
        Dim nomRSDD(10) As Decimal

        EDD = (100 - LDD)

        'Assigning %EDD group to constant n as in the table
        If EDD <= 10 Then
            n = 1
        ElseIf EDD > 10 And EDD <= 20 Then
            n = 2
        ElseIf EDD > 20 And EDD <= 30 Then
            n = 3
        ElseIf EDD > 30 Then
            n = 4
        End If

        'Listing nominal RSDD values to array
        If LDT = "Direct" Then
            Select Case n
                Case 1
                    nomRSDD(1) = 0.98
                    nomRSDD(2) = 0.98
                    nomRSDD(3) = 0.98
                    nomRSDD(4) = 0.97
                    nomRSDD(5) = 0.97
                    nomRSDD(6) = 0.97
                    nomRSDD(7) = 0.97
                    nomRSDD(8) = 0.96
                    nomRSDD(9) = 0.96
                    nomRSDD(10) = 0.96
                Case 2
                    nomRSDD(1) = 0.96
                    nomRSDD(2) = 0.96
                    nomRSDD(3) = 0.95
                    nomRSDD(4) = 0.95
                    nomRSDD(5) = 0.94
                    nomRSDD(6) = 0.94
                    nomRSDD(7) = 0.94
                    nomRSDD(8) = 0.93
                    nomRSDD(9) = 0.92
                    nomRSDD(10) = 0.92
                Case 3
                    nomRSDD(1) = 0.94
                    nomRSDD(2) = 0.94
                    nomRSDD(3) = 0.93
                    nomRSDD(4) = 0.92
                    nomRSDD(5) = 0.91
                    nomRSDD(6) = 0.91
                    nomRSDD(7) = 0.9
                    nomRSDD(8) = 0.89
                    nomRSDD(9) = 0.88
                    nomRSDD(10) = 0.87
                Case 4
                    nomRSDD(1) = 0.92
                    nomRSDD(2) = 0.92
                    nomRSDD(3) = 0.9
                    nomRSDD(4) = 0.9
                    nomRSDD(5) = 0.89
                    nomRSDD(6) = 0.88
                    nomRSDD(7) = 0.87
                    nomRSDD(8) = 0.86
                    nomRSDD(9) = 0.85
                    nomRSDD(10) = 0.83
            End Select
        ElseIf LDT = "Semi-Direct" Then
            Select Case n
                Case 1
                    nomRSDD(1) = 0.97
                    nomRSDD(2) = 0.96
                    nomRSDD(3) = 0.96
                    nomRSDD(4) = 0.95
                    nomRSDD(5) = 0.94
                    nomRSDD(6) = 0.94
                    nomRSDD(7) = 0.93
                    nomRSDD(8) = 0.93
                    nomRSDD(9) = 0.93
                    nomRSDD(10) = 0.93
                Case 2
                    nomRSDD(1) = 0.92
                    nomRSDD(2) = 0.92
                    nomRSDD(3) = 0.91
                    nomRSDD(4) = 0.9
                    nomRSDD(5) = 0.9
                    nomRSDD(6) = 0.89
                    nomRSDD(7) = 0.88
                    nomRSDD(8) = 0.87
                    nomRSDD(9) = 0.87
                    nomRSDD(10) = 0.86
                Case 3
                    nomRSDD(1) = 0.89
                    nomRSDD(2) = 0.88
                    nomRSDD(3) = 0.87
                    nomRSDD(4) = 0.85
                    nomRSDD(5) = 0.84
                    nomRSDD(6) = 0.83
                    nomRSDD(7) = 0.82
                    nomRSDD(8) = 0.81
                    nomRSDD(9) = 0.8
                    nomRSDD(10) = 0.79
                Case 4
                    nomRSDD(1) = 0.84
                    nomRSDD(2) = 0.83
                    nomRSDD(3) = 0.82
                    nomRSDD(4) = 0.8
                    nomRSDD(5) = 0.79
                    nomRSDD(6) = 0.78
                    nomRSDD(7) = 0.77
                    nomRSDD(8) = 0.75
                    nomRSDD(9) = 0.74
                    nomRSDD(10) = 0.72
            End Select
        ElseIf LDT = "Direct-Indirect" Then
            Select Case n
                Case 1
                    nomRSDD(1) = 0.94
                    nomRSDD(2) = 0.94
                    nomRSDD(3) = 0.94
                    nomRSDD(4) = 0.94
                    nomRSDD(5) = 0.93
                    nomRSDD(6) = 0.93
                    nomRSDD(7) = 0.93
                    nomRSDD(8) = 0.93
                    nomRSDD(9) = 0.93
                    nomRSDD(10) = 0.93
                Case 2
                    nomRSDD(1) = 0.87
                    nomRSDD(2) = 0.87
                    nomRSDD(3) = 0.86
                    nomRSDD(4) = 0.86
                    nomRSDD(5) = 0.86
                    nomRSDD(6) = 0.85
                    nomRSDD(7) = 0.84
                    nomRSDD(8) = 0.84
                    nomRSDD(9) = 0.84
                    nomRSDD(10) = 0.84
                Case 3
                    nomRSDD(1) = 0.8
                    nomRSDD(2) = 0.8
                    nomRSDD(3) = 0.79
                    nomRSDD(4) = 0.79
                    nomRSDD(5) = 0.78
                    nomRSDD(6) = 0.78
                    nomRSDD(7) = 0.77
                    nomRSDD(8) = 0.76
                    nomRSDD(9) = 0.76
                    nomRSDD(10) = 0.75
                Case 4
                    nomRSDD(1) = 0.76
                    nomRSDD(2) = 0.75
                    nomRSDD(3) = 0.74
                    nomRSDD(4) = 0.73
                    nomRSDD(5) = 0.72
                    nomRSDD(6) = 0.71
                    nomRSDD(7) = 0.7
                    nomRSDD(8) = 0.69
                    nomRSDD(9) = 0.68
                    nomRSDD(10) = 0.67
            End Select
        ElseIf LDT = "Semi-Indirect" Then
            Select Case n
                Case 1
                    nomRSDD(1) = 0.94
                    nomRSDD(2) = 0.94
                    nomRSDD(3) = 0.94
                    nomRSDD(4) = 0.94
                    nomRSDD(5) = 0.93
                    nomRSDD(6) = 0.93
                    nomRSDD(7) = 0.93
                    nomRSDD(8) = 0.93
                    nomRSDD(9) = 0.93
                    nomRSDD(10) = 0.92
                Case 2
                    nomRSDD(1) = 0.87
                    nomRSDD(2) = 0.87
                    nomRSDD(3) = 0.86
                    nomRSDD(4) = 0.86
                    nomRSDD(5) = 0.86
                    nomRSDD(6) = 0.85
                    nomRSDD(7) = 0.84
                    nomRSDD(8) = 0.84
                    nomRSDD(9) = 0.84
                    nomRSDD(10) = 0.83
                Case 3
                    nomRSDD(1) = 0.8
                    nomRSDD(2) = 0.79
                    nomRSDD(3) = 0.78
                    nomRSDD(4) = 0.78
                    nomRSDD(5) = 0.77
                    nomRSDD(6) = 0.76
                    nomRSDD(7) = 0.76
                    nomRSDD(8) = 0.76
                    nomRSDD(9) = 0.75
                    nomRSDD(10) = 0.75
                Case 4
                    nomRSDD(1) = 0.73
                    nomRSDD(2) = 0.72
                    nomRSDD(3) = 0.71
                    nomRSDD(4) = 0.7
                    nomRSDD(5) = 0.69
                    nomRSDD(6) = 0.68
                    nomRSDD(7) = 0.68
                    nomRSDD(8) = 0.68
                    nomRSDD(9) = 0.67
                    nomRSDD(10) = 0.67
            End Select
        ElseIf LDT = "Indirect" Then
            Select Case n
                Case 1
                    nomRSDD(1) = 0.9
                    nomRSDD(2) = 0.9
                    nomRSDD(3) = 0.9
                    nomRSDD(4) = 0.89
                    nomRSDD(5) = 0.89
                    nomRSDD(6) = 0.89
                    nomRSDD(7) = 0.89
                    nomRSDD(8) = 0.88
                    nomRSDD(9) = 0.88
                    nomRSDD(10) = 0.88
                Case 2
                    nomRSDD(1) = 0.8
                    nomRSDD(2) = 0.8
                    nomRSDD(3) = 0.79
                    nomRSDD(4) = 0.78
                    nomRSDD(5) = 0.78
                    nomRSDD(6) = 0.77
                    nomRSDD(7) = 0.76
                    nomRSDD(8) = 0.76
                    nomRSDD(9) = 0.75
                    nomRSDD(10) = 0.75
                Case 3
                    nomRSDD(1) = 0.7
                    nomRSDD(2) = 0.69
                    nomRSDD(3) = 0.68
                    nomRSDD(4) = 0.67
                    nomRSDD(5) = 0.66
                    nomRSDD(6) = 0.66
                    nomRSDD(7) = 0.65
                    nomRSDD(8) = 0.64
                    nomRSDD(9) = 0.63
                    nomRSDD(10) = 0.62
                Case 4
                    nomRSDD(1) = 0.6
                    nomRSDD(2) = 0.59
                    nomRSDD(3) = 0.58
                    nomRSDD(4) = 0.56
                    nomRSDD(5) = 0.55
                    nomRSDD(6) = 0.54
                    nomRSDD(7) = 0.53
                    nomRSDD(8) = 0.52
                    nomRSDD(9) = 0.51
                    nomRSDD(10) = 0.5
            End Select
        End If

        'Selecting value/s of RSDD from array listed above

        If TrueRCR <= 10 Or TrueRCR >= 1 Then
            i = Math.Round(TrueRCR)
            If i = TrueRCR Then
                RSDD = nomRSDD(i)
                GoTo ReturnLine
            Else
                If i > TrueRCR Then
                    RCRUp = i
                    RCRDown = i - 1
                ElseIf i < TrueRCR Then
                    RCRDown = i
                    RCRUp = i + 1
                End If
            End If
        ElseIf TrueRCR > 10 Then
            RCRDown = 9
            RCRUp = 10
        ElseIf TrueRCR < 1 Then
            RCRDown = 1
            RCRUp = 2
        End If
        RSDD = (((nomRSDD(RCRUp) - nomRSDD(RCRDown)) * (TrueRCR - RCRDown)) / (RCRUp - RCRDown)) + nomRSDD(RCRDown)
ReturnLine: Return RSDD
    End Function

    Public Function NomER(CavityRatio As Integer, BaseReflectance As Integer, WallReflectance As Integer) As Integer
        Dim nomEffRef(50) As Integer

        If BaseReflectance = 90 Then
            Select Case WallReflectance
                Case 90
                    nomEffRef = {90, 90, 89, 89, 88, 88, 88, 88, 87, 87, 86, 86, 86, 85, 85, 85, 85, 84, 84, 84, 83, 83, 83, 83, 82, 82,
                        82, 82, 81, 81, 81, 80, 80, 80, 80, 79, 79, 79, 79, 78, 78, 78, 78, 78, 77, 77, 77, 77, 76, 76, 76}
                Case 70
                    nomEffRef = {90, 89, 88, 87, 86, 85, 84, 83, 82, 81, 80, 79, 78, 78, 77, 76, 75, 74, 73, 73, 72, 71, 70, 69, 68, 68,
                        67, 66, 66, 65, 64, 64, 63, 62, 62, 61, 60, 60, 59, 59, 58, 57, 57, 56, 56, 55, 55, 54, 54, 53, 53}
                Case 50
                    nomEffRef = {90, 88, 86, 85, 83, 81, 80, 78, 77, 76, 74, 73, 72, 70, 69, 68, 66, 65, 64, 63, 62, 61, 60, 59, 58, 57,
                        56, 55, 54, 53, 52, 51, 50, 49, 48, 48, 47, 46, 45, 45, 44, 43, 43, 42, 41, 41, 40, 40, 39, 38, 38}
                Case 30
                    nomEffRef = {90, 87, 85, 83, 81, 78, 76, 74, 73, 71, 69, 67, 65, 64, 62, 61, 59, 58, 56, 55, 53, 52, 51, 50, 48, 47,
                        46, 45, 44, 43, 42, 41, 40, 39, 38, 37, 36, 35, 35, 34, 33, 32, 32, 31, 30, 30, 29, 29, 28, 28, 27}
            End Select
        ElseIf BaseReflectance = 80 Then
            Select Case WallReflectance
                Case 80
                    nomEffRef = {80, 79, 79, 78, 78, 77, 77, 76, 75, 75, 74, 74, 73, 73, 72, 72, 71, 71, 70, 70, 69, 69, 68, 68, 67, 67,
                        66, 66, 66, 65, 65, 64, 64, 64, 63, 63, 62, 62, 62, 61, 61, 60, 60, 60, 59, 59, 59, 58, 58, 58, 57}
                Case 70
                    nomEffRef = {80, 79, 78, 77, 76, 75, 75, 74, 73, 72, 71, 71, 70, 69, 68, 68, 67, 66, 65, 65, 64, 63, 63, 62, 61, 61,
                        60, 60, 59, 58, 58, 57, 57, 56, 56, 55, 54, 54, 53, 53, 52, 52, 51, 51, 51, 50, 50, 49, 49, 49, 48}
                Case 50
                    nomEffRef = {80, 78, 77, 75, 74, 73, 71, 70, 69, 68, 66, 65, 64, 63, 62, 61, 60, 59, 58, 57, 56, 55, 54, 53, 52, 51,
                        50, 49, 48, 48, 47, 46, 45, 44, 44, 43, 42, 42, 41, 40, 40, 39, 39, 38, 38, 37, 37, 36, 36, 35, 35}
                Case 30
                    nomEffRef = {80, 78, 76, 74, 72, 70, 68, 66, 65, 63, 61, 60, 58, 57, 55, 54, 53, 52, 50, 49, 48, 47, 45, 44, 43, 42,
                        41, 40, 39, 38, 38, 37, 36, 35, 34, 33, 33, 32, 31, 30, 30, 29, 29, 28, 28, 27, 26, 26, 25, 25, 25}
            End Select
        ElseIf BaseReflectance = 70 Then
            Select Case WallReflectance
                Case 70
                    nomEffRef = {70, 69, 68, 68, 67, 66, 65, 65, 64, 63, 63, 62, 61, 61, 60, 59, 59, 58, 57, 57, 56, 55, 54, 53, 52, 51,
                        50, 49, 48, 48, 47, 46, 45, 44, 44, 43, 42, 42, 41, 40, 40, 39, 39, 38, 38, 37, 37, 36, 36, 35, 35}
                Case 50
                    nomEffRef = {70, 69, 67, 66, 65, 64, 62, 61, 60, 59, 58, 57, 56, 55, 54, 53, 53, 51, 50, 49, 48, 47, 46, 46, 45, 44,
                        43, 43, 42, 41, 40, 40, 39, 39, 38, 38, 37, 37, 36, 36, 35, 35, 34, 34, 34, 33, 33, 33, 32, 32, 32}
                Case 30
                    nomEffRef = {70, 68, 66, 64, 63, 61, 59, 58, 56, 55, 53, 52, 50, 49, 48, 47, 45, 44, 43, 42, 41, 40, 39, 38, 37, 36,
                        35, 34, 33, 33, 32, 31, 30, 30, 29, 29, 28, 27, 27, 26, 26, 25, 25, 25, 24, 24, 24, 23, 23, 23, 22}
            End Select
        ElseIf BaseReflectance = 50 Then
            Select Case WallReflectance
                Case 70
                    nomEffRef = {50, 49, 49, 49, 48, 48, 47, 47, 47, 46, 46, 46, 45, 45, 45, 44, 44, 44, 43, 43, 43, 43, 42, 42, 42, 41,
                        41, 41, 41, 40, 40, 40, 40, 39, 39, 39, 39, 38, 38, 38, 38, 37, 37, 37, 37, 37, 36, 36, 36, 36, 36}
                Case 50
                    nomEffRef = {50, 49, 48, 47, 46, 46, 45, 44, 43, 43, 42, 41, 41, 40, 40, 39, 39, 38, 37, 37, 37, 36, 36, 35, 35, 34,
                        34, 33, 33, 33, 32, 32, 31, 31, 31, 30, 30, 30, 29, 29, 29, 28, 28, 28, 27, 27, 27, 26, 26, 26, 26}
                Case 30
                    nomEffRef = {50, 48, 47, 46, 45, 44, 43, 42, 41, 40, 39, 38, 37, 36, 35, 34, 33, 32, 32, 31, 30, 29, 29, 28, 28, 27,
                        26, 26, 25, 25, 24, 24, 23, 23, 22, 22, 21, 21, 21, 20, 20, 20, 19, 19, 19, 19, 18, 18, 18, 18, 17}
            End Select
        ElseIf BaseReflectance = 30 Then
            Select Case WallReflectance
                Case 65

                Case 50

                Case 30

                Case 10

            End Select
        End If

        Return nomEffRef(CavityRatio)
    End Function
End Module
