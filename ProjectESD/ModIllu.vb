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

        EDD = 100 - LDD

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

        If TrueRCR <= 10 And TrueRCR >= 1 Then
            i = Math.Truncate(TrueRCR)
            If i = TrueRCR Then
                RSDD = nomRSDD(i)
                GoTo ReturnLine
            Else
                RCRDown = i
                RCRUp = i + 1
            End If
        ElseIf TrueRCR > 10 Then
            RCRDown = 9
            RCRUp = 10
        ElseIf TrueRCR < 1 Then
            RCRDown = 1
            RCRUp = 2
        End If
        RSDD = ((nomRSDD(RCRUp) - nomRSDD(RCRDown)) * (TrueRCR - RCRDown) / (RCRUp - RCRDown)) + nomRSDD(RCRDown)
ReturnLine: Return RSDD
    End Function

#Region "Effective Reflectance"
    Private Function NomER(CavityRatio As Integer, BaseReflectance As Integer, WallReflectance As Integer) As Integer
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
                    nomEffRef = {30, 30, 30, 30, 30, 29, 29, 29, 29, 29, 29, 29, 29, 29, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 27,
                        27, 27, 27, 27, 27, 27, 27, 27, 27, 26, 26, 26, 26, 26, 26, 26, 26, 26, 26, 25, 25, 25, 25, 25, 25}
                Case 50
                    nomEffRef = {30, 30, 29, 29, 29, 28, 28, 28, 27, 27, 27, 26, 26, 26, 26, 25, 25, 25, 25, 25, 24, 24, 24, 24, 24, 23,
                        23, 23, 23, 23, 22, 22, 22, 22, 22, 22, 21, 21, 21, 21, 21, 21, 20, 20, 20, 20, 20, 20, 19, 19, 19}
                Case 30
                    nomEffRef = {30, 29, 29, 28, 27, 27, 26, 26, 25, 25, 24, 24, 23, 23, 22, 22, 21, 21, 21, 20, 20, 20, 19, 19, 19, 18,
                        18, 18, 18, 17, 17, 17, 16, 16, 16, 16, 15, 15, 15, 15, 15, 14, 14, 14, 14, 14, 14, 13, 13, 13, 13}
                Case 10
                    nomEffRef = {30, 29, 28, 27, 26, 25, 25, 24, 23, 22, 22, 21, 20, 20, 19, 18, 18, 17, 17, 16, 16, 16, 15, 15, 14, 14,
                        13, 13, 13, 12, 12, 12, 11, 11, 11, 11, 10, 10, 10, 10, 9, 9, 9, 8, 8, 8, 8, 8, 8, 7, 7}
            End Select
        ElseIf BaseReflectance = 10 Then
            Select Case WallReflectance
                Case 60
                    nomEffRef = {10, 10, 10, 10, 11, 11, 11, 11, 11, 11, 11, 11, 12, 12, 12, 12, 12, 12, 12, 12, 12, 13, 13, 13, 13, 13,
                        13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 14, 14, 14, 14, 14, 14}
                Case 50
                    nomEffRef(0) = 10
                    nomEffRef(1) = 10
                    nomEffRef(2) = 10
                    nomEffRef(4) = 10
                    nomEffRef(6) = 11
                    nomEffRef(8) = 11
                    nomEffRef(10) = 11
                    nomEffRef(12) = 12
                    nomEffRef(14) = 12
                    nomEffRef(16) = 12
                    nomEffRef(18) = 12
                    nomEffRef(20) = 12
                Case 30
                    nomEffRef(0) = 10
                    nomEffRef(1) = 10
                    nomEffRef(2) = 10
                    nomEffRef(4) = 10
                    nomEffRef(6) = 10
                    nomEffRef(8) = 10
                    nomEffRef(10) = 9
                    nomEffRef(12) = 9
                    nomEffRef(14) = 9
                    nomEffRef(16) = 9
                    nomEffRef(18) = 9
                    nomEffRef(20) = 9
                Case 10
                    nomEffRef(0) = 10
                    nomEffRef(1) = 9
                    nomEffRef(2) = 9
                    nomEffRef(4) = 9
                    nomEffRef(6) = 9
                    nomEffRef(8) = 8
                    nomEffRef(10) = 8
                    nomEffRef(12) = 7
                    nomEffRef(14) = 7
                    nomEffRef(16) = 7
                    nomEffRef(18) = 6
                    nomEffRef(20) = 6
            End Select
        End If

        Return nomEffRef(CavityRatio)
    End Function

    Public Function TrueER(CavityRatio As Decimal, BaseReflectance As Integer, WallReflectance As Integer) As Decimal
        Dim CRFactor As Decimal = CavityRatio * 10
        'm and n for cavity ratio boundaries, a and b for base reflectance, i, j, k, and l for wall reflectance
        Dim c, x, w As Decimal
        Dim m, n, a, b, i, j, k, l As Integer
        Dim nomEffRef1, nomEffRef2, nomEffRef3, nomEffRef4, nomEffRef5, nomEffRef6, nomEffRef7, nomEffRef8 As Integer
        Dim CorRef1, CorRef2, CorRef3, CorRef4, FinalRef1, FinalRef2 As Decimal
        Dim TrueEffRef As Decimal

        c = Math.Truncate(CRFactor)
        x = BaseReflectance
        w = WallReflectance

        If BaseReflectance < 30 Then
            If (c > 0 And c < 1) Or (c > 1 And c < 2) Then
                a = c
                b = c + 1
            ElseIf c > 2 And c < 4 Then
                a = 2
                b = 4
            ElseIf c > 4 And c < 6 Then
                a = 4
                b = 6
            ElseIf c > 6 And c < 8 Then
                a = 6
                b = 8
            ElseIf c > 8 And c < 10 Then
                a = 8
                b = 10
            ElseIf c > 10 And c < 12 Then
                a = 10
                b = 12
            ElseIf c > 12 And c < 14 Then
                a = 12
                b = 14
            ElseIf c > 14 And c < 16 Then
                a = 14
                b = 16
            ElseIf c > 16 And c < 18 Then
                a = 16
                b = 18
            ElseIf (c > 18 And c < 20) Or c > 20 Then
                a = 18
                b = 20
            Else
                a = c
                b = c
            End If
        Else
            If c = CRFactor Then
                a = c
                b = c
            Else
                a = c
                b = c + 1
            End If
        End If

        If BaseReflectance = 90 Then
            m = BaseReflectance
            n = BaseReflectance
            If (WallReflectance < 90 And WallReflectance > 70) Or WallReflectance > 90 Then
                i = 90
                j = 70
            ElseIf WallReflectance < 70 And WallReflectance > 50 Then
                i = 70
                j = 50
            ElseIf (WallReflectance < 50 And WallReflectance > 30) Or WallReflectance < 30 Then
                i = 50
                j = 30
            Else
                i = WallReflectance
                j = WallReflectance
            End If
            k = i
            l = j
        ElseIf BaseReflectance = 80 Then
            m = BaseReflectance
            n = BaseReflectance
            If (WallReflectance < 80 And WallReflectance > 70) Or WallReflectance > 80 Then
                i = 80
                j = 70
            ElseIf WallReflectance < 70 And WallReflectance > 50 Then
                i = 70
                j = 50
            ElseIf (WallReflectance < 50 And WallReflectance > 30) Or WallReflectance < 30 Then
                i = 50
                j = 30
            Else
                i = WallReflectance
                j = WallReflectance
            End If
            k = i
            l = j
        ElseIf BaseReflectance = 70 Or BaseReflectance = 50 Then
            m = BaseReflectance
            n = BaseReflectance
            If (WallReflectance < 70 And WallReflectance > 50) Or WallReflectance > 70 Then
                i = 70
                j = 50
            ElseIf (WallReflectance < 50 And WallReflectance > 30) Or WallReflectance < 50 Then
                i = 50
                j = 30
            Else
                i = WallReflectance
                j = WallReflectance
            End If
            k = i
            l = j
        ElseIf BaseReflectance = 30 Then
            m = BaseReflectance
            n = BaseReflectance
            If (WallReflectance < 65 And WallReflectance > 50) Or WallReflectance > 65 Then
                i = 65
                j = 50
            ElseIf WallReflectance < 50 And WallReflectance > 30 Then
                i = 50
                j = 30
            ElseIf (WallReflectance < 30 And WallReflectance > 10) Or WallReflectance < 10 Then
                i = 30
                j = 10
            Else
                i = WallReflectance
                j = WallReflectance
            End If
            k = i
            l = j
        ElseIf BaseReflectance = 10 Then
            m = BaseReflectance
            n = BaseReflectance
            If (WallReflectance < 60 And WallReflectance > 50) Or WallReflectance > 60 Then
                i = 60
                j = 50
            ElseIf WallReflectance < 50 And WallReflectance > 30 Then
                i = 50
                j = 30
            ElseIf (WallReflectance < 30 And WallReflectance > 10) Or WallReflectance < 10 Then
                i = 30
                j = 10
            Else
                i = WallReflectance
                j = WallReflectance
            End If
            k = i
            l = j
        Else
            If (BaseReflectance < 90 And BaseReflectance > 80) Or BaseReflectance > 90 Then
                m = 90
                n = 80
                'for m
                If (WallReflectance < 90 And WallReflectance > 70) Or WallReflectance > 90 Then
                    i = 90
                    j = 70
                ElseIf WallReflectance < 70 And WallReflectance > 50 Then
                    i = 70
                    j = 50
                ElseIf (WallReflectance < 50 And WallReflectance > 30) Or WallReflectance < 30 Then
                    i = 50
                    j = 30
                Else
                    i = WallReflectance
                    j = WallReflectance
                End If

                'for n
                If (WallReflectance < 80 And WallReflectance > 70) Or WallReflectance > 80 Then
                    k = 80
                    l = 70
                ElseIf WallReflectance < 70 And WallReflectance > 50 Then
                    k = 70
                    l = 50
                ElseIf (WallReflectance < 50 And WallReflectance > 30) Or WallReflectance < 30 Then
                    k = 50
                    l = 30
                Else
                    k = WallReflectance
                    l = WallReflectance
                End If

            ElseIf BaseReflectance < 80 And BaseReflectance > 70 Then
                m = 80
                n = 70
                If (WallReflectance < 80 And WallReflectance > 70) Or WallReflectance > 80 Then
                    i = 80
                    j = 70
                ElseIf WallReflectance < 70 And WallReflectance > 50 Then
                    i = 70
                    j = 50
                ElseIf (WallReflectance < 50 And WallReflectance > 30) Or WallReflectance < 30 Then
                    i = 50
                    j = 30
                Else
                    i = WallReflectance
                    j = WallReflectance
                End If

                If (WallReflectance < 70 And WallReflectance > 50) Or WallReflectance > 70 Then
                    k = 70
                    l = 50
                ElseIf (WallReflectance < 50 And WallReflectance > 30) Or WallReflectance < 30 Then
                    k = 50
                    l = 30
                Else
                    k = WallReflectance
                    l = WallReflectance
                End If
            ElseIf BaseReflectance < 70 And BaseReflectance > 50 Then
                m = 70
                n = 50
                If (WallReflectance < 70 And WallReflectance > 50) Or WallReflectance > 70 Then
                    i = 70
                    j = 50
                ElseIf (WallReflectance < 50 And WallReflectance > 30) Or WallReflectance < 30 Then
                    i = 50
                    j = 30
                Else
                    i = WallReflectance
                    j = WallReflectance
                End If
                k = i
                l = j

            ElseIf BaseReflectance < 50 And BaseReflectance > 30 Then
                m = 50
                n = 30
                If (WallReflectance < 70 And WallReflectance > 50) Or WallReflectance > 70 Then
                    i = 70
                    j = 50
                ElseIf (WallReflectance < 50 And WallReflectance > 30) Or WallReflectance < 30 Then
                    i = 50
                    j = 30
                Else
                    i = WallReflectance
                    j = WallReflectance
                End If

                If (WallReflectance < 65 And WallReflectance > 50) Or WallReflectance > 65 Then
                    k = 65
                    l = 50
                ElseIf WallReflectance < 50 And WallReflectance > 30 Then
                    k = 50
                    l = 30
                ElseIf (WallReflectance < 30 And WallReflectance > 10) Or WallReflectance < 10 Then
                    k = 30
                    l = 10
                Else
                    k = WallReflectance
                    l = WallReflectance
                End If
            ElseIf (BaseReflectance < 30 And BaseReflectance > 10) Or BaseReflectance < 10 Then
                m = 30
                n = 10
                If (WallReflectance < 65 And WallReflectance > 50) Or WallReflectance > 65 Then
                    i = 65
                    j = 50
                ElseIf WallReflectance < 50 And WallReflectance > 30 Then
                    i = 50
                    j = 30
                ElseIf (WallReflectance < 30 And WallReflectance > 10) Or WallReflectance < 10 Then
                    i = 30
                    j = 10
                Else
                    i = WallReflectance
                    j = WallReflectance
                End If

                If (WallReflectance < 60 And WallReflectance > 50) Or WallReflectance > 60 Then
                    k = 60
                    l = 50
                ElseIf WallReflectance < 50 And WallReflectance > 30 Then
                    k = 50
                    l = 30
                ElseIf (WallReflectance < 30 And WallReflectance > 10) Or WallReflectance < 10 Then
                    k = 30
                    l = 10
                Else
                    k = WallReflectance
                    l = WallReflectance
                End If
            End If
        End If

        If CavityRatio = 0 Then
            TrueEffRef = BaseReflectance
            GoTo ReturnLine
        Else
            nomEffRef1 = NomER(a, m, i)
            nomEffRef2 = NomER(a, m, j)
            nomEffRef3 = NomER(a, n, k)
            nomEffRef4 = NomER(a, n, l)
            nomEffRef5 = NomER(b, m, i)
            nomEffRef6 = NomER(b, m, j)
            nomEffRef7 = NomER(b, n, k)
            nomEffRef8 = NomER(b, n, l)

            If nomEffRef1 = nomEffRef2 Then
                CorRef1 = nomEffRef1
            Else
                CorRef1 = Interpolate(i, j, nomEffRef1, nomEffRef2, WallReflectance)
            End If

            If nomEffRef3 = nomEffRef4 Then
                CorRef2 = nomEffRef3
            Else
                CorRef2 = Interpolate(k, l, nomEffRef3, nomEffRef4, WallReflectance)
            End If

            If nomEffRef5 = nomEffRef6 Then
                CorRef3 = nomEffRef5
            Else
                CorRef3 = Interpolate(i, j, nomEffRef5, nomEffRef6, WallReflectance)
            End If

            If nomEffRef7 = nomEffRef8 Then
                CorRef4 = nomEffRef7
            Else
                CorRef4 = Interpolate(k, l, nomEffRef7, nomEffRef8, WallReflectance)
            End If

            If CorRef1 = CorRef2 Then
                FinalRef1 = CorRef1
            Else
                FinalRef1 = Interpolate(m, n, CorRef1, CorRef2, BaseReflectance)
            End If

            If CorRef3 = CorRef4 Then
                FinalRef2 = CorRef3
            Else
                FinalRef2 = Interpolate(m, n, CorRef3, CorRef4, BaseReflectance)
            End If

            If FinalRef1 = FinalRef2 Then
                TrueEffRef = FinalRef1
            Else
                TrueEffRef = Interpolate(b / 10, a / 10, FinalRef1, FinalRef2, CavityRatio)
            End If
        End If

ReturnLine: Return TrueEffRef
    End Function
#End Region

#Region "Coefficient of Utilization"
    Private Function NominalCU(RCR As Integer, WallReflectance As Integer, EffectiveReflectance As Integer) As Decimal
        Dim nomCU(10) As Decimal

        If EffectiveReflectance = 80 Then
            Select Case WallReflectance
                Case 70
                    nomCU = {0.81, 0.76, 0.72, 0.68, 0.64, 0.6, 0.56, 0.52, 0.49, 0.45, 0.42}
                Case 50
                    nomCU = {0.81, 0.74, 0.69, 0.63, 0.58, 0.54, 0.49, 0.45, 0.41, 0.37, 0.34}
                Case 30
                    nomCU = {0.81, 0.73, 0.66, 0.6, 0.54, 0.49, 0.45, 0.41, 0.37, 0.33, 0.29}
                Case 10
                    nomCU = {0.81, 0.71, 0.63, 0.57, 0.51, 0.46, 0.42, 0.37, 0.33, 0.29, 0.26}
            End Select
        ElseIf EffectiveReflectance = 70 Then
            Select Case WallReflectance
                Case 70
                    nomCU = {0.79, 0.75, 0.71, 0.67, 0.63, 0.59, 0.55, 0.52, 0.48, 0.44, 0.41}
                Case 50
                    nomCU = {0.79, 0.73, 0.68, 0.62, 0.58, 0.53, 0.49, 0.45, 0.41, 0.37, 0.34}
                Case 30
                    nomCU = {0.79, 0.71, 0.65, 0.59, 0.54, 0.49, 0.45, 0.4, 0.36, 0.32, 0.29}
                Case 10
                    nomCU = {0.79, 0.7, 0.63, 0.56, 0.51, 0.46, 0.42, 0.37, 0.33, 0.29, 0.26}
            End Select
        ElseIf EffectiveReflectance = 50 Then
            Select Case WallReflectance
                Case 50
                    nomCU = {0.75, 0.7, 0.65, 0.61, 0.56, 0.52, 0.48, 0.44, 0.4, 0.36, 0.33}
                Case 30
                    nomCU = {0.75, 0.69, 0.63, 0.58, 0.53, 0.48, 0.44, 0.4, 0.36, 0.32, 0.29}
                Case 10
                    nomCU = {0.75, 0.68, 0.61, 0.55, 0.5, 0.45, 0.41, 0.37, 0.33, 0.29, 0.26}
            End Select
        ElseIf EffectiveReflectance = 30 Then
            Select Case WallReflectance
                Case 50
                    nomCU = {0.72, 0.68, 0.63, 0.59, 0.55, 0.5, 0.47, 0.43, 0.39, 0.36, 0.32}
                Case 30
                    nomCU = {0.72, 0.67, 0.62, 0.57, 0.52, 0.47, 0.43, 0.4, 0.36, 0.32, 0.29}
                Case 10
                    nomCU = {0.72, 0.66, 0.6, 0.54, 0.49, 0.45, 0.41, 0.37, 0.33, 0.29, 0.26}
            End Select
        ElseIf EffectiveReflectance = 10 Then
            Select Case WallReflectance
                Case 50
                    nomCU = {0.69, 0.65, 0.61, 0.57, 0.53, 0.49, 0.46, 0.42, 0.38, 0.35, 0.32}
                Case 30
                    nomCU = {0.69, 0.65, 0.6, 0.55, 0.51, 0.47, 0.43, 0.39, 0.35, 0.31, 0.28}
                Case 10
                    nomCU = {0.69, 0.64, 0.59, 0.54, 0.49, 0.44, 0.41, 0.37, 0.33, 0.29, 0.26}
            End Select
        End If
        Return nomCU(RCR)
    End Function

    Public Function TrueCU(RCR As Decimal, WallReflectance As Integer, EffectiveReflectance As Decimal) As Decimal
        Dim r As Integer = Math.Truncate(RCR)
        Dim w As Integer = WallReflectance
        Dim e As Decimal = EffectiveReflectance

        Dim m, n, a, b, i, j, k, l As Integer
        Dim nomCU1, nomCU2, nomCU3, nomCU4, nomCU5, nomCU6, nomCU7, nomCU8 As Decimal
        Dim CorRef1, CorRef2, CorRef3, CorRef4, FinalRef1, FinalRef2 As Decimal
        Dim TrueCoU As Decimal

        If RCR <= 10 And RCR >= 1 Then
            If r = RCR Then
                a = r
                b = r
            Else
                a = r
                b = r + 1
            End If
        ElseIf RCR > 10 Then
            a = 9
            b = 10
        ElseIf RCR < 1 Then
            a = 1
            b = 2
        End If

        If (e < 80 And e > 70) Or e > 80 Then
            m = 80
            n = 70

            If (w < 70 And w > 50) Or w > 70 Then
                i = 70
                j = 50
            ElseIf w < 50 And w > 30 Then
                i = 50
                j = 30
            ElseIf (w < 30 And w > 10) Or w < 10 Then
                i = 30
                j = 10
            Else
                i = w
                j = w
            End If

            k = i
            l = j
        ElseIf e < 70 And e > 50 Then
            m = 70
            n = 50

            If (w < 70 And w > 50) Or w > 70 Then
                i = 70
                j = 50
            ElseIf w < 50 And w > 30 Then
                i = 50
                j = 30
            ElseIf (w < 30 And w > 10) Or w < 10 Then
                i = 30
                j = 10
            Else
                i = w
                j = w
            End If

            If (w < 50 And w > 30) Or w > 50 Then
                k = 50
                l = 30
            ElseIf (w < 30 And w > 10) Or w < 10 Then
                k = 30
                l = 10
            Else
                k = w
                l = w
            End If
        ElseIf e < 50 And e > 30 Then
            m = 50
            n = 30

            If (w < 50 And w > 30) Or w > 50 Then
                i = 50
                j = 30
                k = 50
                l = 30
            ElseIf (w < 30 And w > 10) Or w < 10 Then
                i = 30
                j = 10
                k = 30
                l = 10
            Else
                i = w
                j = w
                k = w
                l = w
            End If
        ElseIf (e < 30 And e > 10) Or e < 10 Then
            m = 30
            n = 10

            If (w < 50 And w > 30) Or w > 50 Then
                i = 50
                j = 30
                k = 50
                l = 30
            ElseIf (w < 30 And w > 10) Or w < 10 Then
                i = 30
                j = 10
                k = 30
                l = 10
            Else
                i = w
                j = w
                k = w
                l = w
            End If
        Else
            m = e
            n = e
            If e = 80 Or e = 70 Then
                If (w < 70 And w > 50) Or w > 70 Then
                    i = 70
                    j = 50
                ElseIf w < 50 And w > 30 Then
                    i = 50
                    j = 30
                ElseIf (w < 30 And w > 10) Or w < 10 Then
                    i = 30
                    j = 10
                Else
                    i = w
                    j = w
                End If

                k = i
                l = j
            Else
                If (w < 50 And w > 30) Or w > 50 Then
                    i = 50
                    j = 30
                    k = 50
                    l = 30
                ElseIf (w < 30 And w > 10) Or w < 10 Then
                    i = 30
                    j = 10
                    k = 30
                    l = 10
                Else
                    i = w
                    j = w
                    k = w
                    l = w
                End If
            End If
        End If

        nomCU1 = NominalCU(a, i, m)
        nomCU2 = NominalCU(a, j, m)
        nomCU3 = NominalCU(a, k, n)
        nomCU4 = NominalCU(a, l, n)
        nomCU5 = NominalCU(b, i, m)
        nomCU6 = NominalCU(b, j, m)
        nomCU7 = NominalCU(b, k, n)
        nomCU8 = NominalCU(b, l, n)

        If nomCU1 = nomCU2 Then
            CorRef1 = nomCU1
        Else
            CorRef1 = Interpolate(i, j, nomCU1, nomCU2, w)
        End If

        If nomCU3 = nomCU4 Then
            CorRef2 = nomCU3
        Else
            CorRef2 = Interpolate(k, l, nomCU3, nomCU4, w)
        End If

        If nomCU5 = nomCU6 Then
            CorRef3 = nomCU5
        Else
            CorRef3 = Interpolate(i, j, nomCU5, nomCU6, w)
        End If

        If nomCU7 = nomCU8 Then
            CorRef4 = nomCU7
        Else
            CorRef4 = Interpolate(k, l, nomCU7, nomCU8, w)
        End If

        If CorRef1 = CorRef2 Then
            FinalRef1 = CorRef1
        Else
            FinalRef1 = Interpolate(m, n, CorRef1, CorRef2, e)
        End If

        If CorRef3 = CorRef4 Then
            FinalRef2 = CorRef3
        Else
            FinalRef2 = Interpolate(m, n, CorRef3, CorRef4, e)
        End If

        If FinalRef1 = FinalRef2 Then
            TrueCoU = FinalRef1
        Else
            TrueCoU = Interpolate(b, a, FinalRef2, FinalRef1, RCR)
        End If

        Return TrueCoU
    End Function
#End Region

#Region "Correction Factor"
    Private Function NominalCF(RCR As Integer, CeilingReflectance As Integer, WallReflectance As Integer, FloorReflectance As Integer) As Decimal
        Dim NomCF(10) As Decimal

        Select Case FloorReflectance
            Case 20
                NomCF(RCR) = 1
            Case 30
                If CeilingReflectance = 80 Then
                    Select Case WallReflectance
                        Case 70
                            NomCF = {0, 1.092, 1.079, 1.07, 1.062, 1.056, 1.052, 1.047, 1.044, 1.04, 1.037}
                        Case 50
                            NomCF = {0, 1.082, 1.066, 1.054, 1.045, 1.038, 1.033, 1.029, 1.026, 1.024, 1.022}
                        Case 30
                            NomCF = {0, 1.075, 1.055, 1.042, 1.033, 1.026, 1.021, 1.018, 1.015, 1.014, 1.012}
                        Case 10
                            NomCF = {0, 1.068, 1.047, 1.033, 1.024, 1.018, 1.014, 1.011, 1.009, 1.007, 1.006}
                    End Select
                ElseIf CeilingReflectance = 70 Then
                    Select Case WallReflectance
                        Case 70
                            NomCF = {0, 1.077, 1.068, 1.061, 1.055, 1.05, 1.047, 1.043, 1.04, 1.037, 1.034}
                        Case 50
                            NomCF = {0, 1.07, 1.057, 1.048, 1.04, 1.034, 1.03, 1.026, 1.024, 1.022, 1.02}
                        Case 30
                            NomCF = {0, 1.064, 1.048, 1.037, 1.029, 1.024, 1.02, 1.017, 1.015, 1.014, 1.012}
                        Case 10
                            NomCF = {0, 1.059, 1.039, 1.028, 1.021, 1.015, 1.012, 1.009, 1.007, 1.006, 1.005}
                    End Select
                ElseIf CeilingReflectance = 50 Then
                    Select Case WallReflectance
                        Case 50
                            NomCF = {0, 1.049, 1.041, 1.034, 1.03, 1.027, 1.024, 1.022, 1.02, 1.019, 1.017}
                        Case 30
                            NomCF = {0, 1.044, 1.033, 1.027, 1.022, 1.018, 1.015, 1.013, 1.012, 1.011, 1.01}
                        Case 10
                            NomCF = {0, 1.04, 1.027, 1.02, 1.015, 1.012, 1.009, 1.007, 1.006, 1.005, 1.004}
                    End Select
                ElseIf CeilingReflectance = 30 Then
                    Select Case WallReflectance
                        Case 50
                            NomCF = {0, 1.028, 1.026, 1.024, 1.022, 1.02, 1.019, 1.018, 1.017, 1.016, 1.015}
                        Case 30
                            NomCF = {0, 1.026, 1.021, 1.017, 1.015, 1.013, 1.012, 1.01, 1.009, 1.009, 1.009}
                        Case 10
                            NomCF = {0, 1.023, 1.017, 1.012, 1.01, 1.008, 1.006, 1.005, 1.004, 1.004, 1.003}
                    End Select
                ElseIf CeilingReflectance = 10 Then
                    Select Case WallReflectance
                        Case 50
                            NomCF = {0, 1.012, 1.013, 1.014, 1.014, 1.014, 1.014, 1.014, 1.013, 1.013, 1.013}
                        Case 30
                            NomCF = {0, 1.01, 1.01, 1.009, 1.009, 1.009, 1.008, 1.008, 1.007, 1.007, 1.007}
                        Case 10
                            NomCF = {0, 1.006, 1.006, 1.005, 1.004, 1.004, 1.003, 1.003, 1.003, 1.002, 1.002}
                    End Select
                End If
            Case 10
                If CeilingReflectance = 80 Then
                    Select Case WallReflectance
                        Case 70
                            NomCF = {0, 0.923, 0.931, 0.939, 0.944, 0.949, 0.953, 0.957, 0.96, 0.963, 0.965}
                        Case 50
                            NomCF = {0, 0.929, 0.942, 0.951, 0.958, 0.964, 0.969, 0.973, 0.976, 0.978, 0.98}
                        Case 30
                            NomCF = {0, 0.935, 0.95, 0.961, 0.969, 0.975, 0.98, 0.983, 0.986, 0.987, 0.989}
                        Case 10
                            NomCF = {0, 0.94, 0.958, 0.969, 0.978, 0.983, 0.986, 0.991, 0.993, 0.994, 0.995}
                    End Select
                ElseIf CeilingReflectance = 70 Then
                    Select Case WallReflectance
                        Case 70
                            NomCF = {0, 0.933, 0.94, 0.945, 0.95, 0.954, 0.958, 0.961, 0.963, 0.965, 0.967}
                        Case 50
                            NomCF = {0, 0.939, 0.949, 0.957, 0.963, 0.968, 0.972, 0.975, 0.977, 0.979, 0.981}
                        Case 30
                            NomCF = {0, 0.943, 0.957, 0.966, 0.973, 0.976, 0.982, 0.985, 0.987, 0.989, 0.99}
                        Case 10
                            NomCF = {0, 0.948, 0.963, 0.973, 0.98, 0.985, 0.989, 0.991, 0.993, 0.994, 0.995}
                    End Select
                ElseIf CeilingReflectance = 50 Then
                    Select Case WallReflectance
                        Case 50
                            NomCF = {0, 0.956, 0.962, 0.967, 0.972, 0.975, 0.977, 0.979, 0.981, 0.983, 0.984}
                        Case 30
                            NomCF = {0, 0.96, 0.968, 0.975, 0.98, 0.983, 0.985, 0.987, 0.988, 0.99, 0.991}
                        Case 10
                            NomCF = {0, 0.963, 0.974, 0.981, 0.986, 0.989, 0.992, 0.994, 0.995, 0.996, 0.997}
                    End Select
                ElseIf CeilingReflectance = 30 Then
                    Select Case WallReflectance
                        Case 50
                            NomCF = {0, 0.973, 0.976, 0.978, 0.98, 0.981, 0.982, 0.983, 0.984, 0.985, 0.986}
                        Case 30
                            NomCF = {0, 0.976, 0.98, 0.983, 0.986, 0.988, 0.989, 0.99, 0.991, 0.992, 0.993}
                        Case 10
                            NomCF = {0, 0.979, 0.985, 0.988, 0.991, 0.993, 0.995, 0.996, 0.997, 0.998, 0.998}
                    End Select
                ElseIf CeilingReflectance = 10 Then
                    Select Case WallReflectance
                        Case 50
                            NomCF = {0, 0.989, 0.988, 0.988, 0.987, 0.987, 0.987, 0.987, 0.987, 0.986, 0.986}
                        Case 30
                            NomCF = {0, 0.991, 0.991, 0.992, 0.992, 0.992, 0.993, 0.993, 0.994, 0.994, 0.994}
                        Case 10
                            NomCF = {0, 0.993, 0.995, 0.996, 0.996, 0.997, 0.997, 0.998, 0.998, 0.999, 0.999}
                    End Select
                End If
            Case 0
                If CeilingReflectance = 80 Then
                    Select Case WallReflectance
                        Case 70
                            NomCF = {0, 0.859, 0.871, 0.882, 0.893, 0.903, 0.911, 0.917, 0.922, 0.928, 0.933}
                        Case 50
                            NomCF = {0, 0.87, 0.887, 0.904, 0.919, 0.931, 0.94, 0.947, 0.953, 0.958, 0.962}
                        Case 30
                            NomCF = {0, 0.879, 0.903, 0.915, 0.941, 0.953, 0.961, 0.967, 0.971, 0.975, 0.979}
                        Case 10
                            NomCF = {0, 0.886, 0.919, 0.942, 0.958, 0.969, 0.976, 0.981, 0.985, 0.988, 0.991}
                    End Select
                ElseIf CeilingReflectance = 70 Then
                    Select Case WallReflectance
                        Case 70
                            NomCF = {0, 0.873, 0.886, 0.898, 0.908, 0.914, 0.92, 0.924, 0.929, 0.933, 0.937}
                        Case 50
                            NomCF = {0, 0.884, 0.902, 0.918, 0.93, 0.938, 0.945, 0.95, 0.955, 0.959, 0.963}
                        Case 30
                            NomCF = {0, 0.893, 0.916, 0.934, 0.946, 0.958, 0.965, 0.97, 0.975, 0.98, 0.983}
                        Case 10
                            NomCF = {0, 0.9014, 0.928, 0.947, 0.961, 0.97, 0.977, 0.982, 0.986, 0.989, 0.992}
                    End Select
                ElseIf CeilingReflectance = 50 Then
                    Select Case WallReflectance
                        Case 50
                            NomCF = {0, 0.916, 0.925, 0.936, 0.945, 0.951, 0.955, 0.959, 0.963, 0.966, 0.969}
                        Case 30
                            NomCF = {0, 0.923, 0.938, 0.95, 0.961, 0.967, 0.972, 0.975, 0.978, 0.98, 0.982}
                        Case 10
                            NomCF = {0, 0.929, 0.949, 0.964, 0.974, 0.98, 0.985, 0.988, 0.991, 0.993, 0.995}
                    End Select
                ElseIf CeilingReflectance = 30 Then
                    Select Case WallReflectance
                        Case 50
                            NomCF = {0, 0.948, 0.954, 0.958, 0.961, 0.964, 0.966, 0.969, 0.97, 0.971, 0.973}
                        Case 30
                            NomCF = {0, 0.954, 0.963, 0.969, 0.974, 0.977, 0.979, 0.981, 0.983, 0.985, 0.987}
                        Case 10
                            NomCF = {0, 0.96, 0.971, 0.979, 0.984, 0.988, 0.991, 0.993, 0.995, 0.996, 0.997}
                    End Select
                ElseIf CeilingReflectance = 10 Then
                    Select Case WallReflectance
                        Case 50
                            NomCF = {0, 0.979, 0.978, 0.976, 0.975, 0.975, 0.975, 0.975, 0.976, 0.976, 0.977}
                        Case 30
                            NomCF = {0, 0.983, 0.983, 0.984, 0.985, 0.985, 0.986, 0.987, 0.988, 0.988, 0.989}
                        Case 10
                            NomCF = {0, 0.987, 0.991, 0.993, 0.994, 0.995, 0.996, 0.997, 0.998, 0.998, 0.999}
                    End Select
                End If
        End Select
        Return NomCF(RCR)
    End Function

    Public Function TrueCF(RCR As Decimal, CeilingReflectance As Decimal, WallReflectance As Integer, FloorReflectance As Decimal) As Decimal
        Dim NomCF(15), CorValue(7), SemiValue(3), FinalValue(1), TrueCorrection As Decimal
        Dim a, b, c, d, m, n, i, j, k, l As Integer
        Dim r As Integer = Math.Truncate(RCR)
        Dim w As Integer = WallReflectance
        Dim e As Decimal = CeilingReflectance
        Dim fr As Decimal = FloorReflectance

        'Assigning values to a and b with respect to RCR
        If RCR > 10 Then
            a = 9
            b = 10
        ElseIf RCR < 1 Then
            a = 1
            b = 2
        Else
            If RCR = r Then
                a = r
                b = r
            Else
                a = r
                b = r + 1
            End If
        End If

        'Assigning values to c and d with respect to effective floor reflectance
        If fr > 30 Then
            c = 20
            d = 30
        ElseIf fr > 0 And fr < 10 Then
            c = 0
            d = 10
        ElseIf fr > 10 And fr < 20 Then
            c = 10
            d = 20
        ElseIf fr > 20 And fr < 30 Then
            c = 20
            d = 30
        ElseIf fr = 0 Or fr = 10 Or fr = 20 Or fr = 30 Then
            c = fr
            d = fr
        End If

        'Assigning values to m and n with respect to effective ceiling reflectance;
        'and to i, j, k, and l with respect to wall reflectance in its corresponding ceiling reflectance
        If (e < 80 And e > 70) Or e > 80 Then
            m = 80
            n = 70

            If (w < 70 And w > 50) Or w > 70 Then
                i = 70
                j = 50
            ElseIf w < 50 And w > 30 Then
                i = 50
                j = 30
            ElseIf (w < 30 And w > 10) Or w < 10 Then
                i = 30
                j = 10
            Else
                i = w
                j = w
            End If

            k = i
            l = j
        ElseIf e < 70 And e > 50 Then
            m = 70
            n = 50

            If (w < 70 And w > 50) Or w > 70 Then
                i = 70
                j = 50
            ElseIf w < 50 And w > 30 Then
                i = 50
                j = 30
            ElseIf (w < 30 And w > 10) Or w < 10 Then
                i = 30
                j = 10
            Else
                i = w
                j = w
            End If

            If (w < 50 And w > 30) Or w > 50 Then
                k = 50
                l = 30
            ElseIf (w < 30 And w > 10) Or w < 10 Then
                k = 30
                l = 10
            Else
                k = w
                l = w
            End If
        ElseIf e < 50 And e > 30 Then
            m = 50
            n = 30

            If (w < 50 And w > 30) Or w > 50 Then
                i = 50
                j = 30
                k = 50
                l = 30
            ElseIf (w < 30 And w > 10) Or w < 10 Then
                i = 30
                j = 10
                k = 30
                l = 10
            Else
                i = w
                j = w
                k = w
                l = w
            End If
        ElseIf (e < 30 And e > 10) Or e < 10 Then
            m = 30
            n = 10

            If (w < 50 And w > 30) Or w > 50 Then
                i = 50
                j = 30
                k = 50
                l = 30
            ElseIf (w < 30 And w > 10) Or w < 10 Then
                i = 30
                j = 10
                k = 30
                l = 10
            Else
                i = w
                j = w
                k = w
                l = w
            End If
        Else
            m = e
            n = e
            If e = 80 Or e = 70 Then
                If (w < 70 And w > 50) Or w > 70 Then
                    i = 70
                    j = 50
                ElseIf w < 50 And w > 30 Then
                    i = 50
                    j = 30
                ElseIf (w < 30 And w > 10) Or w < 10 Then
                    i = 30
                    j = 10
                Else
                    i = w
                    j = w
                End If

                k = i
                l = j
            Else
                If (w < 50 And w > 30) Or w > 50 Then
                    i = 50
                    j = 30
                    k = 50
                    l = 30
                ElseIf (w < 30 And w > 10) Or w < 10 Then
                    i = 30
                    j = 10
                    k = 30
                    l = 10
                Else
                    i = w
                    j = w
                    k = w
                    l = w
                End If
            End If
        End If

        NomCF = {NominalCF(a, m, i, c), NominalCF(a, m, j, c), NominalCF(a, n, k, c), NominalCF(a, n, l, c), NominalCF(b, m, i, c), NominalCF(b, m, j, c), NominalCF(b, n, k, c), NominalCF(b, n, l, c),
            NominalCF(a, m, i, d), NominalCF(a, m, j, d), NominalCF(a, n, k, d), NominalCF(a, n, l, d), NominalCF(b, m, i, d), NominalCF(b, m, j, d), NominalCF(b, n, k, d), NominalCF(b, n, l, d)}

        For counter As Integer = 0 To 15 Step 4
            If NomCF(counter) = NomCF(counter + 1) Then
                CorValue(counter / 2) = NomCF(counter)
            Else
                CorValue(counter / 2) = Interpolate(i, j, NomCF(counter), NomCF(counter + 1), w)
            End If
        Next counter

        For counter As Integer = 2 To 15 Step 4
            If NomCF(counter) = NomCF(counter + 1) Then
                CorValue(counter / 2) = NomCF(counter)
            Else
                CorValue(counter / 2) = Interpolate(k, l, NomCF(counter), NomCF(counter + 1), w)
            End If
        Next counter

        For counter As Integer = 0 To 7 Step 2
            If CorValue(counter) = CorValue(counter + 1) Then
                SemiValue(counter / 2) = CorValue(counter)
            Else
                SemiValue(counter / 2) = Interpolate(m, n, CorValue(counter), CorValue(counter + 1), e)
            End If
        Next counter

        For counter As Integer = 0 To 3 Step 2
            If SemiValue(counter) = SemiValue(counter + 1) Then
                FinalValue(counter / 2) = SemiValue(counter)
            Else
                FinalValue(counter / 2) = Interpolate(b, a, SemiValue(counter + 1), SemiValue(counter), RCR)
            End If
        Next counter

        If FinalValue(0) = FinalValue(1) Then
            TrueCorrection = FinalValue(0)
        Else
            TrueCorrection = Interpolate(d, c, FinalValue(1), FinalValue(0), FloorReflectance)
        End If

        Return TrueCorrection
    End Function
#End Region
End Module
