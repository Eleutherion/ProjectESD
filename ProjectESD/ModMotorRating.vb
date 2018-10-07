Module ModMotorRating
    Public Function SinglePhaseMotor(HP As Decimal, Voltage As Integer) As Decimal
        Dim Current(3) As Decimal
        Dim v As Integer = Voltage

        Select Case HP
            Case 1 / 6
                Current = {4.4, 2.5, 2.4, 2.2}
            Case 1 / 4
                Current = {5.8, 3.3, 3.2, 2.9}
            Case 1 / 3
                Current = {7.2, 4.1, 4.0, 3.6}
            Case 1 / 2
                Current = {9.8, 5.6, 5.4, 4.9}
            Case 3 / 4
                Current = {13.8, 7.9, 7.6, 6.9}
            Case 1
                Current = {16, 9.2, 8.8, 8.0}
            Case 1 + (1 / 2)
                Current = {20, 11.5, 11, 10}
            Case 2
                Current = {24, 13.8, 13.2, 12}
            Case 3
                Current = {34, 19.6, 18.7, 17}
            Case 5
                Current = {56, 32.2, 30.8, 28}
            Case 7 + (1 / 2)
                Current = {80, 46, 44, 40}
            Case 10
                Current = {100, 57.5, 55, 50}
        End Select

        Return Current(v)
    End Function

    Public Function ThreePhaseInduction(HP As Decimal, Voltage As Integer) As Decimal
        Dim Current(7) As Decimal
        Dim v As Integer = Voltage

        Select Case HP
            Case 1 / 2
                Current = {4.4, 2.5, 2.4, 2.2, 1.3, 1.1, 0.9, 0}
            Case 3 / 4
                Current = {6.4, 3.7, 3.5, 3.2, 1.8, 1.6, 1.3, 0}
            Case 1
                Current = {8.4, 4.8, 4.6, 4.2, 2.3, 2.1, 1.7, 0}
            Case 1.5
                Current = {12, 6.9, 6.6, 6.0, 3.3, 3.0, 2.4, 0}
            Case 2
                Current = {13.6, 7.8, 78.5, 6.8, 4.3, 3.4, 2.7, 0}
            Case 3
                Current = {19.2, 11.0, 10.6, 9.6, 6.1, 4.8, 3.9, 0}
            Case 5
                Current = {30.4, 17.5, 16.7, 15.2, 9.7, 7.6, 6.1, 0}
            Case 7.5
                Current = {44.0, 25.3, 24.2, 22, 14, 11, 9, 0}
            Case 10
                Current = {56.0, 32.2, 30.8, 28, 18, 14, 11, 0}
            Case 15
                Current = {84.0, 48.3, 46.2, 42, 27, 21, 17, 0}
            Case 20
                Current = {108.0, 62.1, 59.4, 54, 34, 27, 22, 0}
            Case 25
                Current = {136.0, 78.2, 74.8, 68, 44, 34, 27, 0}
            Case 30
                Current = {160.0, 92, 88, 80, 51, 40, 32, 0}
            Case 40
                Current = {208.0, 120, 114, 104, 66, 52, 41, 0}
            Case 50
                Current = {260.0, 150, 143, 130, 83, 65, 52, 0}
            Case 60
                Current = {0, 177, 169, 154, 103, 77, 62, 16}
            Case 75
                Current = {0, 221, 211, 192, 128, 96, 77, 20}
            Case 100
                Current = {0, 285, 273, 248, 165, 124, 99, 26}
            Case 125
                Current = {0, 359, 343, 312, 208, 156, 125, 31}
            Case 150
                Current = {0, 414, 396, 360, 240, 180, 144, 37}
            Case 200
                Current = {0, 552, 528, 480, 320, 240, 192, 49}
            Case 250
                Current = {0, 0, 0, 604, 403, 302, 242, 60}
            Case 300
                Current = {0, 0, 0, 722, 482, 361, 289, 72}
            Case 350
                Current = {0, 0, 0, 828, 560, 414, 336, 83}
            Case 400
                Current = {0, 0, 0, 954, 636, 477, 382, 95}
            Case 450
                Current = {0, 0, 0, 1030, 0, 515, 412, 103}
            Case 500
                Current = {0, 0, 0, 1180, 786, 590, 472, 118}
            Case Else
                Current(v) = -1
        End Select

        Return Current(v)
    End Function

    Public Function ThreePhaseSynchronous(HP As Decimal, Voltage As Integer) As Decimal
        Dim Current(3) As Decimal
        Dim v As Integer

        If Voltage = 3 Then
            v = 4
        Else
            v = Voltage
        End If

        Select Case HP
            Case 25
                Current = {53, 26, 21, 0}
            Case 30
                Current = {63, 32, 26, 0}
            Case 40
                Current = {83, 41, 33, 0}
            Case 50
                Current = {104, 52, 42, 0}
            Case 60
                Current = {123, 61, 49, 12}
            Case 75
                Current = {155, 78, 62, 15}
            Case 100
                Current = {202, 101, 81, 20}
            Case 125
                Current = {253, 126, 101, 25}
            Case 150
                Current = {302, 151, 121, 30}
            Case 200
                Current = {400, 201, 161, 40}
            Case Else
                Current(v - 4) = -1
        End Select

        Return Current(v - 4)
    End Function
End Module
