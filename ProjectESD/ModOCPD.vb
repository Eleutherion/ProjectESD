﻿Module ModOCPD
    Public Function OCPDRating(ByVal Current As Decimal) As Integer
        Dim OCPD(37), v As Integer

        OCPD = {0, 15, 20, 25, 30, 35, 40, 45, 50, 60,
            70, 80, 90, 100, 110, 125, 150, 175, 200, 225,
            250, 300, 350, 400, 450, 500, 600, 700, 800, 1000,
            1200, 1600, 2000, 2500, 3000, 4000, 5000, 6000}

        For i As Integer = 1 To 37 Step 1
            If Current > OCPD(i - 1) And Current <= OCPD(i) Then
                v = i
                GoTo ReturnLine
            Else
                GoTo NextLine
            End If
NextLine: Next i

ReturnLine: Return OCPD(v)
    End Function

    Public Function InterruptingCap(ByVal ShortCircuit As Decimal) As Integer
        Dim kAIC(11), v As Integer

        kAIC = {0, 10, 18, 22, 42, 65, 100, 200, 300, 400, 500, 600}

        For i As Integer = 1 To 11 Step 1
            If ShortCircuit > kAIC(i - 1) And ShortCircuit <= kAIC(i) Then
                v = i
                GoTo ReturnLine
            End If
        Next i
ReturnLine: Return kAIC(v)
    End Function
End Module
