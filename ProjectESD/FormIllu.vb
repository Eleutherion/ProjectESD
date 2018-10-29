Public Class FormIllu
    Private Sub BtnCompute_Click(sender As Object, e As EventArgs) Handles BtnCompute.Click
        Dim area, perimeter, length, width, radius, height, replacement, hrc, hcc, hfc, cu, cf, cuf, lat, lv, bf, lsd, lbo, ldd, rsdd, ld, llf As Decimal
        Dim lamp As Integer = TxtLamp.Text
        Dim lum As Decimal
        Dim rating As Integer = TxtRating.Text
        Dim lux As Decimal = TxtIllu.Text

        If RdoIndoor.Checked Then
            hrc = TxthRC.Text
            hcc = TxthCC.Text
            hfc = TxthFC.Text

            height = TxtHeight.Text

            If hrc + hcc + hfc <> height Then
                MessageBox.Show("Sum of cavities must be equal to room height.", "Error")
                GoTo ErrorLine
            Else
                If RadioP.Checked Then
                    area = TxtArea.Text
                    perimeter = TxtPerimeter.Text
                ElseIf RadioLW.Checked Then
                    length = TxtLength.Text
                    width = TxtWidth.Text

                    area = length * width
                    perimeter = (2 * length) + (2 * width)

                    TxtArea.Text = Math.Round(area, 4)
                    TxtPerimeter.Text = Math.Round(perimeter, 4)
                ElseIf RadioC.Checked Then
                    radius = TxtLength.Text

                    area = Math.PI * radius ^ 2
                    perimeter = 2 * Math.PI * radius

                    TxtArea.Text = Math.Round(area, 4)
                    TxtPerimeter.Text = Math.Round(perimeter, 4)
                End If

                'Acquire value of ballast factor
                If CboLampType.Text = "General Service Incandescent" Or CboLampType.Text = "Halogen Incandescent" Then
                    TxtBF.Text = 100
                Else
                    TxtBF.Text = 95
                End If

                'Acquire value of lamp burnout factor
                replacement = TxtReplace.Text
                TxtLBO.Text = Math.Round(100 - replacement, 4)

                'Acquire value of lamp lumen depreciation
                TxtLLD.Text = LLD(CboLampType.Text)

                'Acquire value of luminaire depreciation
                TxtLDD.Text = LDDRating(CboLumCat.Text, CboClean.Text, TxtFreq.Text)

                'Acquire cavity ratio values
                TxtRCR.Text = Math.Round(2.5 * hrc * (perimeter / area), 4)
                TxtCCR.Text = Math.Round(2.5 * hcc * (perimeter / area), 4)
                TxtFCR.Text = Math.Round(2.5 * hfc * (perimeter / area), 4)

                Dim CRFactor As Decimal = TxtFCR.Text * 10

                'Acquire RSDD value
                rsdd = TrueRSDD(TxtRCR.Text, TxtLDD.Text, CboLDT.Text)
                TxtRSDD.Text = Math.Round(rsdd * 100, 4)

                lat = TxtLAT.Text / 100
                lv = TxtLV.Text / 100
                bf = TxtBF.Text / 100
                lsd = TxtLSD.Text / 100
                lbo = TxtLBO.Text / 100
                ld = TxtLLD.Text / 100
                ldd = TxtLDD.Text / 100

                'Acquire effective reflectance values
                Txtpcc.Text = TrueER(TxtCCR.Text, Txtpc.Text, Txtpw.Text)
                Txtpfc.Text = TrueER(TxtFCR.Text, Txtpf.Text, Txtpw.Text)

                'Acquire coefficient of utilization
                cu = TrueCU(TxtRCR.Text, Txtpw.Text, Txtpcc.Text)
                TxtCU.Text = Math.Round(cu, 4)

                'Acquire correction factor
                cf = TrueCF(TxtRCR.Text, Txtpcc.Text, Txtpw.Text, Txtpfc.Text)
                TxtCF.Text = Math.Round(cf, 4)

                'Acquire final coefficient of utilization
                cuf = cu * cf
                TxtCUF.Text = Math.Round(cuf, 4)

                llf = lat * lv * bf * lsd * lbo * ld * ldd * rsdd
                TxtLLF.Text = Math.Round(llf, 4)

                lum = lux * area / (rating * llf * lamp * cuf)

                TxtLuminaire.Text = Math.Round(lum, 4)

                If Math.Round(lum) Mod 2 = 0 Then
                    TxtLuminaireItem.Text = Math.Round(lum)
                Else
                    TxtLuminaireItem.Text = Math.Round(lum) + 1
                End If
            End If

        ElseIf RdoOutdoor.Checked Then
            If RadioP.Checked Then
                area = TxtArea.Text
                perimeter = TxtPerimeter.Text
            ElseIf RadioLW.Checked Then
                length = TxtLength.Text
                width = TxtWidth.Text

                area = length * width
                perimeter = (2 * length) + (2 * width)

                TxtArea.Text = Math.Round(area, 4)
                TxtPerimeter.Text = Math.Round(perimeter, 4)
            ElseIf RadioC.Checked Then
                radius = TxtLength.Text

                area = Math.PI * radius ^ 2
                perimeter = 2 * Math.PI * radius

                TxtArea.Text = Math.Round(area, 4)
                TxtPerimeter.Text = Math.Round(perimeter, 4)
            End If

            'Acquire value of ballast factor
            If CboLampType.Text = "General Service Incandescent" Or CboLampType.Text = "Halogen Incandescent" Then
                TxtBF.Text = 100
            Else
                TxtBF.Text = 95
            End If

            'Acquire value of lamp burnout factor
            replacement = TxtReplace.Text
            TxtLBO.Text = Math.Round(100 - replacement, 4)

            'Acquire value of lamp lumen depreciation
            TxtLLD.Text = LLD(CboLampType.Text)

            'Acquire value of luminaire depreciation
            TxtLDD.Text = LDDRating(CboLumCat.Text, CboClean.Text, TxtFreq.Text)

            'Acquire cavity ratio values
            TxtRCR.Text = Math.Round(2.5 * hrc * (perimeter / area), 4)
            'Acquire RSDD value
            rsdd = TrueRSDD(TxtRCR.Text, TxtLDD.Text, CboLDT.Text)
            TxtRSDD.Text = Math.Round(rsdd * 100, 4)

            lat = TxtLAT.Text / 100
            lv = TxtLV.Text / 100
            bf = TxtBF.Text / 100
            lsd = TxtLSD.Text / 100
            lbo = TxtLBO.Text / 100
            ld = TxtLLD.Text / 100
            ldd = TxtLDD.Text / 100

            llf = lat * lv * bf * lsd * lbo * ld * ldd * rsdd
            TxtLLF.Text = Math.Round(llf, 4)

            lum = lux * area / (rating * llf * lamp * cuf)

            TxtLuminaire.Text = Math.Round(lum, 4)

            'TxtLuminaire.Text = Math.Round(lux * area / (rating * llf * lamp), 4)

            'lum = Math.Truncate(lux * area / (rating * llf * lamp) + 1)
            'TxtLuminaire.Text = Math.Truncate(lum + 1)

            If Math.Round(lum) Mod 2 = 0 Then
                TxtLuminaireItem.Text = Math.Round(lum)
            Else
                TxtLuminaireItem.Text = Math.Round(lum) + 1
            End If
        End If
ErrorLine: End Sub

    Private Sub CboMount_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboMount.SelectedIndexChanged
        If CboMount.Text = "Suspended" Then
            TxthCC.ReadOnly = False
            TxthCC.Text = ""
        Else
            TxthCC.ReadOnly = True
            TxthCC.Text = 0
        End If
    End Sub

    Private Sub RadioP_CheckedChanged(sender As Object, e As EventArgs) Handles RadioP.CheckedChanged
        TxtArea.ReadOnly = False
        TxtPerimeter.ReadOnly = False
        TxtLength.ReadOnly = True
        TxtWidth.Enabled = True
        TxtWidth.ReadOnly = True

        LblLength.Text = "Length"
        LblWidth.Text = "Width"
        LblArea.Text = "Area*"
        LblPerimeter.Text = "Perimeter*"
    End Sub

    Private Sub RadioLW_CheckedChanged(sender As Object, e As EventArgs) Handles RadioLW.CheckedChanged
        TxtArea.ReadOnly = True
        TxtPerimeter.ReadOnly = True
        TxtLength.ReadOnly = False
        TxtWidth.Enabled = True
        TxtWidth.ReadOnly = False

        LblLength.Text = "Length*"
        LblWidth.Text = "Width*"
        LblArea.Text = "Area"
        LblPerimeter.Text = "Perimeter"
    End Sub

    Private Sub RadioC_CheckedChanged(sender As Object, e As EventArgs) Handles RadioC.CheckedChanged
        TxtArea.ReadOnly = True
        TxtPerimeter.ReadOnly = True
        TxtLength.ReadOnly = False
        TxtWidth.ReadOnly = True
        TxtWidth.Enabled = False

        LblLength.Text = "Radius*"
        LblWidth.Text = ""
        LblArea.Text = "Area"
        LblPerimeter.Text = "Perimeter"
    End Sub

    Private Sub FormIllu_FormClosing(sender As Object, e As FormClosingEventArgs)
        'Dim f As Form = FormMain
        'f.Show()

        Dispose()
    End Sub

    Private Sub RdoOutdoor_CheckedChanged(sender As Object, e As EventArgs) Handles RdoOutdoor.CheckedChanged
        Txtpc.Enabled = False
        Txtpw.Enabled = False
        Txtpf.Enabled = False
        TxthRC.Enabled = False
        TxthCC.Enabled = False
        TxthFC.Enabled = False

        Lblpc.Text = "Ceiling"
        Lblpw.Text = "Wall"
        Lblpf.Text = "Floor"
        LblhRC.Text = "hRC"
        LblhCC.Text = "hCC"
        LblhFC.Text = "hFC"
    End Sub

    Private Sub RdoIndoor_CheckedChanged(sender As Object, e As EventArgs) Handles RdoIndoor.CheckedChanged
        Txtpc.Enabled = True
        Txtpw.Enabled = True
        Txtpf.Enabled = True
        TxthRC.Enabled = True
        TxthCC.Enabled = True
        TxthFC.Enabled = True

        Lblpc.Text = "Ceiling*"
        Lblpw.Text = "Wall*"
        Lblpf.Text = "Floor*"
        LblhRC.Text = "hRC*"
        LblhCC.Text = "hCC*"
        LblhFC.Text = "hFC*"
    End Sub
End Class
