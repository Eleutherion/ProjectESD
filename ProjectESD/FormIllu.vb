Public Class FormIllu
    Private Sub ChkPerimeter_CheckStateChanged(sender As Object, e As EventArgs) Handles ChkPerimeter.CheckStateChanged
        'Toggle area/perimeter activity
        If ChkPerimeter.Checked = True Then
            TxtLength.ReadOnly = True
            TxtWidth.ReadOnly = True
            TxtPerimeter.ReadOnly = False
            TxtArea.ReadOnly = False

            LblArea.Text = "Area*"
            LblPerimeter.Text = "Perimeter*"
            LblLength.Text = "Length"
            LblWidth.Text = "Width"

            TxtLength.Text = ""
            TxtWidth.Text = ""

        ElseIf ChkPerimeter.Checked = False Then
            TxtLength.ReadOnly = False
            TxtWidth.ReadOnly = False
            TxtPerimeter.ReadOnly = True
            TxtArea.ReadOnly = True

            LblArea.Text = "Area"
            LblPerimeter.Text = "Perimeter"
            LblLength.Text = "Length*"
            LblWidth.Text = "Width*"

            TxtPerimeter.Text = ""
            TxtArea.Text = ""
        End If
    End Sub

    Private Sub BtnCompute_Click(sender As Object, e As EventArgs) Handles BtnCompute.Click
        Dim area, perimeter, length, width, height, replacement, hrc, hcc, hfc As Decimal

        hrc = TxthRC.Text
        hcc = TxthCC.Text
        hfc = TxthFC.Text

        height = TxtHeight.Text

        If hrc + hcc + hfc <> height Then
            MessageBox.Show("Sum of cavities must be equal to room height.", "Error")
            GoTo ErrorLine
        Else
            'Compute for area and perimeter if needed
            If ChkPerimeter.Checked = False Then
                length = TxtLength.Text
                width = TxtWidth.Text

                area = Math.Round(length * width, 4)
                perimeter = Math.Round((2 * length) + (2 * width), 4)

                TxtArea.Text = area
                TxtPerimeter.Text = perimeter
            Else
                area = TxtArea.Text
                perimeter = TxtPerimeter.Text
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

            'Acquire RSDD value
            TxtRSDD.Text = Math.Round(TrueRSDD(TxtRCR.Text, TxtLDD.Text, CboLDT.Text) * 100, 4)
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
End Class
