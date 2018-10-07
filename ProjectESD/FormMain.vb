Imports System.Text.RegularExpressions
Imports System.Data.SqlClient

Public Class FormMain

    Private Sub FormMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblWire' table. You can move, or remove it, as needed.
        Me.TblWireTableAdapter.Fill(Me.ESD_DatabaseDataSet.tblWire)
        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblConductor' table. You can move, or remove it, as needed.
        Me.TblConductorTableAdapter.Fill(Me.ESD_DatabaseDataSet.tblConductor)
        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblTransGen' table. You can move, or remove it, as needed.
        Me.TblTransGenTableAdapter.Fill(Me.ESD_DatabaseDataSet.tblTransGen)
        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblSubfeeder' table. You can move, or remove it, as needed.
        Me.TblSubfeederTableAdapter.Fill(Me.ESD_DatabaseDataSet.tblSubfeeder)
        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblBranch' table. You can move, or remove it, as needed.
        Me.TblBranchTableAdapter.Fill(Me.ESD_DatabaseDataSet.tblBranch)
        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblMainFeeder' table. You can move, or remove it, as needed.
        Me.TblMainFeederTableAdapter.Fill(Me.ESD_DatabaseDataSet.tblMainFeeder)
        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblProject' table. You can move, or remove it, as needed.
        Me.TblProjectTableAdapter.Fill(Me.ESD_DatabaseDataSet.tblProject)

        SetTextBox.Text = "1"
    End Sub

#Region "Project"
    Private Sub BtnAddProject_Click(sender As Object, e As EventArgs) Handles BtnAddProject.Click
        TblProjectBindingSource.AddNew()
    End Sub

    Private Sub BtnSaveProject_Click(sender As Object, e As EventArgs) Handles BtnSaveProject.Click
        Dim dr As DialogResult = MessageBox.Show("Do you wish to save?", "Save", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            Try
                Validate()
                TblProjectBindingSource.EndEdit()
                TblProjectTableAdapter.Update(ESD_DatabaseDataSet)

                MessageBox.Show("Record saved.")

                TblProjectTableAdapter.Fill(ESD_DatabaseDataSet.tblProject)
            Catch ex As NoNullAllowedException
                MessageBox.Show("Fill in the required fields.", "No Null Allowed Exception", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Catch ex As ConstraintException
                MessageBox.Show("Duplicate records.", "Constraint Exception", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        Else
            Refresh()
        End If
    End Sub


#End Region

#Region "Main Feeder"

#End Region

#Region "Subfeeder"

#End Region

#Region "Branch Circuit"
    Private Sub BtnComputeBranch_Click(sender As Object, e As EventArgs) Handles BtnComputeBranch.Click
        Dim current, power, ocpd, setcurrent As Decimal
        Dim truerating As Decimal = ConvertFraction(MotorRatingTextBox.Text)
        Dim wirenumber As Integer = 1
        Dim conduit As Integer

        'Do the data checking part
        If TypeComboBox.Text = "" Or VoltageComboBox.Text = "" Or PhaseComboBox.Text = "" Or WireTypeComboBox.Text = "" Then
            MessageBox.Show("Fill in the required fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            GoTo ErrorLine

        ElseIf TypeComboBox.SelectedIndex = 0 And (CboLighting1.SelectedIndex = -1 Or TxtRatingLighting1.Text = "" Or TxtItemsLighting1.Text = "" Or TxtItemsLighting1.Text = "0") Then
            MessageBox.Show("Fill in the required fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            GoTo ErrorLine

        ElseIf TypeComboBox.SelectedIndex = 1 And TxtItemsPower1.Text = "0" And TxtItemsPower2.Text = "0" And TxtItemsPower3.Text = "0" And TxtItemsPower4.Text = "0" Then
            MessageBox.Show("Fill in the required fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            GoTo ErrorLine

        ElseIf TypeComboBox.SelectedIndex = 2 Then
            If truerating = -1 Then
                MessageBox.Show("Please input proper value", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                GoTo ErrorLine

            ElseIf truerating = -2 Then
                MessageBox.Show("Fill in the required fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                GoTo ErrorLine

            ElseIf truerating = -3 Then
                MessageBox.Show("Motor rating input not recognized.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                GoTo ErrorLine

            ElseIf PhaseComboBox.Text = "3" Then
                If MotorTypeComboBox.SelectedIndex = 2 And (VoltageComboBox.SelectedIndex < 3 Or VoltageComboBox.SelectedIndex = 4) Then
                    MessageBox.Show("Invalid voltage rating for given motor type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    GoTo ErrorLine

                ElseIf MotorTypeComboBox.SelectedIndex = 2 And CboRatingUnit.SelectedIndex = 0 And ((CDec(MotorRatingTextBox.Text) < 25) Or
                        (CDec(MotorRatingTextBox.Text) > 200) Or
                        (CDec(MotorRatingTextBox.Text) < 60 And VoltageComboBox.SelectedIndex = 7)) Then
                    MessageBox.Show("Invalid motor rating for given voltage.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    GoTo ErrorLine

                ElseIf (VoltageComboBox.SelectedIndex = 7 And truerating < 60) Or
                        (VoltageComboBox.SelectedIndex = 0 And truerating > 50) Or
                        (VoltageComboBox.SelectedIndex <= 2 And truerating > 200) Then
                    MessageBox.Show("Invalid motor rating for given voltage.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    GoTo ErrorLine

                End If

            ElseIf PhaseComboBox.Text IsNot "3" And MotorTypeComboBox.SelectedIndex = 2 Then
                MessageBox.Show("Given motor phase only valid for three-phase system.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                GoTo ErrorLine
            End If
        End If

        'Computation for Lighting
        If TypeComboBox.SelectedIndex = 0 Then
            TxtTotalLighting1.Text = CInt(TxtRatingLighting1.Text) * CInt(TxtItemsLighting1.Text)
            TxtTotalLighting2.Text = CInt(TxtRatingLighting2.Text) * CInt(TxtItemsLighting2.Text)
            TxtTotalLighting3.Text = CInt(TxtRatingLighting3.Text) * CInt(TxtItemsLighting3.Text)

            power = CInt(TxtTotalLighting1.Text) + CInt(TxtTotalLighting2.Text) + CInt(TxtTotalLighting3.Text)

            If power > 1200 Then
                MessageBox.Show("Maximum of 1200VA for lighting circuits.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                GoTo ErrorLine
            End If

            PowerRatingTextBox.Text = power
            current = CInt(PowerRatingTextBox.Text) / CInt(VoltageComboBox.Text)
            FullLoadCurrentTextBox.Text = Math.Round(current, 4)

            MinimumAmpacityTextBox.Text = FullLoadCurrentTextBox.Text

            ocpd = current * 1.25

            ConduitSizeTextBox1.Text = "15"

            If TypeComboBox1.SelectedIndex = 0 Or TypeComboBox1.SelectedIndex = 1 Then
                If ConductorComboBox.SelectedIndex = 2 Then
                    WireSizeTextBox1.Text = "2.0"
                Else
                    WireSizeTextBox1.Text = "3.5"
                End If
                OCPDRatingTextBox1.Text = "15"
            Else
                If ConductorComboBox.SelectedIndex = 2 Then
                    WireSizeTextBox1.Text = "3.5"
                Else
                    WireSizeTextBox1.Text = "5.5"
                End If
                OCPDRatingTextBox1.Text = "20"
            End If

            'Computation for Power
        ElseIf TypeComboBox.SelectedIndex = 1 Then
            If CInt(TxtItemsPower1.Text) + CInt(TxtItemsPower2.Text) + CInt(TxtItemsPower3.Text) > 10 Then
                MessageBox.Show("Exceeded allowable wall-mounted outlet. Maximum allowable wall outlet is 10. Floor-mounted outlets exempted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                GoTo ErrorLine
            End If

            TxtTotalPower1.Text = CInt(TxtItemsPower1.Text) * 180
            TxtTotalPower2.Text = CInt(TxtItemsPower2.Text) * 360
            TxtTotalPower3.Text = CInt(TxtItemsPower3.Text) * 540
            TxtTotalPower4.Text = CInt(TxtItemsPower4.Text) * 360

            PowerRatingTextBox.Text = CInt(TxtTotalPower1.Text) + CInt(TxtTotalPower2.Text) + CInt(TxtTotalPower3.Text) + CInt(TxtTotalPower4.Text)

            current = CInt(PowerRatingTextBox.Text) / CInt(VoltageComboBox.Text)
            FullLoadCurrentTextBox.Text = Math.Round(current, 4)

            MinimumAmpacityTextBox.Text = CStr(Math.Round(current * 1.25, 4))

            ocpd = current * 1.5

            If WireSet(ConductorComboBox.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), ocpd, ConduitTypeComboBox.Text, WireTypeComboBox.Text) = 1 Then
                wirenumber = SetTextBox.Text
            ElseIf WireSet(ConductorComboBox.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), ocpd, ConduitTypeComboBox.Text, WireTypeComboBox.Text) <= SetTextBox.Text Then
                wirenumber = SetTextBox.Text
            Else
                wirenumber = WireSet(ConductorComboBox.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), ocpd, ConduitTypeComboBox.Text, WireTypeComboBox.Text)
                SetTextBox.Text = wirenumber
            End If

            setcurrent = ocpd / wirenumber

            Dim Size As String = WireSize(ConductorComboBox.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), setcurrent, ConduitTypeComboBox.Text, WireTypeComboBox.Text)
            If Size = "2.0" And ConductorComboBox.Text = "Copper" Then
                WireSizeTextBox1.Text = "3.5"
            ElseIf Size = "3.5" And ConductorComboBox.Text IsNot "Copper" Then
                WireSizeTextBox1.Text = "5.5"
            Else
                WireSizeTextBox1.Text = Size
            End If

            If ocpd < 20 Then
                OCPDRatingTextBox1.Text = "20"
            ElseIf ocpd > 6000 Then
                MessageBox.Show("Required OCPD rating exceeds commercially-available rating. Suggested value indicated. Consider customizing")
                OCPDRatingTextBox1.Text = Math.Truncate((ocpd / 1000) + 1) * 1000
            Else
                OCPDRatingTextBox1.Text = OCPDRating(ocpd)
            End If

            conduit = ConduitSize(WireTypeComboBox.Text, WireSizeTextBox1.Text, GroundWireCheckBox1.Checked, PhaseComboBox.Text, ConduitTypeComboBox.Text)

            If conduit = -1 Then
                MessageBox.Show("No conduit size available for given wire and conduit. Consider adding wire set.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                GoTo ErrorLine
            Else
                ConduitSizeTextBox1.Text = conduit
            End If

            'Computation for Special Equipment
        ElseIf TypeComboBox.SelectedIndex = 2 Then
            If CboRatingUnit.SelectedIndex = 0 Then
                If PhaseComboBox.Text = "3" Then
                    If MotorTypeComboBox.SelectedIndex = 2 Then
                        current = ThreePhaseSynchronous(truerating, VoltageComboBox.SelectedIndex)
                        If current = -1 Then
                            MessageBox.Show("Given motor rating not among standard.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            GoTo ErrorLine
                        End If
                    Else
                        current = ThreePhaseInduction(truerating, VoltageComboBox.SelectedIndex)
                        If current = -1 Then
                            MessageBox.Show("Given motor rating not among standard.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            GoTo ErrorLine
                        End If
                    End If
                    FullLoadCurrentTextBox.Text = Math.Round(current, 4)

                    power = current * CInt(VoltageComboBox.Text) * Math.Sqrt(3)
                    PowerRatingTextBox.Text = Math.Round(power, 4)
                Else
                    current = SinglePhaseCurrent(truerating, VoltageComboBox.SelectedIndex)
                    power = CInt(VoltageComboBox.Text) * current
                    FullLoadCurrentTextBox.Text = CStr(Math.Round(current, 4))
                    PowerRatingTextBox.Text = CStr(Math.Round(power, 4))
                End If
            Else
                power = CDec(MotorRatingTextBox.Text) * 1000 * CInt(TxtMotorItem.Text)
                current = power / CInt(VoltageComboBox.Text)

                PowerRatingTextBox.Text = CStr(Math.Round(power, 4))
                FullLoadCurrentTextBox.Text = CStr(Math.Round(current, 4))
            End If

            If MotorTypeComboBox.SelectedIndex = 3 Or MotorTypeComboBox.SelectedIndex = 4 Then
                MinimumAmpacityTextBox.Text = Math.Round(current * 1.15, 4)
                ocpd = current * 1.75
            Else
                MinimumAmpacityTextBox.Text = Math.Round(current * 1.5, 4)
                ocpd = current * 2.5
            End If

            If WireSet(ConductorComboBox.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), ocpd, ConduitTypeComboBox.Text, WireTypeComboBox.Text) = 1 Then
                wirenumber = SetTextBox.Text
            ElseIf WireSet(ConductorComboBox.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), ocpd, ConduitTypeComboBox.Text, WireTypeComboBox.Text) <= SetTextBox.Text Then
                wirenumber = SetTextBox.Text
            Else
                wirenumber = WireSet(ConductorComboBox.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), ocpd, ConduitTypeComboBox.Text, WireTypeComboBox.Text)
                SetTextBox.Text = wirenumber
            End If

            setcurrent = ocpd / wirenumber

            Dim Size As String = WireSize(ConductorComboBox.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), setcurrent, ConduitTypeComboBox.Text, WireTypeComboBox.Text)
            If Size = "2.0" And ConductorComboBox.Text = "Copper" Then
                WireSizeTextBox1.Text = "3.5"
            ElseIf Size = "3.5" And ConductorComboBox.Text IsNot "Copper" Then
                WireSizeTextBox1.Text = "5.5"
            Else
                WireSizeTextBox1.Text = Size
            End If

            If ocpd < 20 Then
                OCPDRatingTextBox1.Text = "20"
            ElseIf ocpd > 6000 Then
                MessageBox.Show("Required OCPD rating exceeds commercially-available rating. Suggested value indicated. Consider customizing")
                OCPDRatingTextBox1.Text = Math.Truncate((ocpd / 1000) + 1) * 1000
            Else
                OCPDRatingTextBox1.Text = OCPDRating(ocpd)
            End If

            conduit = ConduitSize(WireTypeComboBox.Text, WireSizeTextBox1.Text, GroundWireCheckBox1.Checked, PhaseComboBox.Text, ConduitTypeComboBox.Text)

            If conduit = -1 Then
                MessageBox.Show("No conduit size available for given wire and conduit. Consider adding wire set.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                GoTo ErrorLine
            Else
                ConduitSizeTextBox1.Text = conduit
            End If

        End If

ErrorLine: End Sub

    Private Sub CboLighting1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboLighting1.SelectedIndexChanged
        Select Case CboLighting1.SelectedIndex
            Case 0
                TxtRatingLighting1.Text = 100
            Case 1
                TxtRatingLighting1.Text = 40
            Case 2
                TxtRatingLighting1.Text = 40
            Case 3
                TxtRatingLighting1.Text = 25
            Case 4
                TxtRatingLighting1.Text = 1200
            Case Else
                TxtRatingLighting1.Text = 0
        End Select
    End Sub

    Private Sub CboLighting2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboLighting2.SelectedIndexChanged
        Select Case CboLighting2.SelectedIndex
            Case 0
                TxtRatingLighting2.Text = 100
            Case 1
                TxtRatingLighting2.Text = 40
            Case 2
                TxtRatingLighting2.Text = 40
            Case 3
                TxtRatingLighting2.Text = 25
            Case 4
                TxtRatingLighting2.Text = 1200
            Case Else
                TxtRatingLighting2.Text = 0
        End Select
    End Sub

    Private Sub CboLighting3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboLighting3.SelectedIndexChanged
        Select Case CboLighting3.SelectedIndex
            Case 0
                TxtRatingLighting3.Text = 100
            Case 1
                TxtRatingLighting3.Text = 40
            Case 2
                TxtRatingLighting3.Text = 40
            Case 3
                TxtRatingLighting3.Text = 25
            Case 4
                TxtRatingLighting3.Text = 1200
            Case Else
                TxtRatingLighting3.Text = 0
        End Select
    End Sub

    Private Sub TypeComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TypeComboBox.SelectedIndexChanged
        If TypeComboBox.SelectedIndex = 2 Then
            TxtMotorItem.ReadOnly = False
            MotorRatingTextBox.ReadOnly = False
            CboRatingUnit.Enabled = True
            CboRatingUnit.SelectedIndex = 0
            MotorTypeComboBox.Enabled = True
            MotorTypeComboBox.SelectedIndex = 0
            GrpLighting.Enabled = False
            GrpPower.Enabled = False
            TxtMotorItem.Text = "1"
        Else
            TxtMotorItem.ReadOnly = True
            MotorRatingTextBox.ReadOnly = True
            CboRatingUnit.SelectedIndex = -1
            CboRatingUnit.Enabled = False
            MotorTypeComboBox.SelectedIndex = -1
            MotorTypeComboBox.Enabled = False
            TxtMotorItem.Text = ""
            If TypeComboBox.SelectedIndex = 0 Then
                GrpPower.Enabled = False
                GrpLighting.Enabled = True
            ElseIf TypeComboBox.SelectedIndex = 1 Then
                GrpPower.Enabled = True
                GrpLighting.Enabled = False
            End If
        End If
    End Sub


#End Region

#Region "Transformer/Generator"

#End Region

#Region "Misc"
    Private Function ConvertFraction(ByVal InputNfo As String) As Decimal
        Dim regx As New Regex("[a-zA-Z]")
        Dim output As Decimal

        If IsNumeric(InputNfo) = True Then
            output = InputNfo
        ElseIf regx.IsMatch(InputNfo) Then
            output = -1
        ElseIf InputNfo.Contains("/") And InputNfo.Contains(".") = False Then
            output = FractiontoDecimal(InputNfo)
        ElseIf inputNfo = Nothing Or inputNfo = " " Then
            output = -2
        Else
            output = -3
        End If

        Return output
    End Function

    Private Function ThreePhaseCurrent(ByVal HP As Decimal, ByVal Voltage As Integer) As Decimal

        Return 0
    End Function

    Private Function SinglePhaseCurrent(ByVal HP As Decimal, ByVal Voltage As Integer) As Decimal
        Dim Current As Decimal
        Dim a, b As Decimal

        If (HP > 1 / 6 And HP < 1 / 4) Or HP < 1 / 6 Then
            a = 1 / 6
            b = 1 / 4
        ElseIf HP > 1 / 4 And HP < 1 / 3 Then
            a = 1 / 4
            b = 1 / 3
        ElseIf HP > 1 / 3 And HP < 1 / 2 Then
            a = 1 / 3
            b = 1 / 2
        ElseIf HP > 1 / 2 And HP < 3 / 4 Then
            a = 1 / 2
            b = 3 / 4
        ElseIf HP > 3 / 4 And HP < 1 Then
            a = 3 / 4
            b = 1
        ElseIf HP > 1 And HP < 1.5 Then
            a = 1
            b = 1.5
        ElseIf HP > 1.5 And HP < 2 Then
            a = 1.5
            b = 2
        ElseIf HP > 2 And HP < 3 Then
            a = 2
            b = 3
        ElseIf HP > 3 And HP < 5 Then
            a = 3
            b = 5
        ElseIf HP > 5 And HP < 7.5 Then
            a = 5
            b = 7.5
        ElseIf (HP > 7.5 And HP < 10) Or HP > 10 Then
            a = 7.5
            b = 10
        Else
            a = HP
            b = HP
        End If

        If a = b Then
            Current = SinglePhaseMotor(a, Voltage)
        Else
            Current = Interpolate(b, a, SinglePhaseMotor(b, Voltage), SinglePhaseMotor(a, Voltage), HP)
        End If

        Return Current
    End Function

    Private Sub PhaseComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PhaseComboBox.SelectedIndexChanged
        If PhaseComboBox.SelectedIndex = 3 Then
            TypeComboBox.SelectedIndex = 2
            TypeComboBox.Enabled = False
        Else
            TypeComboBox.Enabled = True
        End If
    End Sub

#End Region
End Class