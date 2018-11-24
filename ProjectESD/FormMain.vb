Option Infer On
Imports System.Text.RegularExpressions
Imports System.Numerics

Public Class FormMain
    Private Balanced As Boolean = False

    Private constring As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\ESD_Database.mdf;Integrated Security=True"

    Private TCL As Decimal

    Private Sub FormMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TblWireTableAdapter.Fill(ESD_DatabaseDataSet.tblWire)
        TblConductorTableAdapter.Fill(ESD_DatabaseDataSet.tblConductor)
        TblProjectTableAdapter.Fill(ESD_DatabaseDataSet.tblProject)
    End Sub

#Region "Project"
    Private Sub BtnAddProject_Click(sender As Object, e As EventArgs) Handles BtnAddProject.Click
        TblProjectBindingSource.AddNew()
    End Sub

    Private Sub TypeComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TypeComboBox1.SelectedIndexChanged
        If TypeComboBox1.SelectedIndex < 2 Then
            PhaseComboBox.SelectedIndex = 0
            PhaseComboBox.Enabled = False
        Else
            PhaseComboBox.SelectedIndex = -1
            PhaseComboBox.Enabled = True
        End If
    End Sub

    Private Sub BtnSaveProject_Click(sender As Object, e As EventArgs) Handles BtnSaveProject.Click
        Dim dr As DialogResult = MessageBox.Show("Do you wish to save?", "Save", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            Try
                Validate()
                TblProjectBindingSource.EndEdit()
                TblProjectTableAdapter.Update(ESD_DatabaseDataSet)

                MessageBox.Show("Record saved.")

                TblMainFeederTableAdapter.FillByProject(ESD_DatabaseDataSet.tblMainFeeder, ProjectCodeTextBox.Text)
                TblTransGenTableAdapter.FillByProject(ESD_DatabaseDataSet.tblTransGen, ProjectCodeTextBox.Text)
                TblDistributionTableAdapter.FillByProject(ESD_DatabaseDataSet.tblDistribution, ProjectCodeTextBox.Text)
                TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
                TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub BtnDeleteProject_Click(sender As Object, e As EventArgs) Handles BtnDeleteProject.Click
        Dim dr As DialogResult = MessageBox.Show("Do you wish to delete? Once deleted, the record will not be retrieved.", "Delete", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            Try
                TblProjectTableAdapter.DeleteByPrimary(ProjectCodeTextBox.Text)

                MessageBox.Show("Record deleted.")

                TblProjectTableAdapter.Fill(ESD_DatabaseDataSet.tblProject)
                TblMainFeederTableAdapter.FillByProject(ESD_DatabaseDataSet.tblMainFeeder, ProjectCodeTextBox.Text)
                TblTransGenTableAdapter.FillByProject(ESD_DatabaseDataSet.tblTransGen, ProjectCodeTextBox.Text)
                TblDistributionTableAdapter.FillByProject(ESD_DatabaseDataSet.tblDistribution, ProjectCodeTextBox.Text)
                TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
                TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub BtnNextProject_Click(sender As Object, e As EventArgs) Handles BtnNextProject.Click
        TblProjectBindingSource.MoveNext()
        TblMainFeederTableAdapter.FillByProject(ESD_DatabaseDataSet.tblMainFeeder, ProjectCodeTextBox.Text)
        TblTransGenTableAdapter.FillByProject(ESD_DatabaseDataSet.tblTransGen, ProjectCodeTextBox.Text)
        TblDistributionTableAdapter.FillByProject(ESD_DatabaseDataSet.tblDistribution, ProjectCodeTextBox.Text)
        TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
        TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
    End Sub

    Private Sub BtnPreviousProject_Click(sender As Object, e As EventArgs) Handles BtnPreviousProject.Click
        TblProjectBindingSource.MovePrevious()
        TblMainFeederTableAdapter.FillByProject(ESD_DatabaseDataSet.tblMainFeeder, ProjectCodeTextBox.Text)
        TblTransGenTableAdapter.FillByProject(ESD_DatabaseDataSet.tblTransGen, ProjectCodeTextBox.Text)
        TblDistributionTableAdapter.FillByProject(ESD_DatabaseDataSet.tblDistribution, ProjectCodeTextBox.Text)
        TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
        TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
    End Sub
#End Region

#Region "Main Feeder"
    Private Sub BtnCompute_Click(sender As Object, e As EventArgs) Handles BtnCompute.Click
        Dim imf, hrml, phaseload(2), singlephaseload, threephaseload, setcurrent As Decimal
        Dim voltage As Integer = VoltageLevelComboBox.Text
        Dim wirenumber As Integer

        If PhaseTextBox.Text = "3" Then
            phaseload = {TblBranchTableAdapter.TotalLoadperPhase(ProjectCodeTextBox.Text, "A"),
                TblBranchTableAdapter.TotalLoadperPhase(ProjectCodeTextBox.Text, "B"),
                TblBranchTableAdapter.TotalLoadperPhase(ProjectCodeTextBox.Text, "C")}

            Array.Sort(phaseload)

            singlephaseload = phaseload(2)

            threephaseload = TblBranchTableAdapter.TotalLoadperPhase(ProjectCodeTextBox.Text, "3")

            TCL = (3 * singlephaseload) + threephaseload

            TotalLoadTextBox.Text = Math.Round(TCL, 4)
        Else
            singlephaseload = TblBranchTableAdapter.TotalLoadperPhase(ProjectTextBox.Text, "A")
            TCL = singlephaseload
            TotalLoadTextBox.Text = Math.Round(singlephaseload, 4)
        End If

        If PhaseTextBox.Text = "3" Then
            If TblBranchTableAdapter.GetHRMLSinglePhaseProject(ProjectCodeTextBox.Text) > TblBranchTableAdapter.GetHRMLThreePhaseProject(ProjectCodeTextBox.Text) Then
                hrml = TblBranchTableAdapter.GetHRMLSinglePhaseProject(ProjectCodeTextBox.Text)
                imf = (TCL / (voltage * Math.Sqrt(3))) + (Math.Sqrt(3) * hrml * 0.25 / voltage)
            Else
                hrml = TblBranchTableAdapter.GetHRMLThreePhaseProject(ProjectCodeTextBox.Text)
                imf = (TCL + 0.25 * hrml) / (voltage * Math.Sqrt(3))
            End If

            LineCurrentTextBox.Text = Math.Round(imf, 4)

            OCPDRatingTextBox.Text = OCPDRating(imf * 1.5)

            If WireSet(ConductorComboBoxMain.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxMain.Text), OCPDRatingTextBox.Text, ConduitTypeComboBoxMain.Text, WireTypeComboBoxMain.Text) = 1 Then
                wirenumber = SetTextBox1.Text
            ElseIf WireSet(ConductorComboBoxMain.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxMain.Text), OCPDRatingTextBox.Text, ConduitTypeComboBoxMain.Text, WireTypeComboBoxMain.Text) <= SetTextBox1.Text Then
                wirenumber = SetTextBox1.Text
            Else
                wirenumber = WireSet(ConductorComboBoxMain.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxMain.Text), OCPDRatingTextBox.Text, ConduitTypeComboBoxMain.Text, WireTypeComboBoxMain.Text)
                SetTextBox1.Text = wirenumber
            End If

            setcurrent = OCPDRatingTextBox.Text / wirenumber

            WireSizeTextBox.Text = WireSize(ConductorComboBoxMain.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxMain.Text), setcurrent, ConduitTypeComboBoxMain.Text, WireTypeComboBoxMain.Text)

            ConduitSizeTextBox.Text = ConduitSize(WireTypeComboBoxMain.Text, WireSizeTextBox.Text, GroundWireCheckBox.Checked, "3", ConduitTypeComboBoxMain.Text, False)
        Else
            hrml = TblBranchTableAdapter.GetHRMLSinglePhaseProject(ProjectCodeTextBox.Text)
            imf = (TCL / voltage) + (0.25 * hrml / voltage)

            LineCurrentTextBox.Text = Math.Round(imf, 4)

            If TypeComboBox1.SelectedIndex > 1 And OCPDRating(imf * 1.5) = 15 Then
                OCPDRatingTextBox.Text = 20
            Else
                OCPDRatingTextBox.Text = OCPDRating(imf * 1.5)
            End If

            If WireSet(ConductorComboBoxMain.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxMain.Text), OCPDRatingTextBox.Text, ConduitTypeComboBoxMain.Text, WireTypeComboBoxMain.Text) = 1 Then
                wirenumber = SetTextBox1.Text
            ElseIf WireSet(ConductorComboBoxMain.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxMain.Text), OCPDRatingTextBox.Text, ConduitTypeComboBoxMain.Text, WireTypeComboBoxMain.Text) <= SetTextBox.Text Then
                wirenumber = SetTextBox1.Text
            Else
                wirenumber = WireSet(ConductorComboBoxMain.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxMain.Text), OCPDRatingTextBox.Text, ConduitTypeComboBoxMain.Text, WireTypeComboBoxMain.Text)
                SetTextBox1.Text = wirenumber
            End If

            setcurrent = OCPDRatingTextBox.Text / wirenumber

            WireSizeTextBox.Text = WireSize(ConductorComboBoxMain.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxMain.Text), setcurrent, ConduitTypeComboBoxMain.Text, WireTypeComboBoxMain.Text)

            ConduitSizeTextBox.Text = ConduitSize(WireTypeComboBoxMain.Text, ConductorComboBoxMain.Text, GroundWireCheckBox.Checked, "1", ConduitTypeComboBoxMain.Text, False)
        End If

        If GroundWireCheckBox.Checked Then
            GroundWireSizeTextBoxMain.Text = GroundWire(OCPDRatingTextBox.Text, GroundConductorComboBoxMain.Text)
        End If

        If ShortCircuitCapTextBox.Text = "" Or XRRatioTextBox.Text = "" Then
            GoTo endline
        Else
            Dim count As Integer = TblDistributionTableAdapter.CountDP(ProjectCodeTextBox.Text)
            Dim zreal(count - 1), zimag(count - 1) As Double

            Dim z(count - 1), y(count - 1), ytotal, zeq, zline, zutility, ztrans, z1, y1, yeq As Complex

            Dim dt As DataTable = ESD_DatabaseDataSet.tblDistribution

            TblDistributionTableAdapter.FillImpedance(dt, ProjectCodeTextBox.Text)

            zreal = (From row In dt.AsEnumerable() Select row.Field(Of Double)("ImpedanceReal")).ToArray()
            zimag = (From row In dt.AsEnumerable() Select row.Field(Of Double)("ImpedanceImag")).ToArray()
            zline = LineImpedance(DistancetoSETextBox.Text, ConduitTypeComboBoxMain.Text, WireSizeTextBox.Text, ConductorComboBoxMain.Text, VoltageLevelComboBox.Text, SetTextBox1.Text)

            zutility = New Complex(1 / ShortCircuitCapTextBox.Text * Math.Cos(Math.Atan(XRRatioTextBox.Text)), 1 / ShortCircuitCapTextBox.Text * Math.Sin(Math.Atan(XRRatioTextBox.Text)))

            ztrans = New Complex(TransformerZpuTextBox.Text / 100 * (1 / TransformerRatingTextBox.Text) * Math.Cos(Math.Atan(TransformerXRRatioTextBox.Text)),
                                 TransformerZpuTextBox.Text / 100 * (1 / TransformerRatingTextBox.Text) * Math.Sin(Math.Atan(TransformerXRRatioTextBox.Text)))

            z1 = zline + zutility + ztrans
            y1 = Complex.Reciprocal(z1)

            For i As Integer = 0 To count - 1 Step 1
                z(i) = New Complex(zreal(i), zimag(i))
                y(i) = Complex.Reciprocal(z(i))
            Next i

            For Each item As Complex In y
                ytotal += item
            Next

            yeq = ytotal + y1
            zeq = Complex.Reciprocal(yeq)

            ImpedanceRealTextBoxMain.Text = Math.Round(zeq.Real, 4)
            ImpedanceImagTextBoxMain.Text = Math.Round(zeq.Imaginary, 4)

            Dim ifpu As Complex = Complex.Reciprocal(zeq)
            Dim ibase, iactual As Decimal

            If PhaseTextBox.Text = "3" Then
                ibase = 1000000 / (Math.Sqrt(3) * VoltageLevelComboBox.Text)
            Else
                ibase = 1000000 / VoltageLevelComboBox.Text
            End If

            iactual = ifpu.Magnitude * ibase

            KAICRatingTextBox.Text = InterruptingCap(iactual / 1000)
endline: End If
    End Sub

    Private Sub BtnAddMain_Click(sender As Object, e As EventArgs) Handles BtnAddMain.Click
        TblMainFeederBindingSource.AddNew()
        ProjectCodeTextBox1.Text = ProjectCodeTextBox.Text
        SetTextBox1.Text = "1"
    End Sub

    Private Sub BtnSaveMain_Click(sender As Object, e As EventArgs) Handles BtnSaveMain.Click
        Dim dr As DialogResult = MessageBox.Show("Do you wish to save?", "Save", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            Try
                Validate()
                TblMainFeederBindingSource.EndEdit()
                TblMainFeederTableAdapter.Update(ESD_DatabaseDataSet.tblMainFeeder)

                MessageBox.Show("Record saved.")
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub BtnDeleteMain_Click(sender As Object, e As EventArgs) Handles BtnDeleteMain.Click
        Dim dr As DialogResult = MessageBox.Show("Do you wish to delete? Once deleted, the record will not be retrieved.", "Delete", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            Try
                TblMainFeederTableAdapter.DeletebyPrimary(ProjectCodeTextBox1.Text)

                MessageBox.Show("Record deleted.")
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub GroundWireCheckBox_CheckStateChanged(sender As Object, e As EventArgs) Handles GroundWireCheckBox.CheckStateChanged
        If GroundWireCheckBox.Checked Then
            GroundWireSizeTextBoxMain.Enabled = True
            GroundConductorComboBoxMain.Enabled = True
        Else
            GroundWireSizeTextBoxMain.Enabled = False
            GroundConductorComboBoxMain.Enabled = False

            GroundConductorComboBoxMain.SelectedIndex = -1
        End If
    End Sub
#End Region

#Region "Distribution Panel"
    Private Sub BtnComputeDP_Click(sender As Object, e As EventArgs) Handles BtnComputeDP.Click
        Dim ThreePhase As Boolean = False
        Dim voltage As Integer = VoltageLevelComboBox.Text
        Dim idp, load, hrml, phaseload(2), setcurrent, lighting, power, demandload, equip As Decimal
        Dim wirenumber As Integer

        If VoltageLevelComboBox.Text = "" Or WireTypeComboBoxDP.Text = "" Or ConductorComboBoxDP.Text = "" Or SetTextBoxDP.Text = "" Or ConduitTypeComboBoxDP.Text = "" Or DistancetoMainTextBoxDP.Text = "" Then
            MessageBox.Show("Fill in the required fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            GoTo ErrorLine
        End If

        If IsNothing(TblSubfeederTableAdapter.CurrentSum(CodeTextBoxDP.Text)) Then
            MessageBox.Show("Compute for subfeeders before proceeding to the distribution panelboard computation.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            GoTo ErrorLine
        End If

        TextBox6.Text = TblDistributionTableAdapter.TotalPhaseLoad(CodeTextBoxDP.Text, "A")
        TextBox5.Text = TblDistributionTableAdapter.TotalPhaseLoad(CodeTextBoxDP.Text, "B")
        TextBox4.Text = TblDistributionTableAdapter.TotalPhaseLoad(CodeTextBoxDP.Text, "C")
        TextBox3.Text = TblDistributionTableAdapter.TotalPhaseLoad(CodeTextBoxDP.Text, "3")
        TextBox2.Text = TblDistributionTableAdapter.GetHRMLSinglePhase(CodeTextBoxDP.Text)
        TextBox1.Text = TblDistributionTableAdapter.GetHRMLThreePhase(CodeTextBoxDP.Text)

        If TypeComboBox1.SelectedIndex <= 1 Then
            lighting = TblDistributionTableAdapter.TotalLoadperType(CodeTextBoxDP.Text, "Lighting")
            power = TblDistributionTableAdapter.TotalLoadperType(CodeTextBoxDP.Text, "Power")

            If lighting + power <= 3000 Then
                demandload = lighting + power
            ElseIf lighting + power > 3000 And lighting + power <= 120000 Then
                demandload = 3000 + ((lighting + power - 3000) * 0.35)
            Else
                demandload = 3000 + (117000 * 0.35) + ((lighting + power - 120000) * 0.25)
            End If

            equip = TblDistributionTableAdapter.TotalLoadperType(CodeTextBoxDP.Text, "Motor Equipment") +
                TblDistributionTableAdapter.TotalLoadperType(CodeTextBoxDP.Text, "Non-Motor Equipment") +
                TblDistributionTableAdapter.TotalLoadperType(CodeTextBoxDP.Text, "Spare - Lighting") +
                TblDistributionTableAdapter.TotalLoadperType(CodeTextBoxDP.Text, "Spare - General") +
                (TblDistributionTableAdapter.TotalLoadperType(CodeTextBoxDP.Text, "Electric Range") * 0.8)

            load = equip + demandload

            idp = load / voltage

            CurrentRatingTextBoxDP.Text = Math.Round(idp * 1.25, 4)

            OCPDRatingTextBoxDP.Text = OCPDRating(idp * 1.5)

            If WireSet(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), OCPDRatingTextBoxDP.Text, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text) = 1 Then
                wirenumber = SetTextBoxDP.Text
            ElseIf WireSet(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), OCPDRatingTextBoxDP.Text, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text) <= SetTextBoxDP.Text Then
                wirenumber = SetTextBoxDP.Text
            Else
                wirenumber = WireSet(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), OCPDRatingTextBoxDP.Text, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text)
                SetTextBoxDP.Text = wirenumber
            End If

            setcurrent = OCPDRatingTextBoxDP.Text / wirenumber

            If WireSize(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), OCPDRatingTextBoxDP.Text, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text) = "2.0" Then
                WireSizeTextBoxDP.Text = "3.5"
            Else
                WireSizeTextBoxDP.Text = WireSize(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), OCPDRatingTextBoxDP.Text, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text)
            End If
            ConduitSizeTextBoxDP.Text = ConduitSize(WireTypeComboBoxDP.Text, WireSizeTextBoxDP.Text, False, "1", ConduitTypeComboBoxDP.Text, False)
        Else
            If PhaseTextBox.Text = "3" Then
                ThreePhase = True

                load = TblDistributionTableAdapter.TotalPhaseLoad(CodeTextBoxDP.Text, 3)
                phaseload = {TblDistributionTableAdapter.TotalPhaseLoad(CodeTextBoxDP.Text, "A"), TblDistributionTableAdapter.TotalPhaseLoad(CodeTextBoxDP.Text, "B"), TblDistributionTableAdapter.TotalPhaseLoad(CodeTextBoxDP.Text, "C")}
                Array.Sort(phaseload)

                If TblDistributionTableAdapter.GetHRMLThreePhase(CodeTextBoxDP.Text) >= TblDistributionTableAdapter.GetHRMLSinglePhase(CodeTextBoxDP.Text) Then
                    hrml = TblDistributionTableAdapter.GetHRMLThreePhase(CodeTextBoxDP.Text)
                    idp = ((load + (0.25 * hrml)) / (Math.Sqrt(3) * voltage)) + (Math.Sqrt(3) * phaseload(2) / voltage)
                Else
                    hrml = TblDistributionTableAdapter.GetHRMLSinglePhase(CodeTextBoxDP.Text)
                    idp = (Math.Sqrt(3) * (phaseload(2) + (0.25 * hrml)) / voltage) + (load / (Math.Sqrt(3) * voltage))
                End If

                CurrentRatingTextBoxDP.Text = Math.Round(idp, 4)

                OCPDRatingTextBoxDP.Text = OCPDRating(idp * 1.5)

                If WireSet(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), OCPDRatingTextBoxDP.Text, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text) = 1 Then
                    wirenumber = SetTextBoxDP.Text
                ElseIf WireSet(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), OCPDRatingTextBoxDP.Text, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text) <= SetTextBoxDP.Text Then
                    wirenumber = SetTextBoxDP.Text
                Else
                    wirenumber = WireSet(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), OCPDRatingTextBoxDP.Text, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text)
                    SetTextBoxDP.Text = wirenumber
                End If

                setcurrent = OCPDRatingTextBoxDP.Text / wirenumber

                If WireSize(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), setcurrent, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text) = "2.0" Then
                    WireSizeTextBoxDP.Text = "3.5"
                Else
                    WireSizeTextBoxDP.Text = WireSize(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), setcurrent, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text)
                End If

                ConduitSizeTextBoxDP.Text = ConduitSize(WireTypeComboBoxDP.Text, WireSizeTextBoxDP.Text, False, PhaseTextBox.Text, ConduitTypeComboBoxDP.Text, False)

            Else
                load = TblDistributionTableAdapter.TotalPhaseLoad(CodeTextBoxDP.Text, "A")
                hrml = TblDistributionTableAdapter.GetHRMLSinglePhase(CodeTextBoxDP.Text)

                idp = (load + (0.25 * hrml)) / voltage

                CurrentRatingTextBoxDP.Text = Math.Round(idp, 4)

                OCPDRatingTextBoxDP.Text = OCPDRating(idp * 1.5)
                If TypeComboBox1.SelectedIndex > 1 And OCPDRating(idp * 1.5) = 15 Then
                    OCPDRatingTextBoxDP.Text = 20
                Else
                    OCPDRatingTextBoxDP.Text = OCPDRating(idp * 1.5)
                End If

                If WireSet(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), OCPDRatingTextBoxDP.Text, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text) = 1 Then
                    wirenumber = SetTextBoxDP.Text
                ElseIf WireSet(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), OCPDRatingTextBoxDP.Text, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text) <= SetTextBoxDP.Text Then
                    wirenumber = SetTextBoxDP.Text
                Else
                    wirenumber = WireSet(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), OCPDRatingTextBoxDP.Text, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text)
                    SetTextBoxDP.Text = wirenumber
                End If

                setcurrent = OCPDRatingTextBoxDP.Text / wirenumber

                If WireSize(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), setcurrent, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text) = "2.0" Then
                    WireSizeTextBoxDP.Text = "3.5"
                Else
                    WireSizeTextBoxDP.Text = WireSize(ConductorComboBoxDP.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxDP.Text), setcurrent, ConduitTypeComboBoxDP.Text, WireTypeComboBoxDP.Text)
                End If

                ConduitSizeTextBoxDP.Text = ConduitSize(WireTypeComboBoxDP.Text, WireSizeTextBoxDP.Text, False, "1", ConduitTypeComboBoxDP.Text, False)
            End If
        End If

        Dim count As Integer = TblSubfeederTableAdapter.CountSubfeeder(CodeTextBoxDP.Text)
        Dim zreal(count - 1), zimag(count - 1) As Double

        Dim z(count - 1), y(count - 1), ytotal, zeq, zline, ztotal As Complex

        Dim dt As DataTable = ESD_DatabaseDataSet.tblSubfeeder

        TblSubfeederTableAdapter.FillImpedance(dt, CodeTextBoxDP.Text)

        zreal = (From row In dt.AsEnumerable() Select row.Field(Of Double)("ImpedanceReal")).ToArray()
        zimag = (From row In dt.AsEnumerable() Select row.Field(Of Double)("ImpedanceImag")).ToArray()
        zline = LineImpedance(DistancetoMainTextBoxDP.Text, ConduitTypeComboBoxDP.Text, WireSizeTextBoxDP.Text, ConductorComboBoxDP.Text, VoltageLevelComboBox.Text, SetTextBoxDP.Text)

        For i As Integer = 0 To count - 1 Step 1
            z(i) = New Complex(zreal(i), zimag(i))
            y(i) = Complex.Reciprocal(z(i))
        Next i

        For Each item As Complex In y
            ytotal += item
        Next

        zeq = Complex.Reciprocal(ytotal)

        ztotal = zeq + zline

        ImpedanceRealTextBoxDP.Text = Math.Round(ztotal.Real, 4)
        ImpedanceImagTextBoxDP.Text = Math.Round(ztotal.Imaginary, 4)

ErrorLine: End Sub

    Private Sub BtnAddDP_Click(sender As Object, e As EventArgs) Handles BtnAddDP.Click
        TblDistributionBindingSource.AddNew()
        ProjectCodeTextBoxDP.Text = ProjectCodeTextBox.Text
        SetTextBoxDP.Text = 1

        TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
    End Sub

    Private Sub BtnSaveDP_Click(sender As Object, e As EventArgs) Handles BtnSaveDP.Click
        CodeTextBoxDP.Text = ProjectCodeTextBoxDP.Text + "-DP" + DPNumberTextBox.Text

        Dim dr As DialogResult = MessageBox.Show("Do you wish to save?", "Save", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            Try
                Validate()
                TblDistributionBindingSource.EndEdit()
                TblDistributionTableAdapter.Update(ESD_DatabaseDataSet.tblDistribution)

                MessageBox.Show("Record saved.")
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub BtnDeleteDP_Click(sender As Object, e As EventArgs) Handles BtnDeleteDP.Click
        Dim dr As DialogResult = MessageBox.Show("Do you wish to delete? Once deleted, the record will not be retrieved.", "Delete", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            Try
                TblDistributionTableAdapter.DeletebyPrimary(CodeTextBoxDP.Text)

                MessageBox.Show("Record deleted.")

                TblDistributionTableAdapter.FillByProject(ESD_DatabaseDataSet.tblDistribution, ProjectCodeTextBox.Text)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub BtnPreviousDP_Click(sender As Object, e As EventArgs) Handles BtnPreviousDP.Click
        TblDistributionBindingSource.MovePrevious()
        TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
        TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
    End Sub

    Private Sub BtnNextDP_Click(sender As Object, e As EventArgs) Handles BtnNextDP.Click
        TblDistributionBindingSource.MoveNext()
        TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
        TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
    End Sub
#End Region

#Region "Subfeeder"
    Private Sub BtnComputeSub_Click(sender As Object, e As EventArgs) Handles BtnComputeSub.Click
        Dim ThreePhase As Boolean
        Dim voltage As Integer = VoltageLevelComboBox.Text
        Dim isf, load, hrml, phaseload(2), setcurrent, vd As Decimal
        Dim wirenumber As Integer
        Dim phase As String
        Dim Z As Complex

        If VoltageLevelComboBox.Text = "" Or WireTypeComboBoxSub.Text = "" Or ConductorComboBoxSub.Text = "" Or ConduitTypeComboBoxSub.Text = "" Or SetTextBoxSub.Text = "" Or DistancetoMainTextBox.Text = "" Then
            MessageBox.Show("Fill in the required fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            GoTo ErrorLine
        ElseIf NeutralCheckBox.Checked And (NeutralConductorComboBox.Text = "" Or NeutralWireComboBox.Text = "") Then
            MessageBox.Show("Fill in the required fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            GoTo ErrorLine
        End If

        If Balanced = False Then
            MessageBox.Show("Check balancing of subfeeder before proceeding.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            GoTo ErrorLine
        End If

        If PhaseTextBox.Text = "3" Then
            ThreePhase = True
            phase = 3

            phaseload = {TxtPhaseALoad.Text, TxtPhaseBLoad.Text, TxtPhaseCLoad.Text}
            load = TxtThreePhaseLoad.Text

            Array.Sort(phaseload)

            If TblBranchTableAdapter.GetHRMLSinglePhase(CodeTextBoxSub.Text) > TblBranchTableAdapter.GetHRMLThreePhase(CodeTextBoxSub.Text) Then
                hrml = TblBranchTableAdapter.GetHRMLSinglePhase(CodeTextBoxSub.Text)
                isf = (Math.Sqrt(3) * ((phaseload(2) + (0.25 * hrml)) / voltage)) + (load / (Math.Sqrt(3) * voltage))
            Else
                hrml = TblBranchTableAdapter.GetHRMLThreePhase(CodeTextBoxSub.Text)
                isf = (Math.Sqrt(3) * phaseload(2) / voltage) + (load + (0.25 * hrml)) / (Math.Sqrt(3) * voltage)
            End If
        Else
            ThreePhase = False
            phase = 1

            load = TxtPhaseALoad.Text
            hrml = TxtMotorSinglePhase.Text

            isf = (load + (0.25 * hrml)) / voltage
        End If

        CurrentRatingTextBox.Text = Math.Round(isf, 4)
        MinimumAmpacityTextBoxSub.Text = Math.Round(isf * 1.25, 4)

        OCPDRatingTextBoxSub.Text = OCPDRating(isf * 1.5)

        If WireSet(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), OCPDRatingTextBoxSub.Text, ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text) = 1 Then
            wirenumber = SetTextBoxSub.Text
        ElseIf WireSet(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), OCPDRatingTextBoxSub.Text, ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text) <= SetTextBoxSub.Text Then
            wirenumber = SetTextBoxSub.Text
        Else
            wirenumber = WireSet(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), OCPDRatingTextBoxSub.Text, ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text)
            SetTextBoxSub.Text = wirenumber
        End If

        setcurrent = OCPDRatingTextBoxSub.Text / wirenumber

        If WireSize(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxSub.Text), setcurrent, ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text) = "2.0" Then
            WireSizeTextBoxSub.Text = "3.5"
        Else
            WireSizeTextBoxSub.Text = WireSize(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxSub.Text), setcurrent, ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text)
        End If

        ConduitSizeTextBoxSub.Text = ConduitSize(WireTypeComboBoxSub.Text, WireSizeTextBoxSub.Text, GroundWireCheckBoxSub.Checked, phase, ConduitTypeComboBoxSub.Text, NeutralCheckBox.Checked)

        'If TblBranchTableAdapter.CountThreePhaseLoad(CodeTextBoxSub.Text) > 0 Then
        '    load = TblBranchTableAdapter.TotalThreePhasePower(CodeTextBoxSub.Text)
        '    hrml = TblBranchTableAdapter.GetHRMLThreePhase(CodeTextBoxSub.Text)

        '    isf = (load + (0.25 * hrml)) / (Math.Sqrt(3) * voltage)

        '    CurrentRatingTextBox.Text = Math.Round(isf, 4)
        '    MinimumAmpacityTextBoxSub.Text = Math.Round(isf * 1.25, 4)

        '    OCPDRatingTextBoxSub.Text = OCPDRating(isf * 1.5)

        '    If WireSet(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), OCPDRatingTextBoxSub.Text, ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text) = 1 Then
        '        wirenumber = SetTextBoxSub.Text
        '    ElseIf WireSet(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), OCPDRatingTextBoxSub.Text, ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text) <= SetTextBoxSub.Text Then
        '        wirenumber = SetTextBoxSub.Text
        '    Else
        '        wirenumber = WireSet(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), OCPDRatingTextBoxSub.Text, ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text)
        '        SetTextBoxSub.Text = wirenumber
        '    End If

        '    setcurrent = OCPDRatingTextBoxSub.Text / wirenumber

        '    WireSizeTextBoxSub.Text = WireSize(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), setcurrent, ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text)

        '    ConduitSizeTextBoxSub.Text = ConduitSize(WireTypeComboBoxSub.Text, ConductorComboBoxSub.Text, GroundWireCheckBoxSub.Checked, "3", ConduitTypeComboBoxSub.Text, True)
        'Else
        '    phaseload = {TblBranchTableAdapter.TotalPhaseAPower(CodeTextBoxSub.Text), TblBranchTableAdapter.TotalPhaseBPower(CodeTextBoxSub.Text), TblBranchTableAdapter.TotalPhaseCPower(CodeTextBoxSub.Text)}
        '    If IsNothing(TblBranchTableAdapter.GetHRMLSinglePhase(CodeTextBoxSub.Text)) Then
        '        hrml = 0
        '    Else
        '        hrml = TblBranchTableAdapter.GetHRMLSinglePhase(CodeTextBoxSub.Text)
        '    End If

        '    Array.Sort(phaseload)

        '    isf = Math.Sqrt(3) * (phaseload(2) + (0.25 * hrml)) / voltage

        '    CurrentRatingTextBox.Text = Math.Round(isf, 4)
        '    MinimumAmpacityTextBoxSub.Text = Math.Round(isf * 1.25, 4)

        '    OCPDRatingTextBoxSub.Text = OCPDRating(isf * 1.5)

        '    If WireSet(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), OCPDRating(isf * 1.5), ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text) = 1 Then
        '        wirenumber = SetTextBoxSub.Text
        '    ElseIf WireSet(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBox.Text), OCPDRating(isf * 1.5), ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text) <= SetTextBoxSub.Text Then
        '        wirenumber = SetTextBoxSub.Text
        '    Else
        '        wirenumber = WireSet(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxSub.Text), OCPDRating(isf * 1.5), ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text)
        '        SetTextBoxSub.Text = wirenumber
        '    End If

        '    setcurrent = OCPDRatingTextBoxSub.Text / wirenumber

        '    If WireSize(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxSub.Text), setcurrent, ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text) = "2.0" Then
        '        WireSizeTextBoxSub.Text = "3.5"
        '    Else
        '        WireSizeTextBoxSub.Text = WireSize(ConductorComboBoxSub.Text, TblWireTableAdapter.GetNominalTemp(WireTypeComboBoxSub.Text), setcurrent, ConduitTypeComboBoxSub.Text, WireTypeComboBoxSub.Text)
        '    End If

        '    ConduitSizeTextBoxSub.Text = ConduitSize(WireTypeComboBoxSub.Text, ConductorComboBoxSub.Text, GroundWireCheckBoxSub.Checked, "3", ConduitTypeComboBoxSub.Text, False)
        'End If

        vd = VoltageDrop(isf / SetTextBoxSub.Text, DistancetoMainTextBox.Text, ConduitTypeComboBoxSub.Text, WireSizeTextBoxSub.Text, ConductorComboBoxSub.Text, ThreePhase, VoltageLevelComboBox.Text)

        Do While vd > 2
            Dim size = {"2.0", "3.5", "5.5", "8.0", "14", "22", "30", "38", "50", "60", "80", "100", "125", "150", "175", "200", "250", "325", "375", "400", "500"}
            For i As Integer = 0 To 20 Step 1
                If size(i) = WireSizeTextBoxSub.Text And i <> 20 Then
                    WireSizeTextBoxSub.Text = size(i + 1)
                    GoTo vdline
                ElseIf i = 20 Then
                    MessageBox.Show("Voltage drop exceeds standard voltage drop for largest available wire. Consider adding wire set.")
                    GoTo ErrorLine
                End If
            Next
vdline:     vd = VoltageDrop(isf / SetTextBoxSub.Text, DistancetoMainTextBox.Text, ConduitTypeComboBoxSub.Text, WireSizeTextBoxSub.Text, ConductorComboBoxSub.Text, ThreePhase, VoltageLevelComboBox.Text)
        Loop

        If NeutralCheckBox.Checked Then
            Dim ineutral As Decimal
            Dim neutralset As Integer

            ineutral = 0.7 * isf

            If WireSet(NeutralConductorComboBox.Text, TblWireTableAdapter.GetNominalTemp(NeutralWireComboBox.Text), ineutral, ConduitTypeComboBoxSub.Text, NeutralWireComboBox.Text) = 1 Then
                neutralset = NeutralSetTextBox.Text
            ElseIf WireSet(NeutralConductorComboBox.Text, TblWireTableAdapter.GetNominalTemp(NeutralWireComboBox.Text), ineutral, ConduitTypeComboBoxSub.Text, NeutralWireComboBox.Text) <= SetTextBoxSub.Text Then
                neutralset = NeutralSetTextBox.Text
            Else
                neutralset = WireSet(NeutralConductorComboBox.Text, TblWireTableAdapter.GetNominalTemp(NeutralWireComboBox.Text), ineutral, ConduitTypeComboBoxSub.Text, NeutralWireComboBox.Text)
                NeutralSetTextBox.Text = wirenumber
            End If

            NeutralSizeTextBox.Text = WireSize(NeutralConductorComboBox.Text, TblWireTableAdapter.GetNominalTemp(NeutralWireComboBox.Text), ineutral / neutralset, ConduitTypeComboBoxSub.Text, NeutralWireComboBox.Text)
        End If

        VoltageDropTextBox.Text = Math.Round(vd, 4)

        If GroundWireCheckBoxSub.Checked Then
            GroundWireSizeTextBoxSub.Text = GroundWire(OCPDRatingTextBoxSub.Text, GroundConductorComboBoxSub.Text)
        End If

        Balanced = False

        If PhaseTextBox.Text = "3" Then
            Z = LineImpedance(DistancetoMainTextBox.Text, ConduitTypeComboBoxSub.Text, WireSizeTextBoxSub.Text, ConductorComboBoxSub.Text, VoltageLevelComboBox.Text, SetTextBoxSub.Text) +
                LoadImpedance(Math.Sqrt(3) * VoltageLevelComboBox.Text * CurrentRatingTextBox.Text)
        Else
            Z = LineImpedance(DistancetoMainTextBox.Text, ConduitTypeComboBoxSub.Text, WireSizeTextBoxSub.Text, ConductorComboBoxSub.Text, VoltageLevelComboBox.Text, SetTextBoxSub.Text) +
                LoadImpedance(VoltageLevelComboBox.Text * CurrentRatingTextBox.Text)
        End If
        ImpedanceRealTextBox.Text = Math.Round(Z.Real, 4)
        ImpedanceImagTextBox.Text = Math.Round(Z.Imaginary, 4)
ErrorLine: End Sub

    Private Sub BtnBalance_Click(sender As Object, e As EventArgs) Handles BtnBalance.Click
        If TblBranchTableAdapter.CountBranch(CodeTextBoxSub.Text) = 0 Then
            MessageBox.Show("No branch circuit input yet. Add necessary branch circuit.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            GoTo ErrorLine
        End If

        TxtPhaseALoad.Text = TblBranchTableAdapter.TotalPhaseAPower(CodeTextBoxSub.Text)
        TxtPhaseBLoad.Text = TblBranchTableAdapter.TotalPhaseBPower(CodeTextBoxSub.Text)
        TxtPhaseCLoad.Text = TblBranchTableAdapter.TotalPhaseCPower(CodeTextBoxSub.Text)
        TxtThreePhaseLoad.Text = TblBranchTableAdapter.TotalThreePhasePower(CodeTextBoxSub.Text)
        TxtMotorSinglePhase.Text = TblBranchTableAdapter.GetHRMLSinglePhase(CodeTextBoxSub.Text)
        TxtMotorThreePhase.Text = TblBranchTableAdapter.GetHRMLThreePhase(CodeTextBoxSub.Text)

        If TxtPhaseALoad.Text = "" Then
            TxtPhaseALoad.Text = 0
        End If
        If TxtPhaseBLoad.Text = "" Then
            TxtPhaseBLoad.Text = 0
        End If
        If TxtPhaseCLoad.Text = "" Then
            TxtPhaseCLoad.Text = 0
        End If

        Dim ave, tolerance(2) As Decimal

        If PhaseTextBox.Text = "1" Or (PhaseTextBox.Text = "3" And TxtPhaseALoad.Text = "0" And TxtPhaseBLoad.Text = "0" And TxtPhaseCLoad.Text = "0") Then
            MessageBox.Show("Balancing not necessary.")
            Balanced = True
        Else
            ave = (CDec(TxtPhaseALoad.Text) + CDec(TxtPhaseBLoad.Text) + CDec(TxtPhaseCLoad.Text)) / 3

            tolerance(0) = Math.Abs(ave - CDec(TxtPhaseALoad.Text)) / ave * 100
            tolerance(1) = Math.Abs(ave - CDec(TxtPhaseBLoad.Text)) / ave * 100
            tolerance(2) = Math.Abs(ave - CDec(TxtPhaseCLoad.Text)) / ave * 100

            For i As Integer = 0 To 2
                If tolerance(i) > 2 Then
                    MessageBox.Show("Unbalanced subfeeder.")
                    GoTo ErrorLine
                End If
            Next i

            MessageBox.Show("Balanced subfeeder.")
            Balanced = True
        End If
ErrorLine: End Sub

    Private Sub BtnAddSub_Click(sender As Object, e As EventArgs) Handles BtnAddSub.Click
        TblSubfeederBindingSource.AddNew()

        ProjectCodeTextBoxSub.Text = ProjectCodeTextBox.Text
        DPCodeTextBox.Text = CodeTextBoxDP.Text
        SetTextBoxSub.Text = "1"

        NeutralCheckBox.Checked = False
        GroundWireCheckBoxSub.Checked = False
    End Sub

    Private Sub BtnSaveSub_Click(sender As Object, e As EventArgs) Handles BtnSaveSub.Click
        CodeTextBoxSub.Text = ProjectCodeTextBoxSub.Text + "-SF" + NumberTextBox.Text
        'Dim con As New SqlConnection(constring)

        'Dim query As String

        'If TblSubfeederTableAdapter.CountRow(CodeTextBoxSub.Text) = 0 Then
        '    query = "INSERT INTO tblSubfeeder (Code, ProjectCode, Number, DPCode, Voltage, CurrentRating, MinimumAmpacity, WireType, WireSize, [Set], Conductor, Neutral, NeutralWire, NeutralSize, NeutralConductor, NeutralSet, ConduitType, ConduitSize, OCPDRating, 
        '                BreakerType, DistancetoMain, VoltageDrop, GroundWire, GroundWireSize, GroundConductor, ImpedanceReal, ImpedanceImag) VALUES 
        '                (@Code,@ProjectCode,@Number,@DPCode,@Voltage,@CurrentRating,@MinimumAmpacity,@WireType,@WireSize,@Set,@Conductor,@Neutral,@NeutralWire,
        '                @NeutralSize,@NeutralConductor,@NeutralSet,@ConduitType,@ConduitSize,@OCPDRating,@BreakerType,@DistancetoMain,@VoltageDrop,@GroundWire,@GroundWireSize,@GroundConductor,@ImpedanceReal,@ImpedanceImag)"
        'Else
        '    query = "UPDATE tblSubfeeder SET Code = @Code, ProjectCode = @ProjectCode, Number = @Number, DPCode = @DPCode, Voltage = @Voltage, CurrentRating = @CurrentRating, MinimumAmpacity = @MinimumAmpacity, WireType = @WireType, 
        '                 WireSize = @WireSize, [Set] = @Set, Conductor = @Conductor, Neutral = @Neutral, NeutralWire = @NeutralWire, NeutralSize = @NeutralSize, NeutralConductor = @NeutralConductor, NeutralSet = @NeutralSet, 
        '                 ConduitType = @ConduitType, ConduitSize = @ConduitSize, OCPDRating = @OCPDRating, BreakerType = @BreakerType, DistancetoMain = @DistancetoMain, VoltageDrop = @VoltageDrop, GroundWire = @GroundWire, 
        '                 GroundWireSize = @GroundWireSize, GroundConductor = @GroundConductor, ImpedanceReal = @ImpedanceReal, ImpedanceImag = @ImpedanceImag WHERE (Code = @Code)"
        'End If

        'Dim cmd As New SqlCommand(query, con)

        Dim dr As DialogResult = MessageBox.Show("Do you wish to save?", "Save", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            'Try
            '    Validate()

            '    con.Open()

            '    cmd.Parameters.AddWithValue("@Code", CodeTextBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCodeTextBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@Number", NumberTextBox.Text)
            '    cmd.Parameters.AddWithValue("@DPCode", DPCodeTextBox.Text)
            '    cmd.Parameters.AddWithValue("@Voltage", VoltageComboBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@CurrentRating", CurrentRatingTextBox.Text)
            '    cmd.Parameters.AddWithValue("@MinimumAmpacity", MinimumAmpacityTextBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@WireType", WireTypeComboBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@WireSize", WireSizeTextBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@Set", SetTextBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@Conductor", ConductorComboBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@Neutral", NeutralCheckBox.CheckState)
            '    cmd.Parameters.AddWithValue("@NeutralWire", NeutralWireComboBox.Text)
            '    cmd.Parameters.AddWithValue("@NeutralSize", NeutralSizeTextBox.Text)
            '    cmd.Parameters.AddWithValue("@NeutralConductor", NeutralConductorComboBox.Text)
            '    cmd.Parameters.AddWithValue("@NeutralSet", NeutralSetTextBox.Text)
            '    cmd.Parameters.AddWithValue("@ConduitType", ConduitTypeComboBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@ConduitSize", ConduitSizeTextBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@OCPDRating", OCPDRatingTextBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@BreakerType", BreakerTypeTextBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@DistancetoMain", DistancetoMainTextBox.Text)
            '    cmd.Parameters.AddWithValue("@VoltageDrop", VoltageDropTextBox.Text)
            '    cmd.Parameters.AddWithValue("@GroundWire", GroundWireCheckBoxSub.CheckState)
            '    cmd.Parameters.AddWithValue("@GroundWireSize", GroundWireSizeTextBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@GroundConductor", GroundConductorComboBoxSub.Text)
            '    cmd.Parameters.AddWithValue("@ImpedanceReal", ImpedanceRealTextBox.Text)
            '    cmd.Parameters.AddWithValue("@ImpedanceImag", ImpedanceImagTextBox.Text)
            '    cmd.ExecuteNonQuery()

            '    MessageBox.Show("Record saved.")

            '    TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
            'Catch ex As Exception
            '    MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Finally
            '    con.Close()
            'End Try

            Try
                Validate()
                TblSubfeederBindingSource.EndEdit()
                TblSubfeederTableAdapter.Update(ESD_DatabaseDataSet.tblSubfeeder)

                MessageBox.Show("Record saved.")

                TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End If
    End Sub

    Private Sub BtnNextSub_Click(sender As Object, e As EventArgs) Handles BtnNextSub.Click
        Try
            TblSubfeederBindingSource.MoveNext()
            TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub BtnPreviousSub_Click(sender As Object, e As EventArgs) Handles BtnPreviousSub.Click
        Try
            TblSubfeederBindingSource.MovePrevious()
            TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub NeutralCheckBox_CheckStateChanged(sender As Object, e As EventArgs) Handles NeutralCheckBox.CheckStateChanged
        If NeutralCheckBox.Checked = True Then
            NeutralWireComboBox.Enabled = True
            NeutralConductorComboBox.Enabled = True
            NeutralSizeTextBox.ReadOnly = False
            NeutralSizeTextBox.BackColor = SystemColors.Window
            NeutralSetTextBox.ReadOnly = False
            NeutralSetTextBox.BackColor = SystemColors.Window
            NeutralSetTextBox.Text = 1
        Else
            NeutralSizeTextBox.ReadOnly = True
            NeutralSizeTextBox.BackColor = Color.Silver
            NeutralWireComboBox.Enabled = False
            NeutralConductorComboBox.Enabled = False
            NeutralSetTextBox.ReadOnly = True
            NeutralSetTextBox.BackColor = Color.Silver
            NeutralSetTextBox.Text = ""
        End If
    End Sub

    Private Sub BtnDeleteSub_Click(sender As Object, e As EventArgs) Handles BtnDeleteSub.Click
        Dim dr As DialogResult = MessageBox.Show("Do you wish to delete? Once deleted, the record will not be retrieved.", "Delete", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            Try
                TblSubfeederTableAdapter.DeletebyPrimary(CodeTextBoxSub.Text)

                MessageBox.Show("Record deleted.")

                TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub GroundWireCheckBoxSub_CheckStateChanged(sender As Object, e As EventArgs) Handles GroundWireCheckBoxSub.CheckStateChanged
        If GroundWireCheckBoxSub.Checked Then
            GroundWireSizeTextBoxSub.Enabled = True
            GroundConductorComboBoxSub.Enabled = True
        Else
            GroundWireSizeTextBoxSub.Enabled = False
            GroundConductorComboBoxSub.Enabled = False
        End If
    End Sub

#End Region

#Region "Branch Circuit"
    Private Sub BtnComputeBranch_Click(sender As Object, e As EventArgs) Handles BtnComputeBranch.Click
        Dim current, power, ocpd, setcurrent, vd As Decimal
        Dim truerating As Decimal = ConvertFraction(MotorRatingTextBox.Text)
        Dim wirenumber As Integer = 1
        Dim conduit As Integer
        Dim ThreePhase As Boolean

        If PhaseComboBox.Text = "3" Then
            ThreePhase = True
        Else
            ThreePhase = False
        End If

        'Do the data checking part
        If TypeComboBox.Text = "" Or VoltageLevelComboBox.Text = "" Or PhaseComboBox.Text = "" Or ((WireTypeComboBox.Text = "" Or ConductorComboBox.Text = "" Or ConduitTypeComboBox.Text = "") And TypeComboBox.SelectedIndex < 5) Then
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
                If MotorTypeComboBox.SelectedIndex = 2 And (VoltageLevelComboBox.SelectedIndex < 3 Or VoltageLevelComboBox.SelectedIndex = 4) Then
                    MessageBox.Show("Invalid voltage rating for given motor type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    GoTo ErrorLine

                ElseIf MotorTypeComboBox.SelectedIndex = 2 And CboRatingUnit.SelectedIndex = 0 And ((CDec(MotorRatingTextBox.Text) < 25) Or
                        (CDec(MotorRatingTextBox.Text) > 200) Or
                        (CDec(MotorRatingTextBox.Text) < 60 And VoltageLevelComboBox.SelectedIndex = 7)) Then
                    MessageBox.Show("Invalid motor rating for given voltage.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    GoTo ErrorLine

                ElseIf (VoltageLevelComboBox.SelectedIndex = 7 And truerating < 60) Or
                        (VoltageLevelComboBox.SelectedIndex = 0 And truerating > 50) Or
                        (VoltageLevelComboBox.SelectedIndex <= 2 And truerating > 200) Then
                    MessageBox.Show("Invalid motor rating for given voltage.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    GoTo ErrorLine

                End If

            ElseIf PhaseComboBox.Text IsNot "3" And MotorTypeComboBox.SelectedIndex = 2 Then
                MessageBox.Show("Given motor phase only valid for three-phase system.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                GoTo ErrorLine
            End If

        ElseIf (TypeComboBox.SelectedIndex = 3 Or TypeComboBox.SelectedIndex = 4) And PowerRatingTextBox.Text = "" Then
            MessageBox.Show("Fill in the required fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            GoTo ErrorLine

        ElseIf TypeComboBox.SelectedIndex < 5 And DistancetoSFTextBox.Text = "" Then
            MessageBox.Show("Fill in the required fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            GoTo ErrorLine
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
            current = CInt(PowerRatingTextBox.Text) / CInt(VoltageLevelComboBox.Text)
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

            current = CInt(PowerRatingTextBox.Text) / CInt(VoltageLevelComboBox.Text)
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
            If Size = "2.0" And ConductorComboBox.Text = "Cu" Then
                WireSizeTextBox1.Text = "3.5"
            ElseIf Size = "3.5" And ConductorComboBox.Text IsNot "Cu" Then
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

            conduit = ConduitSize(WireTypeComboBox.Text, WireSizeTextBox1.Text, GroundWireCheckBox1.Checked, PhaseComboBox.Text, ConduitTypeComboBox.Text, False)

            If conduit = -1 Then
                MessageBox.Show("No conduit size available for given wire and conduit. Consider adding wire set.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                GoTo ErrorLine
            Else
                ConduitSizeTextBox1.Text = conduit
            End If

            'Computation for Motors
        ElseIf TypeComboBox.SelectedIndex = 2 Then
            If CboRatingUnit.SelectedIndex = 0 Then
                If PhaseComboBox.Text = "3" Then
                    If MotorTypeComboBox.SelectedIndex = 2 Then
                        current = ThreePhaseSynchronous(truerating, VoltageLevelComboBox.SelectedIndex)
                        If current = -1 Then
                            MessageBox.Show("Given motor rating not among standard.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            GoTo ErrorLine
                        End If
                    Else
                        current = ThreePhaseInduction(truerating, VoltageLevelComboBox.SelectedIndex)
                        If current = -1 Then
                            MessageBox.Show("Given motor rating not among standard.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            GoTo ErrorLine
                        End If
                    End If
                    FullLoadCurrentTextBox.Text = Math.Round(current, 4)

                    power = current * CInt(VoltageLevelComboBox.Text) * Math.Sqrt(3) * CInt(TxtMotorItem.Text)
                    PowerRatingTextBox.Text = Math.Round(power, 4)
                Else
                    current = SinglePhaseCurrent(truerating, VoltageLevelComboBox.SelectedIndex)
                    power = CInt(VoltageLevelComboBox.Text) * current * CInt(TxtMotorItem.Text)
                    FullLoadCurrentTextBox.Text = CStr(Math.Round(current, 4))
                    PowerRatingTextBox.Text = CStr(Math.Round(power, 4))
                End If
            ElseIf CboRatingUnit.SelectedIndex = 1 Then
                If PhaseComboBox.Text = "3" Then
                    power = CDec(MotorRatingTextBox.Text) * 1000 * CInt(TxtMotorItem.Text)
                    current = power / (CInt(VoltageLevelComboBox.Text) * Math.Sqrt(3))
                Else
                    power = CDec(MotorRatingTextBox.Text) * 1000 * CInt(TxtMotorItem.Text)
                    current = power / CInt(VoltageLevelComboBox.Text)
                End If

                PowerRatingTextBox.Text = CStr(Math.Round(power, 4))
                FullLoadCurrentTextBox.Text = CStr(Math.Round(current, 4))
            Else
                If PhaseComboBox.Text = "3" Then
                    power = CDec(MotorRatingTextBox.Text) * CInt(TxtMotorItem.Text)
                    current = power / (CInt(VoltageLevelComboBox.Text) * Math.Sqrt(3))
                Else
                    power = CDec(MotorRatingTextBox.Text) * CInt(TxtMotorItem.Text)
                    current = power / CInt(VoltageLevelComboBox.Text)
                End If

                PowerRatingTextBox.Text = CStr(Math.Round(power, 4))
                FullLoadCurrentTextBox.Text = CStr(Math.Round(current, 4))
            End If

            MotorRatingTextBox.Text = Math.Round(power / CInt(TxtMotorItem.Text), 4)
            CboRatingUnit.SelectedIndex = 2

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
            If Size = "2.0" And ConductorComboBox.Text = "Cu" Then
                WireSizeTextBox1.Text = "3.5"
            ElseIf Size = "3.5" And ConductorComboBox.Text IsNot "Cu" Then
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

            conduit = ConduitSize(WireTypeComboBox.Text, WireSizeTextBox1.Text, GroundWireCheckBox1.Checked, PhaseComboBox.Text, ConduitTypeComboBox.Text, False)

            If conduit = -1 Then
                MessageBox.Show("No conduit size available for given wire and conduit. Consider adding wire set.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                GoTo ErrorLine
            Else
                ConduitSizeTextBox1.Text = conduit
            End If

            'Computation for Non-Motors
        ElseIf TypeComboBox.SelectedIndex = 3 Or TypeComboBox.SelectedIndex = 4 Then

            power = CDec(PowerRatingTextBox.Text) * CInt(TxtMotorItem.Text)

            current = power / CInt(VoltageLevelComboBox.Text)
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
            If Size = "2.0" And ConductorComboBox.Text = "Cu" Then
                WireSizeTextBox1.Text = "3.5"
            ElseIf Size = "3.5" And ConductorComboBox.Text IsNot "Cu" Then
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

            conduit = ConduitSize(WireTypeComboBox.Text, WireSizeTextBox1.Text, GroundWireCheckBox1.Checked, PhaseComboBox.Text, ConduitTypeComboBox.Text, False)

            If conduit = -1 Then
                MessageBox.Show("No conduit size available for given wire and conduit. Consider adding wire set.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                GoTo ErrorLine
            Else
                ConduitSizeTextBox1.Text = conduit
            End If
        ElseIf TypeComboBox.SelectedIndex = 5 Then
            PowerRatingTextBox.Text = 1200
            If TypeComboBox1.SelectedIndex <= 1 Then
                OCPDRatingTextBox1.Text = 15
            Else
                OCPDRatingTextBox1.Text = 20
            End If
            FullLoadCurrentTextBox.Text = ""
            MinimumAmpacityTextBox.Text = ""
            WireSizeTextBox1.Text = ""
            ConduitSizeTextBox1.Text = ""
        ElseIf TypeComboBox.SelectedIndex = 6 Then
            If PowerRatingTextBox.Text = "" Then
                PowerRatingTextBox.Text = "1500"
            End If

            If PhaseComboBox.Text IsNot "3" Then
                ocpd = OCPDRating(PowerRatingTextBox.Text * 1.5 / VoltageLevelComboBox.Text)
            Else
                ocpd = OCPDRating(PowerRatingTextBox.Text * 1.5 / (VoltageLevelComboBox.Text / Math.Sqrt(3)))
            End If

            If ocpd = 15 Then
                OCPDRatingTextBox1.Text = 20
            Else
                OCPDRatingTextBox1.Text = ocpd
            End If
            FullLoadCurrentTextBox.Text = ""
            MinimumAmpacityTextBox.Text = ""
            WireSizeTextBox1.Text = ""
            ConduitSizeTextBox1.Text = ""
        End If

        If GroundWireCheckBox1.Checked = True Then
            GroundWireSizeTextBox.Text = GroundWire(CInt(OCPDRatingTextBox1.Text), GroundConductorComboBox.Text)
        End If

        If TypeComboBox.SelectedIndex < 5 Then
            vd = VoltageDrop(FullLoadCurrentTextBox.Text / SetTextBox.Text, DistancetoSFTextBox.Text, ConduitTypeComboBox.Text, WireSizeTextBox1.Text, ConductorComboBox.Text, ThreePhase, VoltageLevelComboBox.Text)

            Do While vd > 3
                Dim size = {"2.0", "3.5", "5.5", "8.0", "14", "22", "30", "38", "50", "60", "80", "100", "125", "150", "175", "200", "250", "325", "375", "400", "500"}
                For i As Integer = 0 To 20 Step 1
                    If size(i) = WireSizeTextBox1.Text And i <> 20 Then
                        WireSizeTextBox1.Text = size(i + 1)
                        GoTo vdline
                    ElseIf i = 20 Then
                        MessageBox.Show("Voltage drop exceeds standard voltage drop for largest available wire. Consider adding wire set.")
                        GoTo ErrorLine
                    End If
                Next
vdline:         vd = VoltageDrop(FullLoadCurrentTextBox.Text / SetTextBox.Text, DistancetoSFTextBox.Text, ConduitTypeComboBox.Text, WireSizeTextBox1.Text, ConductorComboBox.Text, ThreePhase, VoltageLevelComboBox.Text)
            Loop

            VoltageDropTextBox1.Text = Math.Round(vd, 4)
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
            Case 5 To 8
                TxtRatingLighting1.Text = 550
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
            Case 5 To 8
                TxtRatingLighting2.Text = 550
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
            Case 5 To 8
                TxtRatingLighting3.Text = 550
            Case Else
                TxtRatingLighting3.Text = 0
        End Select
    End Sub

    Private Sub TypeComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TypeComboBox.SelectedIndexChanged
        If TypeComboBox.SelectedIndex = 6 Then
            PowerRatingTextBox.ReadOnly = False
            PowerRatingTextBox.BackColor = SystemColors.Window
        Else
            PowerRatingTextBox.ReadOnly = True
            PowerRatingTextBox.BackColor = Color.Silver
        End If

        If TypeComboBox.SelectedIndex = 2 Then
            MotorRatingTextBox.ReadOnly = False
            MotorRatingTextBox.BackColor = SystemColors.Window
            CboRatingUnit.Enabled = True
            CboRatingUnit.SelectedIndex = 0
            MotorTypeComboBox.Enabled = True
            MotorTypeComboBox.SelectedIndex = 0
            GrpLighting.Enabled = False
            GrpPower.Enabled = False
            TxtMotorItem.Text = "1"
            PowerRatingTextBox.ReadOnly = True
            PowerRatingTextBox.BackColor = Color.Silver
            WireTypeComboBox.Enabled = True
            ConductorComboBox.Enabled = True
            ConduitTypeComboBox.Enabled = True

            For Each c As Control In GrpLighting.Controls
                If c.GetType Is GetType(TextBox) Or c.GetType Is GetType(ComboBox) Then
                    c.Text = ""
                End If
            Next

            For Each c As Control In GrpPower.Controls
                If c.GetType Is GetType(TextBox) Or c.GetType Is GetType(ComboBox) Then
                    c.Text = ""
                End If
            Next

        Else
            MotorRatingTextBox.ReadOnly = True
            MotorRatingTextBox.BackColor = Color.Silver
            CboRatingUnit.SelectedIndex = -1
            CboRatingUnit.Enabled = False
            MotorTypeComboBox.SelectedIndex = -1
            MotorTypeComboBox.Enabled = False

            If TypeComboBox.SelectedIndex = 0 Then
                GrpPower.Enabled = False
                GrpLighting.Enabled = True
                PowerRatingTextBox.ReadOnly = True
                PowerRatingTextBox.BackColor = Color.Silver
                TxtMotorItem.ReadOnly = True
                TxtMotorItem.BackColor = Color.Silver
                TxtMotorItem.Text = ""
                WireTypeComboBox.Enabled = True
                ConductorComboBox.Enabled = True
                ConduitTypeComboBox.Enabled = True

                For Each c As Control In GrpPower.Controls
                    If c.GetType Is GetType(TextBox) Or c.GetType Is GetType(ComboBox) Then
                        c.Text = ""
                    End If
                Next

                TxtRatingLighting1.Text = "0"
                TxtRatingLighting2.Text = "0"
                TxtRatingLighting3.Text = "0"
                TxtItemsLighting1.Text = "0"
                TxtItemsLighting2.Text = "0"
                TxtItemsLighting3.Text = "0"

            ElseIf TypeComboBox.SelectedIndex = 1 Then
                GrpPower.Enabled = True
                GrpLighting.Enabled = False
                PowerRatingTextBox.ReadOnly = True
                PowerRatingTextBox.BackColor = Color.Silver
                TxtMotorItem.ReadOnly = True
                TxtMotorItem.BackColor = Color.Silver
                TxtMotorItem.Text = ""
                WireTypeComboBox.Enabled = True
                ConductorComboBox.Enabled = True
                ConduitTypeComboBox.Enabled = True

                For Each c As Control In GrpLighting.Controls
                    If c.GetType Is GetType(TextBox) Or c.GetType Is GetType(ComboBox) Then
                        c.Text = ""
                    End If
                Next

                TxtItemsPower1.Text = "0"
                TxtItemsPower2.Text = "0"
                TxtItemsPower3.Text = "0"
                TxtItemsPower4.Text = "0"

            ElseIf TypeComboBox.SelectedIndex = 3 Or TypeComboBox.SelectedIndex = 4 Then
                GrpPower.Enabled = False
                GrpLighting.Enabled = False
                PowerRatingTextBox.ReadOnly = False
                PowerRatingTextBox.BackColor = Color.Silver
                TxtMotorItem.Text = "1"
                WireTypeComboBox.Enabled = True
                ConductorComboBox.Enabled = True
                ConduitTypeComboBox.Enabled = True

                For Each c As Control In GrpLighting.Controls
                    If c.GetType Is GetType(TextBox) Or c.GetType Is GetType(ComboBox) Then
                        c.Text = ""
                    End If
                Next

                For Each c As Control In GrpPower.Controls
                    If c.GetType Is GetType(TextBox) Or c.GetType Is GetType(ComboBox) Then
                        c.Text = ""
                    End If
                Next

            ElseIf TypeComboBox.SelectedIndex >= 5 Then
                GrpPower.Enabled = False
                GrpLighting.Enabled = False
                WireTypeComboBox.Enabled = False
                ConductorComboBox.Enabled = False
                ConduitTypeComboBox.Enabled = False

                For Each c As Control In GrpLighting.Controls
                    If c.GetType Is GetType(TextBox) Or c.GetType Is GetType(ComboBox) Then
                        c.Text = ""
                    End If
                Next

                For Each c As Control In GrpPower.Controls
                    If c.GetType Is GetType(TextBox) Or c.GetType Is GetType(ComboBox) Then
                        c.Text = ""
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub BtnAddBranch_Click(sender As Object, e As EventArgs) Handles BtnAddBranch.Click
        TblBranchBindingSource.AddNew()
        ProjectTextBox.Text = ProjectCodeTextBox.Text
        SubfeederTextBox.Text = CodeTextBoxSub.Text
        SetTextBox.Text = "1"
        GroundWireCheckBox1.CheckState = False
        TypeComboBox.Enabled = True

        If TypeComboBox1.SelectedIndex < 2 Then
            PhaseComboBox.Text = "A"
        End If
    End Sub

    Private Sub BtnSaveBranch_Click(sender As Object, e As EventArgs) Handles BtnSaveBranch.Click
        CodeTextBox.Text = SubfeederTextBox.Text + "-BC" + CircuitNoTextBox.Text

        Dim dr As DialogResult = MessageBox.Show("Do you wish to save?", "Save", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            Try
                Validate()
                TblBranchBindingSource.EndEdit()
                TblBranchTableAdapter.Update(ESD_DatabaseDataSet.tblBranch)

                MessageBox.Show("Record saved.")

                For Each c As Control In GrpLighting.Controls
                    If c.GetType Is GetType(TextBox) Or c.GetType Is GetType(ComboBox) Then
                        c.Text = ""
                    End If
                Next

                For Each c As Control In GrpPower.Controls
                    If c.GetType Is GetType(TextBox) Or c.GetType Is GetType(ComboBox) Then
                        c.Text = ""
                    End If
                Next
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If

    End Sub

    Private Sub BtnDeleteBranch_Click(sender As Object, e As EventArgs) Handles BtnDeleteBranch.Click
        Dim dr As DialogResult = MessageBox.Show("Do you wish to delete? Once deleted, the record will not be retrieved.", "Delete", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            Try
                TblBranchTableAdapter.DeleteByPrimary(CodeTextBox.Text)

                MessageBox.Show("Record deleted.")

                TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub BtnNextBranch_Click(sender As Object, e As EventArgs) Handles BtnNextBranch.Click
        TblBranchBindingSource.MoveNext()
        CboRatingUnit.SelectedIndex = 2
    End Sub

    Private Sub BtnPreviousBranch_Click(sender As Object, e As EventArgs) Handles BtnPreviousBranch.Click
        TblBranchBindingSource.MovePrevious()
        CboRatingUnit.SelectedIndex = 2
    End Sub

    Private Sub GroundWireCheckBox1_CheckStateChanged(sender As Object, e As EventArgs) Handles GroundWireCheckBox1.CheckStateChanged
        If GroundWireCheckBox1.Checked = True Then
            GroundConductorComboBox.Enabled = True
            GroundConductorComboBox.SelectedIndex = 0
        Else
            GroundConductorComboBox.Enabled = False
            GroundConductorComboBox.SelectedIndex = -1
        End If
    End Sub
#End Region

#Region "Transformer/Generator"
    Private Sub BtnComputeTG_Click(sender As Object, e As EventArgs) Handles BtnComputeTG.Click
        If IsNothing(TblDistributionTableAdapter.CurrentSum(ProjectCodeTextBox.Text)) Then
            MessageBox.Show("Compute for distribution panelboards before proceeding to the transformer/generator computation.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            GoTo ErrorLine
        End If

        If TypeComboBox1.SelectedIndex >= 2 Then
            If DemandFactorTextBox.Text = "" Then
                DemandFactorTextBox.Text = 0.6
            End If

            If DiversityFactorTextBox.Text = "" Then
                DiversityFactorTextBox.Text = 1.3
            End If

            If TransformerXRRatioTextBox.Text = "" Then
                TransformerXRRatioTextBox.Text = 10
            End If

            If TransformerZpuTextBox.Text = "" Then
                TransformerZpuTextBox.Text = 5.75
            End If

            If NumberofGeneratorsTextBox.Text = "" Or NumberofGeneratorsTextBox.Text = "0" Then
                NumberofGeneratorsTextBox.Text = 1
            End If

            Dim singlephaseload, threephaseload, phaseload(2) As Decimal
            Dim voltage As Integer = VoltageLevelComboBox.Text

            If PhaseTextBox.Text = "3" Then
                phaseload = {TblBranchTableAdapter.TotalLoadperPhase(ProjectCodeTextBox.Text, "A"),
                    TblBranchTableAdapter.TotalLoadperPhase(ProjectCodeTextBox.Text, "B"),
                    TblBranchTableAdapter.TotalLoadperPhase(ProjectCodeTextBox.Text, "C")}

                Array.Sort(phaseload)

                singlephaseload = phaseload(2)

                threephaseload = TblBranchTableAdapter.TotalLoadperPhase(ProjectCodeTextBox.Text, "3")

                TCL = (3 * singlephaseload) + threephaseload

                TotalLoadTextBox.Text = Math.Round(TCL, 4)
            Else
                singlephaseload = TblBranchTableAdapter.TotalLoadperPhase(ProjectTextBox.Text, "A")
                TCL = singlephaseload
                TotalLoadTextBox.Text = Math.Round(singlephaseload, 4)
            End If

            TransformerRatingTextBox.Text = CStr(TransformerRating(TCL * CDec(DemandFactorTextBox.Text) / CDec(DiversityFactorTextBox.Text)))

            GeneratorRatingTextBox.Text = CStr(TransformerRating(TCL * 1.25 / CDec(NumberofGeneratorsTextBox.Text)))
        Else
            GeneratorRatingTextBox.Text = CStr(TransformerRating(TCL * 1.25 / CDec(NumberofGeneratorsTextBox.Text)))
        End If

ErrorLine: End Sub

    Private Sub BtnAddTG_Click(sender As Object, e As EventArgs) Handles BtnAddTG.Click
        TblTransGenBindingSource.AddNew()
        ProjectCodeTextBox3.Text = ProjectCodeTextBox.Text
    End Sub

    Private Sub BtnSaveTG_Click(sender As Object, e As EventArgs) Handles BtnSaveTG.Click
        Dim dr As DialogResult = MessageBox.Show("Do you wish to save?", "Save", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            Try
                Validate()
                TblTransGenBindingSource.EndEdit()
                TblTransGenTableAdapter.Update(ESD_DatabaseDataSet.tblTransGen)

                MessageBox.Show("Record saved.")
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else

        End If
    End Sub

    Private Sub BtnDeleteTG_Click(sender As Object, e As EventArgs) Handles BtnDeleteTG.Click
        Dim dr As DialogResult = MessageBox.Show("Do you wish to delete? Once deleted, the record will not be retrieved.", "Delete", MessageBoxButtons.YesNo)
        If dr = DialogResult.Yes Then
            Try
                TblTransGenTableAdapter.DeletebyPrimary(ProjectCodeTextBox3.Text)

                MessageBox.Show("Record deleted.")
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub
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
        ElseIf InputNfo = Nothing Or InputNfo = " " Then
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

    Private Sub IntegerKeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtItemsPower4.KeyPress, TxtItemsPower3.KeyPress, TxtItemsPower2.KeyPress, TxtItemsPower1.KeyPress, TxtItemsLighting3.KeyPress, TxtItemsLighting2.KeyPress, TxtItemsLighting1.KeyPress, SetTextBoxSub.KeyPress, SetTextBoxDP.KeyPress, SetTextBox1.KeyPress, SetTextBox.KeyPress, PhaseTextBox.KeyPress, NumberTextBox.KeyPress, NumberofGeneratorsTextBox.KeyPress, DPNumberTextBox.KeyPress, CircuitNoTextBox.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub NumericKeyPress(sender As Object, e As KeyPressEventArgs) Handles XRRatioTextBox.KeyPress, TxtRatingLighting3.KeyPress, TxtRatingLighting2.KeyPress, TxtRatingLighting1.KeyPress, TransformerZpuTextBox.KeyPress, TransformerXRRatioTextBox.KeyPress, ShortCircuitCapTextBox.KeyPress, DiversityFactorTextBox.KeyPress, DistancetoSETextBox.KeyPress, DistancetoMainTextBoxDP.KeyPress, DistancetoMainTextBox.KeyPress, DemandFactorTextBox.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not e.KeyChar = "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub NumericTextChange(sender As Object, e As EventArgs) Handles TxtItemsPower4.TextChanged, TxtItemsPower3.TextChanged, TxtItemsPower2.TextChanged, TxtItemsPower1.TextChanged, TxtItemsLighting3.TextChanged, TxtItemsLighting2.TextChanged, TxtItemsLighting1.TextChanged, SetTextBoxSub.TextChanged, SetTextBoxDP.TextChanged, SetTextBox1.TextChanged, SetTextBox.TextChanged, PhaseTextBox.TextChanged, NumberTextBox.TextChanged, NumberofGeneratorsTextBox.TextChanged, DPNumberTextBox.TextChanged, CircuitNoTextBox.TextChanged
        Dim digitsOnly As Regex = New Regex("[^\d]")
        sender.Text = digitsOnly.Replace(sender.Text, "")
    End Sub

    Private Sub IlluminationCalculatorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IlluminationCalculatorToolStripMenuItem.Click
        'MessageBox.Show("Under construction.")

        Dim f As Form = FormIllu
        f.ShowDialog()
        'Hide()

    End Sub

    Private Sub ManualToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ManualToolStripMenuItem.Click
        Process.Start(Application.StartupPath & "\User-Manual.pdf")
    End Sub

    Private Sub AboutUsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutUsToolStripMenuItem.Click
        Dim f As New FormAbout
        f.ShowDialog()
    End Sub

    Private Report As FormReport

    Private Sub GenerateReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerateReportToolStripMenuItem.Click
        'Dim f As New FormReport
        'f.ProjectCode = ProjectCodeTextBox.Text

        'f.Show()
        'Hide()
        Dim initialValue As String
        initialValue = ProjectCodeTextBox.Text

        Report = New FormReport(initialValue)

        Report.ShowDialog()
    End Sub

    Private Sub FormMain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Application.Exit()
    End Sub

    Private Sub LoadRecordsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadRecordsToolStripMenuItem.Click
        TblMainFeederTableAdapter.FillByProject(ESD_DatabaseDataSet.tblMainFeeder, ProjectCodeTextBox.Text)
        TblTransGenTableAdapter.FillByProject(ESD_DatabaseDataSet.tblTransGen, ProjectCodeTextBox.Text)
        TblDistributionTableAdapter.FillByProject(ESD_DatabaseDataSet.tblDistribution, ProjectCodeTextBox.Text)
        TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
        TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        Select Case CboSearch.SelectedIndex
            Case 0
                If TblProjectTableAdapter.CountCode(TxtSearch.Text) > 0 Then
                    TblProjectTableAdapter.FillByProject(ESD_DatabaseDataSet.tblProject, TxtSearch.Text)
                    TblMainFeederTableAdapter.FillByProject(ESD_DatabaseDataSet.tblMainFeeder, ProjectCodeTextBox.Text)
                    TblDistributionTableAdapter.FillByProject(ESD_DatabaseDataSet.tblDistribution, ProjectCodeTextBox.Text)
                    TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
                    TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
                    TblTransGenTableAdapter.FillByProject(ESD_DatabaseDataSet.tblTransGen, ProjectCodeTextBox.Text)
                    MessageBox.Show("Search results ended. " + CStr(TblProjectTableAdapter.CountCode(TxtSearch.Text)) + " record/s found.")
                Else
                    MessageBox.Show("No records found.")
                End If
            Case 1
                If TblProjectTableAdapter.CountName(TxtSearch.Text) > 0 Then
                    TblProjectTableAdapter.FillByProjectName(ESD_DatabaseDataSet.tblProject, TxtSearch.Text)
                    TblMainFeederTableAdapter.FillByProject(ESD_DatabaseDataSet.tblMainFeeder, ProjectCodeTextBox.Text)
                    TblDistributionTableAdapter.FillByProject(ESD_DatabaseDataSet.tblDistribution, ProjectCodeTextBox.Text)
                    TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
                    TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
                    TblTransGenTableAdapter.FillByProject(ESD_DatabaseDataSet.tblTransGen, ProjectCodeTextBox.Text)
                    MessageBox.Show("Search results ended. " + CStr(TblProjectTableAdapter.CountName(TxtSearch.Text)) + " record/s found.")
                Else
                    MessageBox.Show("No records found.")
                End If
            Case 2
                If TblProjectTableAdapter.CountOwner(TxtSearch.Text) > 0 Then
                    TblProjectTableAdapter.FillByOwner(ESD_DatabaseDataSet.tblProject, TxtSearch.Text)
                    TblMainFeederTableAdapter.FillByProject(ESD_DatabaseDataSet.tblMainFeeder, ProjectCodeTextBox.Text)
                    TblDistributionTableAdapter.FillByProject(ESD_DatabaseDataSet.tblDistribution, ProjectCodeTextBox.Text)
                    TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
                    TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
                    TblTransGenTableAdapter.FillByProject(ESD_DatabaseDataSet.tblTransGen, ProjectCodeTextBox.Text)
                    MessageBox.Show("Search results ended. " + CStr(TblProjectTableAdapter.CountOwner(TxtSearch.Text)) + " record/s found.")
                Else
                    MessageBox.Show("No records found.")
                End If
            Case 3
                If TblDistributionTableAdapter.VerifySearch(TxtSearch.Text, ProjectCodeTextBox.Text) = 1 Then
                    TblDistributionTableAdapter.FillByCode(ESD_DatabaseDataSet.tblDistribution, TxtSearch.Text, ProjectCodeTextBox.Text)
                    TblSubfeederTableAdapter.FillByDP(ESD_DatabaseDataSet.tblSubfeeder, CodeTextBoxDP.Text)
                    TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
                    MessageBox.Show("Search results ended. " + CStr(TblDistributionTableAdapter.VerifySearch(TxtSearch.Text, ProjectCodeTextBox.Text)) + " record/s found.")
                Else
                    MessageBox.Show("No records found.")
                End If
            Case 4
                If TblSubfeederTableAdapter.VerifySearch(TxtSearch.Text, CodeTextBoxDP.Text) = 1 Then
                    TblSubfeederTableAdapter.FillByCode(ESD_DatabaseDataSet.tblSubfeeder, TxtSearch.Text, CodeTextBoxDP.Text)
                    TblBranchTableAdapter.FillBySubfeeder(ESD_DatabaseDataSet.tblBranch, CodeTextBoxSub.Text)
                    MessageBox.Show("Search results ended. " + CStr(TblSubfeederTableAdapter.VerifySearch(TxtSearch.Text, CodeTextBoxDP.Text)) + " record/s found.")
                Else
                    MessageBox.Show("No records found.")
                End If
            Case 5
                If TblBranchTableAdapter.VerifySearch(TxtSearch.Text, CodeTextBoxSub.Text) = 1 Then
                    TblBranchTableAdapter.FillByCode(ESD_DatabaseDataSet.tblBranch, TxtSearch.Text, CodeTextBoxSub.Text)
                    MessageBox.Show("Search results ended. " + CStr(TblBranchTableAdapter.VerifySearch(TxtSearch.Text, CodeTextBoxSub.Text)) + " record/s found.")
                Else
                    MessageBox.Show("No records found.")
                End If
        End Select
    End Sub
#End Region
End Class