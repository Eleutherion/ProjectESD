Public Class FormReport
    Private Sub FormReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TblProjectTableAdapter.FillByProject(ESD_DatabaseDataSet.tblProject, ProjectCode)
        TblBranchTableAdapter.FillByProject(ESD_DatabaseDataSet.tblBranch, ProjectCode)
        TblSubfeederTableAdapter.FillByProject(ESD_DatabaseDataSet.tblSubfeeder, ProjectCode)
        TblDistributionTableAdapter.FillByProject(ESD_DatabaseDataSet.tblDistribution, ProjectCode)
        TblMainFeederTableAdapter.FillByProject(ESD_DatabaseDataSet.tblMainFeeder, ProjectCode)
        TblTransGenTableAdapter.FillByProject(ESD_DatabaseDataSet.tblTransGen, ProjectCode)

        ReportESD1.SetDataSource(ESD_DatabaseDataSet)
    End Sub

    Public Property ProjectCode As String

    Private Sub FormReport_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim f As New FormMain
        f.Show()

        Dispose()
    End Sub
End Class