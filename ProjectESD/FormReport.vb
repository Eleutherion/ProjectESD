Public Class FormReport
    Private Property ProjectCode As String

    Private Sub FormReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public WriteOnly Property Code() As String
        Set(ByVal value As String)
            ProjectCode = value
        End Set
    End Property

    Public Sub New(ByVal CodeValue As String)
        MyBase.New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Code = CodeValue

        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblProject' table. You can move, or remove it, as needed.
        TblProjectTableAdapter.FillByProject(ESD_DatabaseDataSet.tblProject, CodeValue)
        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblTransGen' table. You can move, or remove it, as needed.
        TblTransGenTableAdapter.FillByProject(ESD_DatabaseDataSet.tblTransGen, CodeValue)
        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblMainFeeder' table. You can move, or remove it, as needed.
        TblMainFeederTableAdapter.FillByProject(ESD_DatabaseDataSet.tblMainFeeder, CodeValue)
        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblDistribution' table. You can move, or remove it, as needed.
        TblDistributionTableAdapter.FillByProject(ESD_DatabaseDataSet.tblDistribution, CodeValue)
        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblSubfeeder' table. You can move, or remove it, as needed.
        TblSubfeederTableAdapter.FillByProject(ESD_DatabaseDataSet.tblSubfeeder, CodeValue)
        'TODO: This line of code loads data into the 'ESD_DatabaseDataSet.tblBranch' table. You can move, or remove it, as needed.
        TblBranchTableAdapter.FillByProject(ESD_DatabaseDataSet.tblBranch, CodeValue)

        reportProject1.SetDataSource(ESD_DatabaseDataSet)
    End Sub

    Private Sub FormReport_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Dispose()
    End Sub
End Class