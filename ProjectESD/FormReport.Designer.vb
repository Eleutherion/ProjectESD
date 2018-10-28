<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormReport
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormReport))
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.reportProject1 = New ProjectESD.ReportProject()
        Me.ESD_DatabaseDataSet = New ProjectESD.ESD_DatabaseDataSet()
        Me.tblBranchBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.TblBranchTableAdapter = New ProjectESD.ESD_DatabaseDataSetTableAdapters.tblBranchTableAdapter()
        Me.tblSubfeederBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.TblSubfeederTableAdapter = New ProjectESD.ESD_DatabaseDataSetTableAdapters.tblSubfeederTableAdapter()
        Me.tblDistributionBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.TblDistributionTableAdapter = New ProjectESD.ESD_DatabaseDataSetTableAdapters.tblDistributionTableAdapter()
        Me.tblMainFeederBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.TblMainFeederTableAdapter = New ProjectESD.ESD_DatabaseDataSetTableAdapters.tblMainFeederTableAdapter()
        Me.tblTransGenBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.TblTransGenTableAdapter = New ProjectESD.ESD_DatabaseDataSetTableAdapters.tblTransGenTableAdapter()
        Me.tblProjectBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.TblProjectTableAdapter = New ProjectESD.ESD_DatabaseDataSetTableAdapters.tblProjectTableAdapter()
        CType(Me.ESD_DatabaseDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblBranchBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblSubfeederBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblDistributionBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblMainFeederBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblTransGenBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblProjectBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = 0
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.Cursor = System.Windows.Forms.Cursors.Default
        Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ReportSource = Me.reportProject1
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(1305, 784)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'ESD_DatabaseDataSet
        '
        Me.ESD_DatabaseDataSet.DataSetName = "ESD_DatabaseDataSet"
        Me.ESD_DatabaseDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'tblBranchBindingSource
        '
        Me.tblBranchBindingSource.DataMember = "tblBranch"
        Me.tblBranchBindingSource.DataSource = Me.ESD_DatabaseDataSet
        '
        'TblBranchTableAdapter
        '
        Me.TblBranchTableAdapter.ClearBeforeFill = True
        '
        'tblSubfeederBindingSource
        '
        Me.tblSubfeederBindingSource.DataMember = "tblSubfeeder"
        Me.tblSubfeederBindingSource.DataSource = Me.ESD_DatabaseDataSet
        '
        'TblSubfeederTableAdapter
        '
        Me.TblSubfeederTableAdapter.ClearBeforeFill = True
        '
        'tblDistributionBindingSource
        '
        Me.tblDistributionBindingSource.DataMember = "tblDistribution"
        Me.tblDistributionBindingSource.DataSource = Me.ESD_DatabaseDataSet
        '
        'TblDistributionTableAdapter
        '
        Me.TblDistributionTableAdapter.ClearBeforeFill = True
        '
        'tblMainFeederBindingSource
        '
        Me.tblMainFeederBindingSource.DataMember = "tblMainFeeder"
        Me.tblMainFeederBindingSource.DataSource = Me.ESD_DatabaseDataSet
        '
        'TblMainFeederTableAdapter
        '
        Me.TblMainFeederTableAdapter.ClearBeforeFill = True
        '
        'tblTransGenBindingSource
        '
        Me.tblTransGenBindingSource.DataMember = "tblTransGen"
        Me.tblTransGenBindingSource.DataSource = Me.ESD_DatabaseDataSet
        '
        'TblTransGenTableAdapter
        '
        Me.TblTransGenTableAdapter.ClearBeforeFill = True
        '
        'tblProjectBindingSource
        '
        Me.tblProjectBindingSource.DataMember = "tblProject"
        Me.tblProjectBindingSource.DataSource = Me.ESD_DatabaseDataSet
        '
        'TblProjectTableAdapter
        '
        Me.TblProjectTableAdapter.ClearBeforeFill = True
        '
        'FormReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1305, 784)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FormReport"
        Me.Text = "Report"
        CType(Me.ESD_DatabaseDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblBranchBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblSubfeederBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblDistributionBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblMainFeederBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblTransGenBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblProjectBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents reportProject1 As ReportProject
    Friend WithEvents ESD_DatabaseDataSet As ESD_DatabaseDataSet
    Friend WithEvents tblBranchBindingSource As BindingSource
    Friend WithEvents TblBranchTableAdapter As ESD_DatabaseDataSetTableAdapters.tblBranchTableAdapter
    Friend WithEvents tblSubfeederBindingSource As BindingSource
    Friend WithEvents TblSubfeederTableAdapter As ESD_DatabaseDataSetTableAdapters.tblSubfeederTableAdapter
    Friend WithEvents tblDistributionBindingSource As BindingSource
    Friend WithEvents TblDistributionTableAdapter As ESD_DatabaseDataSetTableAdapters.tblDistributionTableAdapter
    Friend WithEvents tblMainFeederBindingSource As BindingSource
    Friend WithEvents TblMainFeederTableAdapter As ESD_DatabaseDataSetTableAdapters.tblMainFeederTableAdapter
    Friend WithEvents tblTransGenBindingSource As BindingSource
    Friend WithEvents TblTransGenTableAdapter As ESD_DatabaseDataSetTableAdapters.tblTransGenTableAdapter
    Friend WithEvents tblProjectBindingSource As BindingSource
    Friend WithEvents TblProjectTableAdapter As ESD_DatabaseDataSetTableAdapters.tblProjectTableAdapter
End Class
