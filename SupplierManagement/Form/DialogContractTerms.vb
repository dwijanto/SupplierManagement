Imports System.Windows.Forms

Public Class DialogContractTerms
    Private _BS As New BindingSource

    Private bs1 As New BindingSource
    Private bs2 As New BindingSource
    Private bs3 As New BindingSource
    Private GeneralDT As DataTable
    Private SupplyChainDT As DataTable
    Private QualityDT As DataTable

    Const CONTRACT_GENERAL_CONTRACT As Integer = 32
    Const CONTRACT_QUALITY_APPENDIX As Integer = 33
    Const CONTRACT_SUPPLY_CHAIN_APPENDIX As Integer = 35

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Sub New(ByRef ContractBS As BindingSource, ByRef BS1 As BindingSource, ByRef BS2 As BindingSource, ByRef BS3 As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        Dim dt As DataTable = DirectCast(ContractBS.DataSource, DataTable).Copy
        Dim MyView As DataView = dt.AsDataView
        MyView.RowFilter = "doctypeid = 32"
        GeneralDT = MyView.ToTable("General")
        MyView.RowFilter = "doctypeid = 33"
        QualityDT = MyView.ToTable("Quality")
        MyView.RowFilter = "doctypeid = 35"
        SupplyChainDT = MyView.ToTable("SupplyChain")

        Me._BS.DataSource = ContractBS
        Me.bs1.DataSource = BS1
        Me.bs2.DataSource = BS2

        Me.bs3.DataSource = BS3
        ' Add any initialization after the InitializeComponent() call.
        InitData()
    End Sub

    Private Sub InitData()
        DataGridView4.AutoGenerateColumns = False
        DataGridView5.AutoGenerateColumns = False
        DataGridView6.AutoGenerateColumns = False
        DataGridView4.DataSource = bs1
        DataGridView5.DataSource = bs3
        DataGridView6.DataSource = bs2
        If GeneralDT.Rows.Count > 0 Then
            Dim drv = GeneralDT.Rows(0)
            TextBox1.Text = "" & drv.Item("payt") & " " & drv.Item("details")
        End If

        If SupplyChainDT.Rows.Count > 0 Then
            Dim drv = SupplyChainDT.Rows(0)
            TextBox12.Text = "" & drv.Item("sasl") & " %"
            TextBox11.Text = "" & drv.Item("leadtime")
        End If

        
        If QualityDT.Rows.Count > 0 Then
            Dim drv = QualityDT.Rows(0)
            TextBox13.Text = "" & drv.Item("nqsu")
        End If
        

    End Sub



End Class
