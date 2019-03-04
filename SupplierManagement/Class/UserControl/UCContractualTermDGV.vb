Imports SupplierManagement.PublicClass

Public Class UCContractualTermDGV
    Implements IUTODetail

    Dim _sapbs As BindingSource
    Dim _contractBS As BindingSource
    Dim _authLetter As BindingSource
    Dim _projectSpec As BindingSource
    Dim _citiprogram As BindingSource
    Dim _vendorcurrency As BindingSource

    Public Property VendorPaymentDate As Date
    Public Property VFPDate As Date
    Public Event ShowPopUp()

    Public Property SAPBS As BindingSource
        Get
            Return _sapbs
        End Get
        Set(ByVal value As BindingSource)
            _sapbs = value
        End Set
    End Property
    Public Property ContractBS As BindingSource
        Get
            Return _contractBS
        End Get
        Set(ByVal value As BindingSource)
            _contractBS = value
        End Set
    End Property
    Public Property AuthLeter As BindingSource
        Get
            Return _authLetter
        End Get
        Set(ByVal value As BindingSource)
            _authLetter = value
        End Set
    End Property

    Public Property ProjectSpec As BindingSource
        Get
            Return _projectSpec
        End Get
        Set(ByVal value As BindingSource)
            _projectSpec = value
        End Set
    End Property

    Public Property CitiProgram As BindingSource
        Get
            Return _citiprogram
        End Get
        Set(ByVal value As BindingSource)
            _citiprogram = value
        End Set
    End Property

    Public Property VendorCurrency As BindingSource
        Get
            Return _vendorcurrency
        End Get
        Set(ByVal value As BindingSource)
            _vendorcurrency = value
        End Set
    End Property


    Public Sub DisplayLatestUpdate()
        Label10.Text = String.Format("Latest Import date of SAP Vendor Payment term : {0:dd-MMM-yyyy}", VendorPaymentDate)
        Label12.Text = String.Format("Latest Import date of VFP Program Status : {0:dd-MMM-yyyy}", VFPDate)
    End Sub

    Public Sub DisplayValue() Implements IUTODetail.DisplayValue
        ClearValue()
        ListBox3.DataSource = _sapbs
        ListBox3.DisplayMember = "sappaymentterm"
        'ListBox1.DataSource = _authLetter
        'ListBox1.DisplayMember = "description"
        'ListBox2.DataSource = _projectSpec
        'ListBox2.DisplayMember = "description"
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = _authLetter
        DataGridView2.AutoGenerateColumns = False
        DataGridView2.DataSource = _projectSpec
        DataGridView3.AutoGenerateColumns = False
        DataGridView3.DataSource = _citiprogram
        DataGridView4.AutoGenerateColumns = False
        DataGridView4.DataSource = _vendorcurrency
        For Each drv As DataRowView In _contractBS.List
            Select Case drv.Item("doctypeid")
                Case 32
                    TextBox2.Text = "" & drv.Item("payt") & " " & drv.Item("details")
                    TextBox6.Text = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("docdate"))
                    TextBox8.Text = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("expireddate"))
                Case 33
                    TextBox4.Text = "" & drv.Item("nqsu")
                    TextBox5.Text = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("remarks"))
                    TextBox11.Text = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("docdate"))
                    TextBox10.Text = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("expireddate"))

                Case 35
                    TextBox15.Text = "" & drv.Item("sasl") & " %"
                    TextBox19.Text = "" & drv.Item("leadtime")
                    TextBox14.Text = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("remarks"))
                    TextBox13.Text = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("docdate"))
                    TextBox12.Text = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("expireddate"))
            End Select
        Next
    End Sub

    Public Sub initLabel() Implements IUTODetail.initLabel

    End Sub

    Public Sub ClearValue()
        For Each obj As Control In Me.Controls
            If obj.GetType() Is GetType(TextBox) Then
                Dim Txt As TextBox = CType(obj, TextBox)
                Txt.Text = ""
            End If
        Next
    End Sub





    Private Sub DataGridView2_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        HelperClass1.previewdoc(ProjectSpec, HelperClass1.document)
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        HelperClass1.previewdoc(AuthLeter, HelperClass1.document)
    End Sub


    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        RaiseEvent ShowPopUp()
    End Sub
End Class
