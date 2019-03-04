Public Class FormPriceHistory
    Dim model As ModelPriceHistory
    Dim vendorcode As String
    Dim shortname As String
    Dim drv As DataRowView


    Private Sub FormPriceHistory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        model.loadData()
    End Sub


    Public Sub New(ByVal bs As BindingSource, ByVal vendorquery As FormSupplierDashboard.VendorQuery, ByVal vendorcode As Object, ByVal shortname As Object)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        model = New ModelPriceHistory(Me, bs, vendorquery, vendorcode, shortname)
        drv = bs.Current
        model.f = AddressOf callback

    End Sub

    Private Sub callback(ByRef obj As Object, ByRef e As EventArgs)
        Dim dt As DataTable = DirectCast(obj, DataTable)
        Dim bs As New BindingSource
        bs.DataSource = dt
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = bs
        TextBox1.Text = drv.Row.Item("cmmf")
        TextBox2.Text = "" & drv.Row.Item("commercialref")
        TextBox3.Text = "" & drv.Row.Item("materialdesc")
    End Sub



   
End Class
