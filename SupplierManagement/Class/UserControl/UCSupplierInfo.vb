Public Class UCSupplierInfo
    Public Shortname As String
    Public vendorcode As Long
    Public Property BS As BindingSource

    Public Sub RefreshataGrid()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = BS
        DataGridView1.Invalidate()
    End Sub

End Class
