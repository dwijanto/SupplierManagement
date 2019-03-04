Imports SupplierManagement.PublicClass

Public Class UCSustainableDevelopment
    Public Property CharterBS As BindingSource
    Public Property SocialAuditBS As BindingSource

    Public Sub DisplayDataGrid()
        DataGridView1.AutoGenerateColumns = False
        DataGridView2.AutoGenerateColumns = False
        DataGridView1.DataSource = CharterBS
        DataGridView2.DataSource = SocialAuditBS
        DataGridView1.Invalidate()
        DataGridView2.Invalidate()

    End Sub



    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        HelperClass1.previewdoc(CharterBS, HelperClass1.document)
    End Sub

    Private Sub DataGridView2_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        HelperClass1.previewdoc(SocialAuditBS, HelperClass1.document)
    End Sub


End Class
