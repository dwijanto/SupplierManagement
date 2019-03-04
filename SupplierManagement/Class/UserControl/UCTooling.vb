Public Class UCTooling

    Public Property ToolingHD As BindingSource
    Public Property ToolingDT As BindingSource

    Public Sub DisplayDataGrid()
        DataGridView1.AutoGenerateColumns = False
        DataGridView2.AutoGenerateColumns = False
        DataGridView1.DataSource = ToolingHD
        DataGridView2.DataSource = ToolingDT
        DataGridView1.Invalidate()
        DataGridView2.Invalidate()

    End Sub

End Class
