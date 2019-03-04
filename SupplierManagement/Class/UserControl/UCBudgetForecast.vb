Public Class UCBudgetForecast
    Public Property title As String

    Public Sub init()
        Label2.Text = title
    End Sub

    Private Sub UCBudgetForecast_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
