Public Class DialogCMMFHelper
    Private Shared myInstance As DialogCMMFHelper

    Public Shared Function GetInstance() As DialogCMMFHelper
        If IsNothing(myInstance) Then
            myInstance = New DialogCMMFHelper
        End If
        Return myInstance
    End Function



    Private Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim bs As BindingSource = DataGridView1.DataSource
        Try
            bs.Filter = String.Format("[" & DataGridView1.Columns(0).DataPropertyName & "] like '*{0}*'", TextBox1.Text.Replace("'", "''"))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
End Class