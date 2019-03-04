Public Class FormDialogDeleteFactoryContact
    Public Property BS As BindingSource
    Public Property MstBS As BindingSource

    Public Sub DataBinding()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = BS
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)




    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub DeleteAssignmentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteAssignmentToolStripMenuItem.Click
        If Not IsNothing(BS.Current) Then
            If MessageBox.Show("Delete selected record?", "Delete Record(s)", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    BS.RemoveAt(drv.Index)
                Next
                If BS.Count = 0 Then
                    'Remove MstBS
                    'MstBS.RemoveCurrent()
                End If
            End If
        End If
    End Sub
End Class