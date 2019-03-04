Imports System.Windows.Forms

Public Class DialogReguiredDocument
    Private BS As BindingSource
    Private CBBS As BindingSource
    Public Sub New(ByVal bs As BindingSource, ByVal CBBS As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        Me.BS = bs
        Me.CBBS = CBBS
        ' Add any initialization after the InitializeComponent() call.
        InitData()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        BS.EndEdit()
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        BS.CancelEdit()
        Me.Close()
    End Sub

    Private Sub InitData()
        Dim cb = DirectCast(DataGridView1.Columns(0), DataGridViewComboBoxColumn)
        cb.DisplayMember = "cvalue"
        cb.ValueMember = "paramdtid"
        cb.DataPropertyName = "doctypeid"
        cb.DataSource = CBBS

        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = BS

       

    End Sub

    Private Sub AddRecodToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddRecodToolStripMenuItem.Click
        Dim drv = BS.AddNew()
        If drv.Row.RowState = DataRowState.Detached Then
            Dim cb = DirectCast(DataGridView1.Columns(0), DataGridViewComboBoxColumn)
            'cb.SelectedIndex = -1            
            'cb.DisplayIndex = 1
        End If
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If Not IsNothing(BS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    BS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Me.Validate()
    End Sub

    Private Sub DialogReguiredDocument_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        Me.Validate()
    End Sub

End Class
