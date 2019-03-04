Imports System.Windows.Forms

Public Class DialogAssetPurchaseComments
    Private drv As DataRowView
    Public Sub New(ByRef drv As DataRowView)
        Me.drv = drv
        ' This call is required by the designer.
        InitializeComponent()
        TextBox1.DataBindings.Add(New Binding("Text", drv, "comments", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Overloads Function Validate()
        Dim myret As Boolean = False
        If TextBox1.TextLength > 0 Then
            myret = True
        End If
        Return myret
    End Function
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Not Me.Validate Then
            If MessageBox.Show("Comment is blank. Do you want to continue?", "Blank Comment", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Cancel Then
                drv.CancelEdit()
                Exit Sub
            End If
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        drv.CancelEdit()
        Me.Close()
    End Sub

End Class
