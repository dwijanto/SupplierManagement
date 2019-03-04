Imports System.Windows.Forms

Public Class DialogAssetPurchaseFinance
    Private drv As DataRowView
    Public Sub New(ByRef drv As DataRowView)

        ' This call is required by the designer.
        InitializeComponent()
        Me.drv = drv
        TextBox1.DataBindings.Add(New Binding("Text", drv, "financeassetno", True, DataSourceUpdateMode.OnPropertyChanged, ""))
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
        'If Me.Validate Then
        '    Me.DialogResult = System.Windows.Forms.DialogResult.OK
        '    Me.Close()
        'End If
        If Not Me.Validate Then
            If MessageBox.Show("Finance Asset# is blank. Do you want to continue?", "Blank Finance Asset#", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Cancel Then
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
