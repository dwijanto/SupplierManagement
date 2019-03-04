Imports System.Windows.Forms

Public Class DialogFinanceAssetNo
    Public Property FinanceAssetNo As String
    Dim errorProvider1 As New ErrorProvider

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        errorProvider1.SetError(TextBox1, "")
        If TextBox1.TextLength > 0 Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            FinanceAssetNo = TextBox1.Text
            Me.Close()
        Else
            errorProvider1.SetError(TextBox1, "Value cannot be blank.")
        End If

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
