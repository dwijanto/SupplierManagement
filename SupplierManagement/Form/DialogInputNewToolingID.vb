Imports System.Windows.Forms

Public Class DialogInputNewToolingID
    Public Property Result As String
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.validate Then

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Result = TextBox1.Text
            Me.Close()
        End If
        
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(TextBox1, "")
        If TextBox1.Text.Length = 0 Then
            ErrorProvider1.SetError(TextBox1, "Value cannot be blank")
            myret = False
        End If
        Return myret
    End Function
End Class
