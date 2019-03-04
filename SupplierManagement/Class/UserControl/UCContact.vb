Imports System.Text

Public Class UCContact
    Public Property bs As BindingSource

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Width = GroupBox1.Width + 3
        Me.Height = GroupBox1.Height + 3
    End Sub

    Public Sub DataBinding()

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        TextBox6.DataBindings.Clear()
        TextBox7.DataBindings.Clear()
        CheckBox1.DataBindings.Clear()

        TextBox1.DataBindings.Add(New Binding("Text", bs, "contactname", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox2.DataBindings.Add(New Binding("Text", bs, "title", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox3.DataBindings.Add(New Binding("Text", bs, "email", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox4.DataBindings.Add(New Binding("Text", bs, "officeph", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox5.DataBindings.Add(New Binding("Text", bs, "factoryph", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox6.DataBindings.Add(New Binding("Text", bs, "officemb", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox7.DataBindings.Add(New Binding("Text", bs, "factorymb", True, DataSourceUpdateMode.OnPropertyChanged))
        CheckBox1.DataBindings.Add(New Binding("checked", bs, "isecoqualitycontact", True, DataSourceUpdateMode.OnPropertyChanged))

    End Sub

    Public Function ValidateInput(ByVal errormessage As String) As Boolean
        Dim myret As Boolean = True
        Dim sb As StringBuilder = New StringBuilder
        If Not IsNothing(bs.Current) Then
            Dim drv As DataRowView = bs.Current
            If IsDBNull(drv.Row.Item("contactname")) Then
                myret = False
                sb.Append("Name cannot be blank.")
                ErrorProvider1.SetError(TextBox1, "Name cannot be blank.")
            Else
                If drv.Row.Item("contactname") = "" Then
                    myret = False
                    ErrorProvider1.SetError(TextBox1, "Name cannot be blank.")
                    sb.Append("Name cannot be blank.")
                End If
            End If
            If drv.Row.RowState <> DataRowState.Added Then 'Only modify updated one
                If myret Then 'No Error Found
                    If Not drv.Row.RowState = DataRowState.Modified Then drv.Row.SetModified()
                    bs.EndEdit()
                End If
            End If
        End If
        errormessage = sb.ToString
        Return myret
    End Function

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        ErrorProvider1.SetError(TextBox1, "")
    End Sub
End Class
