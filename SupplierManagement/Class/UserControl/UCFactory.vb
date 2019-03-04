Imports System.Text

Public Class UCFactory

    Public Property bs As BindingSource       
    Public Property ProvinceBS As BindingSource
    Public Property CountryBS As BindingSource

  

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Width = GroupBox1.Width + 3
        Me.Height = GroupBox1.Height + 3

    End Sub
    Public Sub DataBinding()

        TextBox1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()

        ComboBox1.DataSource = ProvinceBS
        ComboBox1.DisplayMember = "province"
        ComboBox1.ValueMember = "id"

        ComboBox2.DataSource = CountryBS
        ComboBox2.DisplayMember = "country"
        ComboBox2.ValueMember = "id"

        TextBox1.DataBindings.Add(New Binding("Text", bs, "chinesename", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox2.DataBindings.Add(New Binding("Text", bs, "englishname", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox3.DataBindings.Add(New Binding("Text", bs, "chineseaddress", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox4.DataBindings.Add(New Binding("Text", bs, "englishaddress", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox5.DataBindings.Add(New Binding("Text", bs, "area", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox6.DataBindings.Add(New Binding("Text", bs, "city", True, DataSourceUpdateMode.OnPropertyChanged))

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", bs, "provinceid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", bs, "countryid", True, DataSourceUpdateMode.OnPropertyChanged))
    End Sub


    Public Function ValidateInput(ByRef errormessage As String) As Boolean
        Dim myret As Boolean = True
        Dim sb As StringBuilder = New StringBuilder
        If Not IsNothing(bs.Current) Then
            Dim drv As DataRowView = bs.Current
            If IsDBNull(drv.Row.Item("englishname")) Or drv.Row.Item("englishname") = "" Then 'Put Initial value when created row
                myret = False
                sb.Append("Factory Name (English) cannot be blank.")
                ErrorProvider1.SetError(TextBox2, "Factory Name (English) cannot be blank.")
            End If
            If IsDBNull(drv.Row.Item("englishaddress")) Or drv.Row.Item("englishaddress") = "" Then
                myret = False
                sb.Append("Address Name (English) cannot be blank.")
                ErrorProvider1.SetError(TextBox4, "Address Name (English) cannot be blank.")           
            End If
            If drv.Row.RowState <> DataRowState.Added Then 'Only modify updated one
                If myret Then 'No Error Found
                    If Not drv.Row.RowState = DataRowState.Modified Then drv.Row.SetModified()
                    'bs.Item("status")
                    bs.EndEdit()
                End If
            End If
        End If
        errormessage = sb.ToString
        Return myret
    End Function



    Private Sub TextBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        ErrorProvider1.SetError(TextBox2, "")
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        ErrorProvider1.SetError(TextBox4, "")
    End Sub
End Class
