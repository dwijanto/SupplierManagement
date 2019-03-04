Imports System.Windows.Forms

Public Class DialogSupplierPhotos
    Private drv As DataRowView
    Private combodict As New Dictionary(Of Integer, String)
    Public Shared Event FinishTx()

    Public Sub New(ByVal drv As DataRowView)
        InitializeComponent()
        Me.drv = drv
        combodict.Add(1, "Factory")
        combodict.Add(2, "Product")
        ComboBox1.DisplayMember = "Value"
        ComboBox1.ValueMember = "key"
        ComboBox1.DataSource = New BindingSource(combodict, Nothing)
        initData()
        

    End Sub


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.validate Then
            drv.EndEdit()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
            RaiseEvent FinishTx()
        End If

       
    End Sub
    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True

        ErrorProvider1.SetError(TextBox1, "")
        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Description cannot be blank.")
            myret = False
        End If


        ErrorProvider1.SetError(ComboBox1, "")
        If IsNothing(ComboBox1.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from list.")
            myret = False
        End If

        'if new record, filename cannot be blank
        ErrorProvider1.SetError(TextBox3, "")
        If drv.Row.Item("id") <= 0 Then
            If TextBox3.Text = "" Then
                ErrorProvider1.SetError(TextBox3, "Filename cannot be blank.")
                myret = False
            Else
                drv.Row.Item("filename") = IO.Path.GetFileName(TextBox3.Text)
            End If
        End If
        If TextBox3.Text <> "" Then
            drv.Row.Item("filename") = IO.Path.GetFileName(TextBox3.Text)
        End If
        Dim cbdr As System.Collections.Generic.KeyValuePair(Of Integer, String) = ComboBox1.SelectedItem
        drv.Row.Item("phototypename") = cbdr.Value
        Return myret
    End Function

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        drv.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub


    Private Sub initData()
        TextBox1.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        TextBox6.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()
        DateTimePicker1.DataBindings.Clear()
        CheckBox1.DataBindings.Clear()
        CheckBox2.DataBindings.Clear()
        CheckBox3.DataBindings.Clear()


        TextBox1.DataBindings.Add(New Binding("Text", drv, "description", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("Text", drv, "fullpath", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox4.DataBindings.Add(New Binding("Text", drv, "lineorderfp", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox5.DataBindings.Add(New Binding("Text", drv, "lineordercp", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox6.DataBindings.Add(New Binding("Text", drv, "lineordergeneral", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", drv, "phototype", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker1.DataBindings.Add(New Binding("text", drv, "createddate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        CheckBox1.DataBindings.Add(New Binding("checked", drv, "fp", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        CheckBox2.DataBindings.Add(New Binding("checked", drv, "cp", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        CheckBox3.DataBindings.Add(New Binding("checked", drv, "general", True, DataSourceUpdateMode.OnPropertyChanged, ""))


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim myfolder As New OpenFileDialog
        If myfolder.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TextBox3.Text = myfolder.FileName

        End If
    End Sub
End Class

