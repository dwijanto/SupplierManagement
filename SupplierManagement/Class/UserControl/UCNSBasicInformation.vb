Public Class UCNSBasicInformation
    Dim myDRV As DataRowView
    Dim bsDoctype As BindingSource
    Dim bsDoctypehelper As BindingSource
    Dim bsDocLevel As BindingSource
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub BindingControl(ByVal drv As DataRowView, ByVal bsDocType As BindingSource, ByVal bsDocTypeHelper As BindingSource, ByVal bsDocLevel As BindingSource)
        'InitializeComponent()
        Me.myDRV = drv
        Me.bsDoctype = bsDocType
        Me.bsDocLevel = bsDocLevel
        InitData()
    End Sub

    Private Sub InitData()
        ComboBox3.DataBindings.Clear()
        ComboBox4.DataBindings.Clear()

        DateTimePicker2.DataBindings.Clear()
        DateTimePicker3.DataBindings.Clear()

        TextBox8.DataBindings.Clear()
        TextBox9.DataBindings.Clear()
        TextBox10.DataBindings.Clear()

        ComboBox3.DataSource = bsDoctype
        ComboBox3.DisplayMember = "doctypename"
        ComboBox3.ValueMember = "id"

        ComboBox4.DataSource = bsDocLevel
        ComboBox4.DisplayMember = "levelname"
        ComboBox4.ValueMember = "id"

        ComboBox3.DataBindings.Add("SelectedValue", myDRV, "doctypeid", True, DataSourceUpdateMode.OnPropertyChanged)
        ComboBox4.DataBindings.Add("SelectedValue", myDRV, "doclevelid", True, DataSourceUpdateMode.OnPropertyChanged)
        DateTimePicker2.DataBindings.Add(New Binding("Text", myDRV, "docdate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker3.DataBindings.Add(New Binding("Text", myDRV, "expireddate", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        TextBox8.DataBindings.Add(New Binding("Text", myDRV, "remarks", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox9.DataBindings.Add(New Binding("Text", myDRV, "version", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox10.DataBindings.Add(New Binding("Text", myDRV, "filename", True, DataSourceUpdateMode.OnPropertyChanged, "")) 'if no value means update
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim myobj As Button = CType(sender, Button)
        Try
            Select Case myobj.Name

                Case "Button9"
                    Dim myform = New FormHelper(bsDocTypeHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "doctypename"
                    If Not myform.ShowDialog = DialogResult.OK Then                      
                        'checkcombobox3()
                        Dim drv As DataRowView = bsDocTypeHelper.Current                       
                        mydrv.Row.Item("doctypename") = drv.Row.Item("doctypename")
                        mydrv.Row.Item("doctypeid") = drv.Row.Item("id")

                        'Sync with bsDocType
                        Dim itemfound As Integer = bsDoctype.Find("id", drv.Row.Item("id"))
                        bsDoctype.Position = itemfound
                    End If
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim openfiledialog1 As New OpenFileDialog
        If openfiledialog1.ShowDialog = DialogResult.OK Then            
            mydrv.Row.Item("docname") = IO.Path.GetFileName(openfiledialog1.FileName)
            mydrv.Row.Item("docext") = IO.Path.GetExtension(openfiledialog1.FileName)            
            TextBox10.Text = openfiledialog1.FileName
        End If
    End Sub
End Class
