Imports System.Windows.Forms
Imports SupplierManagement.PublicClass
Public Class DialogActionPlan

    Dim DRV As DataRowView
    Dim ShortNameBS As New BindingSource
    Dim ValidatorBS As New BindingSource
    Dim VendorModel1 As New VendorModel
    Dim OfficerSEBModel1 As New OfficerSEBModel

    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)
    Public PriorityList As New List(Of Priority)
    Dim TXMode As TXModeEnum
    Public Sub New(ByVal drv As DataRowView, ByVal TXMode As TXModeEnum)

        ' This call is required by the designer.
        InitializeComponent()
        Me.DRV = drv
        ' Add any initialization after the InitializeComponent() call.
        ShortNameBS = VendorModel1.GetShortNameBS
        ValidatorBS = OfficerSEBModel1.GetValidator
        Me.TXMode = TXMode
    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox
        Dim myret As Boolean = True
        DRV.Item("modifiedby") = HelperClass1.UserId
        DRV.Item("uploaddate") = Date.Today
        Dim item = ComboBox1.SelectedItem
        ErrorProvider1.SetError(ComboBox3, "")
        If item = "Closed" Then
            If ComboBox3.SelectedIndex <= 0 Then
                ErrorProvider1.SetError(ComboBox3, "Please select from the list.")
                myret = False
            End If
        End If
        'If DRV.Item("validatorflag") Then
        '    If DRV.Item("status") <> "Validated" Then
        '        DRV.Item("status") = "Validated"
        '        ErrorProvider1.SetError(ComboBox1, "Sorry, cannot change the status.")
        '        myret = False
        '    End If
        'End If

        'Dim cbdrv As DataRowView = ComboBox1.SelectedItem
        'If Not IsNothing(cbdrv) Then
        '    DRV.Item("brand") = cbdrv.Item("brand")
        'End If

        'ErrorProvider1.SetError(ComboBox4, "")
        'Dim cbdrv4 As DataRowView = ComboBox4.SelectedItem
        'If Not IsNothing(cbdrv4) Then
        '    DRV.Item("productlinegpsname") = cbdrv4.Item("productlinegpsname")
        '    DRV.Item("familylv2") = DRV.Item("familylv2")
        'Else
        '    myret = False
        '    ErrorProvider1.SetError(ComboBox4, "Product Line cannot be blank.")
        'End If


        'Dim cbdrv5 As DataRowView = ComboBox5.SelectedItem
        'If Not IsNothing(cbdrv5) Then
        '    DRV.Item("familyname") = cbdrv5.Item("familyname")
        '    DRV.Item("nsp1") = PriceDRV.Item("nsp1")
        '    DRV.Item("nsp2") = PriceDRV.Item("nsp2")
        'End If
        'ErrorProvider1.SetError(TextBox1, "")
        'If IsDBNull(DRV.Item("cmmf")) Then
        '    ErrorProvider1.SetError(TextBox1, "CMMF cannot be blank.")
        '    Return False
        'End If

        'PriceDRV.Item("cmmf") = DRV.Item("cmmf")


        Return myret
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim obj As Button = DirectCast(sender, Button)
        If obj.Text = "Validate" Then
            DRV.Item("Status") = "Validated"
            DRV.Item("validatorflag") = True
        End If
        If Me.Validate Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            DRV.EndEdit()

            RaiseEvent RefreshDataGridView(Me, e)
            Me.Close()
        Else
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        DRV.CancelEdit()

        RaiseEvent RefreshDataGridView(Me, e)
        Me.Close()
    End Sub
    Private Sub initData()
       
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox7.DataBindings.Clear()
        TextBox8.DataBindings.Clear()
        DateTimePicker1.DataBindings.Clear()
        DateTimePicker2.DataBindings.Clear()
        DateTimePicker3.DataBindings.Clear()

        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        ComboBox6.DataBindings.Clear()

        PriorityList = New List(Of Priority)
        PriorityList.Add(New Priority With {.priority = "H"})
        PriorityList.Add(New Priority With {.priority = "M"})
        PriorityList.Add(New Priority With {.priority = "L"})




        ComboBox2.DataSource = ShortNameBS
        ComboBox2.DisplayMember = "shortname"
        ComboBox2.ValueMember = "shortname"

        ComboBox3.DataSource = ValidatorBS
        ComboBox3.DisplayMember = "name"
        ComboBox3.ValueMember = "userid"

        ComboBox6.DataSource = PriorityList
        ComboBox6.DisplayMember = "priority"
        ComboBox6.ValueMember = "priority"

        TextBox1.DataBindings.Add(New Binding("Text", DRV, "target", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("Text", DRV, "result", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("Text", DRV, "responsibleperson", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox7.DataBindings.Add(New Binding("Text", DRV, "proposal", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox8.DataBindings.Add(New Binding("Text", DRV, "situation", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox1.DataBindings.Add(New Binding("Text", DRV, "status", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", DRV, "shortname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox3.DataBindings.Add(New Binding("SelectedValue", DRV, "validator", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox6.DataBindings.Add(New Binding("SelectedValue", DRV, "priority", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker1.DataBindings.Add(New Binding("Text", DRV, "startdate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker2.DataBindings.Add(New Binding("Text", DRV, "enddate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker3.DataBindings.Add(New Binding("Text", DRV, "finishdate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        If DRV.Row.RowState = DataRowState.Detached Then           
            ComboBox1.SelectedIndex = -1
            ComboBox6.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
        End If
        ComboBox6.Focus()
        OK_Button.Text = "OK"
        If TXMode = TXModeEnum.ValidateMode Then
            OK_Button.Text = "Validate"
        End If
        If ComboBox1.Text = "Validated" Then
            ComboBox1.Enabled = False
        End If
    End Sub

    'Private Sub validcombo4()
    '    Dim drv = ComboBox4.SelectedItem
    '    If Not IsNothing(drv) Then
    '        Me.DRV.Item("productlinegpsname") = drv.item("productlinegpsname")
    '        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    '    End If
    'End Sub
    'Private Sub validcombo5()
    '    Dim drv = ComboBox5.SelectedItem
    '    If Not IsNothing(drv) Then
    '        Me.DRV.Item("familyname") = drv.item("familyname")
    '        Me.DRV.Item("familylv2") = drv.item("familylv2")

    '        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    '    End If
    'End Sub



    'Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs)
    '    validcombo4()
    'End Sub
    'Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs)
    '    validcombo5()
    'End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim myobj As Button = CType(sender, Button)
        Select Case myobj.Name
            Case "Button1"
                'Dim myform = New FormHelper(FamilyHelperBS)
                'myform.DataGridView1.Columns(0).DataPropertyName = "familyname"
                'If myform.ShowDialog = DialogResult.OK Then
                '    Dim drv As DataRowView = FamilyHelperBS.Current

                '    Me.DRV.BeginEdit()
                '    Me.DRV.Row.Item("familyname") = drv.Row.Item("familyname")

                '    'Need bellow code to sync with combobox
                '    Dim myposition = FamilyBS.Find("familyid", drv.Row.Item("familyid"))
                '    FamilyBS.Position = myposition
                'End If

        End Select
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub

    Private Sub DialogActionPlan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        initData()
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'show helper

        Dim myform = New FormHelper(ShortNameBS)
        myform.DataGridView1.Columns(0).DataPropertyName = "shortname"
        If myform.ShowDialog = DialogResult.OK Then
            Dim drv As DataRowView = ShortNameBS.Current
            TextBox1.Text = drv.Item("shortname")
        End If        
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim myItem = ComboBox1.SelectedItem
        Label9.Visible = False
        ComboBox3.Visible = False
        If myItem = "Closed" Then
            ComboBox3.Visible = True
            Label9.Visible = True
            If TXMode = TXModeEnum.ValidateMode Then
                OK_Button.Text = "Validate"
            End If
        Else
            ComboBox3.SelectedIndex = -1
            If TXMode = TXModeEnum.ValidateMode Then
                OK_Button.Text = "Reject"
            End If
        End If


    End Sub
End Class

Public Class Priority
    Public Property priority As String

    Public Overrides Function ToString() As String
        Return priority.ToString
    End Function
End Class