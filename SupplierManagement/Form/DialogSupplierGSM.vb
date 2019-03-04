Imports System.Windows.Forms

Public Class DialogSupplierGSM
    Dim DRV As DataRowView
    Dim GSMBS As BindingSource
    Dim VendorBS As BindingSource
    Dim VendorHelperBS As BindingSource
    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView, ByVal VendorBS As BindingSource, ByVal VendorHelperBS As BindingSource, ByVal GSMBS As BindingSource)
        InitializeComponent()
        Me.DRV = drv
        Me.DRV.BeginEdit()
        Me.GSMBS = GSMBS
        Me.VendorBS = VendorBS
        Me.VendorHelperBS = VendorHelperBS


    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox
        Dim cbdrv As DataRowView = ComboBox1.SelectedItem
        DRV.Item("vendorcode") = cbdrv.Item("vendorcode")


        Dim cbdrv2 As DataRowView = ComboBox2.SelectedItem
        DRV.Item("gsmid") = cbdrv2.Item("gsmid")

        'Check Effectivedate
        If Not DateTimePicker1.Checked Then
            DRV.Item("effectivedate") = DBNull.Value

        End If
        
        Return True
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
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
        ComboBox1.DataSource = VendorBS
        ComboBox1.DisplayMember = "description"
        ComboBox1.ValueMember = "vendorcode"

        ComboBox2.DataSource = GSMBS
        ComboBox2.DisplayMember = "gsm"
        ComboBox2.ValueMember = "gsmid"

        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", DRV, "gsmid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker1.DataBindings.Add(New Binding("Text", DRV, "effectivedate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        If DRV.Row.RowState = DataRowState.Detached Then
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
        End If

    End Sub

    Private Sub validcombo1()
        Dim drv = ComboBox1.SelectedItem
        If Not IsNothing(drv) Then
            Me.DRV.Item("vendorname") = drv.item("vendorname")
            Me.DRV.Item("shortname") = drv.item("shortname")
            RaiseEvent RefreshDataGridView(Me, New EventArgs)            
        End If
    End Sub
    Private Sub validcombo2()
        Dim drv = ComboBox2.SelectedItem
        If Not IsNothing(drv) Then
            Me.DRV.Item("gsmid") = drv.item("gsmid")
            Me.DRV.Item("gsm") = drv.item("gsm")
            RaiseEvent RefreshDataGridView(Me, New EventArgs)
        End If
    End Sub
    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
        'validcombo1()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        'validcombo1()
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        validcombo1()
    End Sub
    Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectionChangeCommitted
        validcombo2()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim myobj As Button = CType(sender, Button)
        Select Case myobj.Name
            Case "Button1"
                Dim myform = New FormHelper(VendorHelperBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "description"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = VendorHelperBS.Current

                    Me.DRV.BeginEdit()
                    Me.DRV.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
                    Me.DRV.Row.Item("vendorname") = drv.Row.Item("vendorname")
                    Me.DRV.Row.Item("shortname") = drv.Row.Item("shortname")
                    'Need bellow code to sync with combobox
                    Dim myposition = VendorBS.Find("vendorcode", drv.Row.Item("vendorcode"))
                    VendorBS.Position = myposition

                End If

        End Select
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub
End Class
