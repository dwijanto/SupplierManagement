Imports System.Windows.Forms

Public Class DialogVendorFamilyPM

    Dim DRV As DataRowView

    Dim FamilyBS As BindingSource
    Dim FamilyHelperBS As BindingSource

    Dim VendorBS As BindingSource
    Dim VendorHelperBS As BindingSource

    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView, ByVal VendorBS As BindingSource, ByVal VendorHelperBS As BindingSource, ByVal FamilyBS As BindingSource, ByVal FamilyHelperBS As BindingSource)
        InitializeComponent()
        Me.DRV = drv        
        Me.FamilyBS = FamilyBS
        Me.FamilyHelperBS = FamilyHelperBS
        Me.VendorBS = VendorBS
        Me.VendorHelperBS = VendorHelperBS

    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox
        'Dim cbdrv As DataRowView = ComboBox1.SelectedItem
        'DRV.Item("vendorname") = cbdrv.Item("vendorname").ToString.Trim
        'DRV.Item("shortname") = cbdrv.Item("shortname").ToString.Trim
        validcombo1()
        validcombo2()
        'Dim cbdrv2 As DataRowView = ComboBox2.SelectedItem
        'Me.DRV.Item("familyname") = cbdrv2.Item("familyname")
        'DRV.Item("pm") = cbdrv2.Item("pm")
        'DRV.Item("spm") = cbdrv2.Item("spm")
        If DateTimePicker1.Checked = False Then
            DRV.Item("pmeffectivedate") = DBNull.Value
        End If
        If DateTimePicker2.Checked = False Then
            DRV.Item("spmeffectivedate") = DBNull.Value
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
        ComboBox1.DisplayMember = "vendordesc"
        ComboBox1.ValueMember = "vendorcode"

        ComboBox2.DataSource = FamilyBS
        ComboBox2.DisplayMember = "familydesc2"
        ComboBox2.ValueMember = "familyid"

        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        DateTimePicker1.DataBindings.Clear()
        DateTimePicker2.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", DRV, "familyid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker1.DataBindings.Add(New Binding("Text", DRV, "pmeffectivedate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker2.DataBindings.Add(New Binding("Text", DRV, "spmeffectivedate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        If DRV.Row.RowState = DataRowState.Detached Then
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
        End If
        'If DRV.Item(0) < 0 Then
        '    ComboBox1.SelectedIndex = -1
        '    ComboBox2.SelectedIndex = -1
        'End If
    End Sub

    Private Sub validcombo1()
        Dim drv = ComboBox1.SelectedItem
        If Not IsNothing(drv) Then
            Me.DRV.Item("vendorname") = drv.item("vendorname").ToString.Trim
            Me.DRV.Item("shortname") = drv.item("shortname").ToString.Trim
            RaiseEvent RefreshDataGridView(Me, New EventArgs)
        End If
    End Sub
    Private Sub validcombo2()
        Dim drv = ComboBox2.SelectedItem
        If Not IsNothing(drv) Then
            'Me.DRV.Item("familyname") = drv.item("familyname")
            Me.DRV.Item("familydesc") = drv.item("familydesc")
            Me.DRV.Item("pm") = drv.item("pm")
            Me.DRV.Item("spm") = drv.item("spm")
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
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button2.Click
        Dim myobj As Button = CType(sender, Button)
        Select Case myobj.Name
            Case "Button1"
                Dim myform = New FormHelper(VendorHelperBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "vendordesc"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = VendorHelperBS.Current

                    'Me.DRV.BeginEdit()
                    Me.DRV.Row.Item("vendorname") = drv.Row.Item("vendorname").ToString.Trim
                    Me.DRV.Row.Item("shortname") = drv.Row.Item("shortname").ToString.Trim

                    'Need bellow code to sync with combobox
                    Dim myposition = VendorBS.Find("vendorcode", drv.Row.Item("vendorcode"))
                    VendorBS.Position = myposition

                End If
            Case "Button2"
                Dim myform = New FormHelper(FamilyHelperBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "familydesc"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = FamilyHelperBS.Current

                    'Me.DRV.BeginEdit()
                    Me.DRV.Row.Item("familyname") = drv.Row.Item("familyname")
                    Me.DRV.Row.Item("pm") = drv.Row.Item("pm")
                    Me.DRV.Row.Item("spm") = drv.Row.Item("spm")
                    'Need bellow code to sync with combobox
                    Dim myposition = FamilyBS.Find("familyid", drv.Row.Item("familyid"))
                    FamilyBS.Position = myposition

                End If
        End Select
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub

End Class
