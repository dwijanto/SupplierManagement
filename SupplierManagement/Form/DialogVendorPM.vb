Imports System.Windows.Forms

Public Class DialogVendorPM

    Dim DRV As DataRowView

    Dim PMBS As BindingSource
    Dim PMBSHelperBS As BindingSource

    Dim VendorBS As BindingSource
    Dim VendorHelperBS As BindingSource

    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView, ByVal VendorBS As BindingSource, ByVal VendorHelperBS As BindingSource, ByVal PMBS As BindingSource, ByVal PMBSHelperBS As BindingSource)
        InitializeComponent()
        Me.DRV = drv
        Me.PMBS = PMBS
        Me.PMBSHelperBS = PMBSHelperBS
        Me.VendorBS = VendorBS
        Me.VendorHelperBS = VendorHelperBS

    End Sub

    Public Overloads Function Validate() As Boolean
        validcombo1()
        validcombo2()
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

        ComboBox2.DataSource = PMBS
        ComboBox2.DisplayMember = "pm"
        ComboBox2.ValueMember = "pmid"

        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        DateTimePicker1.DataBindings.Clear()
        DateTimePicker2.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", DRV, "pmid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
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
                Dim myform = New FormHelper(PMBSHelperBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "pm"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = PMBSHelperBS.Current

                    'Me.DRV.BeginEdit()
                    'Me.DRV.Row.Item("familyname") = drv.Row.Item("familyname")
                    Me.DRV.Row.Item("pm") = drv.Row.Item("pm")
                    Me.DRV.Row.Item("spm") = drv.Row.Item("spm")
                    'Need bellow code to sync with combobox
                    Dim myposition = PMBS.Find("pmid", drv.Row.Item("pmid"))
                    PMBS.Position = myposition

                End If
        End Select
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub

End Class
