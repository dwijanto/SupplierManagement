Imports System.Windows.Forms

Public Class DialogVendorFamilyPMEx

    Dim DRV As DataRowView

    Dim FamilyBS As BindingSource
    Dim FamilyHelperBS As BindingSource
    Dim VendorBS As BindingSource
    Dim VendorHelperBS As BindingSource
    Dim PMBS As BindingSource
    Dim PMBSHelper As BindingSource

    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView, ByVal VendorBS As BindingSource, ByVal VendorHelperBS As BindingSource, ByVal FamilyBS As BindingSource, ByVal FamilyHelperBS As BindingSource, ByVal PMBS As BindingSource, ByVal PMBSHelper As BindingSource)
        InitializeComponent()
        Me.DRV = drv
        Me.FamilyBS = FamilyBS
        Me.FamilyHelperBS = FamilyHelperBS
        Me.VendorBS = VendorBS
        Me.VendorHelperBS = VendorHelperBS
        Me.PMBS = PMBS
        Me.PMBSHelper = PMBSHelper

    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox
        validcombo1()
        'validcombo2()
        validcombo3()

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
        ComboBox1.ValueMember = "id"

        'ComboBox2.DataSource = FamilyBS
        'ComboBox2.DisplayMember = "familydesc2"
        'ComboBox2.ValueMember = "familyid"

        ComboBox3.DataSource = PMBS
        ComboBox3.DisplayMember = "pm"
        ComboBox3.ValueMember = "pmid"

        ComboBox1.DataBindings.Clear()
        'ComboBox2.DataBindings.Clear()
        ComboBox3.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "id", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        'ComboBox2.DataBindings.Add(New Binding("SelectedValue", DRV, "familyid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox3.DataBindings.Add(New Binding("SelectedValue", DRV, "pmid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        If DRV.Row.RowState = DataRowState.Detached Then
            ComboBox1.SelectedIndex = -1
            'ComboBox2.SelectedIndex = -1
            ComboBox3.SelectedIndex = -1
        End If       
    End Sub

    Private Sub validcombo1()
        Dim drv = ComboBox1.SelectedItem
        If Not IsNothing(drv) Then
            Me.DRV.Item("id") = drv.item("id")
            Me.DRV.Item("vendorcode") = drv.item("vendorcode")
            Me.DRV.Item("vendorname") = drv.item("vendorname").ToString.Trim
            Me.DRV.Item("shortname") = drv.item("shortname").ToString.Trim
            Me.DRV.Item("familyid") = drv.item("familyid")
            Me.DRV.Item("familydesc") = drv.item("familydesc")
            RaiseEvent RefreshDataGridView(Me, New EventArgs)
        End If
    End Sub
    'Private Sub validcombo2()
    '    Dim drv = ComboBox2.SelectedItem
    '    If Not IsNothing(drv) Then            
    '        Me.DRV.Item("familydesc") = drv.item("familydesc")
    '        Me.DRV.Item("pm") = drv.item("pm")
    '        'Me.DRV.Item("spm") = drv.item("spm")
    '        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    '    End If
    'End Sub

    Private Sub validcombo3()
        Dim drv = ComboBox3.SelectedItem
        If Not IsNothing(drv) Then
            Me.DRV.Item("pm") = drv.item("pm")
            'Me.DRV.Item("spm") = drv.item("spm")
            RaiseEvent RefreshDataGridView(Me, New EventArgs)
        End If
    End Sub

    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()       
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        'validcombo1()
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        validcombo1()
    End Sub
    'Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs)
    '    validcombo2()
    'End Sub
    Private Sub ComboBox3_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectionChangeCommitted
        validcombo3()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button3.Click
        Dim myobj As Button = CType(sender, Button)
        Select Case myobj.Name
            Case "Button1"
                Dim myform = New FormHelper(VendorHelperBS)
                myform.Size = New Size(800, 469)
                myform.DataGridView1.Columns(0).DataPropertyName = "description"
                myform.DataGridView1.Columns(0).Width = 700

                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = VendorHelperBS.Current

                    'Me.DRV.BeginEdit()
                    Me.DRV.Item("id") = drv.Item("id")
                    Me.DRV.Item("vendorcode") = drv.Item("vendorcode")
                    Me.DRV.Row.Item("vendorname") = drv.Row.Item("vendorname").ToString.Trim
                    Me.DRV.Row.Item("shortname") = drv.Row.Item("shortname").ToString.Trim
                    Me.DRV.Row.Item("familyid") = drv.Row.Item("familyid")
                    Me.DRV.Row.Item("familydesc") = drv.Row.Item("familydesc").ToString.Trim
                    'Need bellow code to sync with combobox
                    Dim myposition = VendorBS.Find("id", drv.Row.Item("id"))
                    VendorBS.Position = myposition

                End If
            Case "Button2"
                Dim myform = New FormHelper(FamilyHelperBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "familydesc"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = FamilyHelperBS.Current

                    'Me.DRV.BeginEdit()
                    Me.DRV.Row.Item("familyname") = drv.Row.Item("familyname")
                    'Me.DRV.Row.Item("pm") = drv.Row.Item("pm")
                    'Me.DRV.Row.Item("spm") = drv.Row.Item("spm")
                    'Need bellow code to sync with combobox
                    Dim myposition = FamilyBS.Find("familyid", drv.Row.Item("familyid"))
                    FamilyBS.Position = myposition

                End If
            Case "Button3"
                Dim myform = New FormHelper(PMBSHelper)
                myform.DataGridView1.Columns(0).DataPropertyName = "pm"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = PMBSHelper.Current

                    'Me.DRV.BeginEdit()
                    'Me.DRV.Row.Item("familyname") = drv.Row.Item("familyname")
                    Me.DRV.Row.Item("pm") = drv.Row.Item("pm")
                    'Me.DRV.Row.Item("spm") = drv.Row.Item("spm")
                    'Need bellow code to sync with combobox
                    Dim myposition = PMBS.Find("pmid", drv.Row.Item("pmid"))
                    PMBS.Position = myposition

                End If
        End Select
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub


End Class
