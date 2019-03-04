Imports System.Windows.Forms

Public Class DialogFamilyPM

    Dim DRV As DataRowView
    Dim PMBS As BindingSource
    Dim FamilyBS As BindingSource
    Dim FamilyHelperBS As BindingSource
    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView, ByVal FamilyBS As BindingSource, ByVal FamilyHelperBS As BindingSource, ByVal PMBS As BindingSource)
        InitializeComponent()
        Me.DRV = drv
        Me.DRV.BeginEdit()
        Me.PMBS = PMBS
        Me.FamilyBS = FamilyBS
        Me.FamilyHelperBS = FamilyHelperBS
    End Sub

    Public Overloads Function Validate() As Boolean
        'Check combobox
        Dim cbdrv As DataRowView = ComboBox1.SelectedItem
        DRV.Item("familydesc") = cbdrv.Item("familydesc")


        Dim cbdrv2 As DataRowView = ComboBox2.SelectedItem
        'DRV.Item("username") = cbdrv2.Item("pm")
        DRV.Item("pm") = cbdrv2.Item("pm")


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
        ComboBox1.DataSource = FamilyBS
        ComboBox1.DisplayMember = "familydesc"
        ComboBox1.ValueMember = "familyid"

        ComboBox2.DataSource = PMBS
        ComboBox2.DisplayMember = "PM"
        ComboBox2.ValueMember = "PMID"

        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "familyid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", DRV, "pmid", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        If DRV.Row.RowState = DataRowState.Detached Then
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
        End If
        'If IsDBNull(DRV.Item(0)) Then
        ' ComboBox1.SelectedIndex = -1
        ' ComboBox2.SelectedIndex = -1
        ' End If
    End Sub

    Private Sub validcombo1()
        Dim drv = ComboBox1.SelectedItem
        If Not IsNothing(drv) Then
            Me.DRV.Item("familydesc") = drv.item("familydesc")
            RaiseEvent RefreshDataGridView(Me, New EventArgs)
        End If
    End Sub
    Private Sub validcombo2()
        Dim drv = ComboBox2.SelectedItem
        If Not IsNothing(drv) Then            
            'Me.DRV.Item("username") = drv.item("pm")
            Me.DRV.Item("pm") = drv.item("pm")
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
                Dim myform = New FormHelper(FamilyHelperBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "familydesc"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = FamilyHelperBS.Current

                    Me.DRV.BeginEdit()
                    Me.DRV.Row.Item("familydesc") = drv.Row.Item("familydesc")
                    
                    'Need bellow code to sync with combobox
                    Dim myposition = FamilyBS.Find("familyid", drv.Row.Item("familyid"))
                    FamilyBS.Position = myposition

                End If

        End Select
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub
End Class
