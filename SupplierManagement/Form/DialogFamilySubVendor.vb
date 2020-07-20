Imports System.Windows.Forms

Public Class DialogFamilySubVendor
    Dim FamilyBS As BindingSource
    Dim SubFamilyBS As BindingSource
    Public Sub New(ByRef FamilyBS As BindingSource, ByRef SubFamilyBS As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.FamilyBS = FamilyBS
        Me.SubFamilyBS = SubFamilyBS

        bindingControls()
    End Sub


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub bindingControls()
        ComboBox1.DataSource = FamilyBS
        ComboBox1.DisplayMember = "familyvcdesc"
        ComboBox1.ValueMember = "familyvc"

        ComboBox2.DataSource = SubFamilyBS
        ComboBox2.DisplayMember = "subfamilyvcdesc"
        ComboBox2.ValueMember = "subfamilyvc"
    End Sub



    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim mydr = ComboBox1.SelectedItem
        SubFamilyBS.Filter = String.Format("familyvc='{0}'", mydr.Row.Item("familyvc"))
    End Sub
End Class
