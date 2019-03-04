Public Class UCFactoryContact
    Implements IUTODetail

    Public VendorBS As BindingSource
    Public FactoryBS As BindingSource
    Public ContactBS As BindingSource
    Public SupplierTechnologyBS As BindingSource
    Public PanelHistoryBS As BindingSource

    Public Sub DisplayValue() Implements IUTODetail.DisplayValue
        DataGridView1.AutoGenerateColumns = False
        DataGridView2.AutoGenerateColumns = False
        DataGridView3.AutoGenerateColumns = False
        DataGridView4.AutoGenerateColumns = False
        DataGridView5.AutoGenerateColumns = False
        DataGridView1.DataSource = VendorBS
        DataGridView2.DataSource = FactoryBS
        DataGridView3.DataSource = ContactBS
        DataGridView4.DataSource = SupplierTechnologyBS
        DataGridView5.DataSource = PanelHistoryBS
        DataGridView3.DataSource = ContactBS
        TextBox19.DataBindings.Clear()
        TextBox19.DataBindings.Add(New Binding("text", FactoryBS, "customname", True))
    End Sub

    Public Sub initLabel() Implements IUTODetail.initLabel

    End Sub
End Class
