Public Class UCNSGeneralContract
    Dim myDRV As DataRowView
    Dim bsPaymentTerm As BindingSource
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByVal drv As DataRowView, ByVal bsPaymentTerm As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        Me.myDRV = drv
        Me.bsPaymentTerm = bsPaymentTerm
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub InitData()
        ComboBox5.DataBindings.Clear()
        ComboBox5.DataSource = bsPaymentTerm
        ComboBox5.DisplayMember = "payt"
        ComboBox5.ValueMember = "paymenttermid"
        ComboBox5.DataBindings.Add("SelectedValue", myDRV, "paymentcode", True, DataSourceUpdateMode.OnPropertyChanged)
    End Sub
End Class
