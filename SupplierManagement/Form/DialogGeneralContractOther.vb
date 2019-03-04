Imports System.Windows.Forms

Public Class DialogGeneralContractOther
    Private DRV As DataRowView
    Private PaymentTermBS As BindingSource
    Public Sub New(ByRef DRV As DataRowView, ByVal PaymentTermBS As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.DRV = DRV
        Dim dt As DataTable = DirectCast(PaymentTermBS.DataSource, DataTable).Copy
        Me.PaymentTermBS = New BindingSource
        Me.PaymentTermBS.DataSource = dt
        initialDataRow()
    End Sub
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.validate Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()       
        End If

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub initialDataRow()
        ComboBox1.DataSource = PaymentTermBS
        ComboBox1.DisplayMember = "payt"
        ComboBox1.ValueMember = "paymenttermid"

        ComboBox1.DataBindings.Add(New Binding("selectedvalue", DRV, "paymentcode", False, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker1.DataBindings.Add(New Binding("Text", DRV, "otherdate", False, DataSourceUpdateMode.OnPropertyChanged))
        CheckBox1.DataBindings.Add(New Binding("checked", DRV, "status", False, DataSourceUpdateMode.OnPropertyChanged))
    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        Dim cbdrv = ComboBox1.SelectedItem
        ErrorProvider1.SetError(ComboBox1, "")
        ErrorProvider1.SetError(DateTimePicker1, "")
        If IsNothing(cbdrv) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from the list.")
            myret = False
        Else
            DRV.Item("payt") = cbdrv.item("payt")
        End If
        'If DateTimePicker1.Value.Date = Today.Date Then
        '    ErrorProvider1.SetError(DateTimePicker1, "Please change your date!")
        '    myret = False
        'End If

        Return myret
    End Function



End Class
