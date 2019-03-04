Imports System.Windows.Forms

Public Class DialogPaymentTermDate

    Dim DRV As DataRowView

    Dim PaymentTermBS As BindingSource
    Dim PaymentTermBSHelper As BindingSource
    Dim VendorBS As BindingSource
    Dim VendorHelperBS As BindingSource


    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView, ByVal VendorBS As BindingSource, ByVal VendorHelperBS As BindingSource, ByVal PaymentTermBS As BindingSource, ByVal PaymentTermBSHelper As BindingSource)
        InitializeComponent()
        Me.DRV = drv
        Me.PaymentTermBS = PaymentTermBS
        Me.PaymentTermBSHelper = PaymentTermBSHelper
        Me.VendorBS = VendorBS
        Me.VendorHelperBS = VendorHelperBS
  
    End Sub

    Public Overloads Function Validate() As Boolean       
        Return validcombo1() And validcombo2()
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.Validate Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            DRV.EndEdit()
            RaiseEvent RefreshDataGridView(Me, e)
            Me.Close()
        Else
            MessageBox.Show("Please select from the list.")
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

        ComboBox2.DataSource = PaymentTermBS
        ComboBox2.DisplayMember = "paymenttermdesc"
        ComboBox2.ValueMember = "paymenttermid"

      
        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        DateTimePicker1.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", DRV, "paymenttermid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker1.DataBindings.Add(New Binding("Text", DRV, "effectivedate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        If DRV.Row.RowState = DataRowState.Detached Then
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1

        End If
    End Sub

    Private Function validcombo1() As Boolean
        Dim myret As Boolean = False
        Dim drv = ComboBox1.SelectedItem
        If Not IsNothing(drv) Then
            Me.DRV.Item("vendorname") = drv.item("vendorname").ToString.Trim
            Me.DRV.Item("shortname") = drv.item("shortname").ToString.Trim
            RaiseEvent RefreshDataGridView(Me, New EventArgs)
            myret = True
        End If
        Return myret
    End Function
    Private Function validcombo2() As Boolean
        Dim myret As Boolean = False
        Dim drv = ComboBox2.SelectedItem
        If Not IsNothing(drv) Then
            Me.DRV.Item("paymenttermdesc") = drv.item("paymenttermdesc")
            myret = True
            RaiseEvent RefreshDataGridView(Me, New EventArgs)
        End If
        Return myret
    End Function



    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
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
                Dim myform = New FormHelper(PaymentTermBSHelper)
                myform.DataGridView1.Columns(0).DataPropertyName = "paymenttermdesc"
                If myform.ShowDialog = DialogResult.OK Then                   
                    Dim drv As DataRowView = PaymentTermBSHelper.Current
                    'Me.DRV.BeginEdit()
                    Me.DRV.Row.Item("paymenttermdesc") = drv.Row.Item("paymenttermdesc")                    
                    'Me.DRV.Row.Item("pm") = drv.Row.Item("pm")
                    'Me.DRV.Row.Item("spm") = drv.Row.Item("spm")
                    'Need bellow code to sync with combobox
                    Dim myposition = PaymentTermBS.Find("paymenttermid", drv.Row.Item("paymenttermid"))                   
                    PaymentTermBS.Position = myposition
                End If
            
        End Select
        RaiseEvent RefreshDataGridView(Me, New EventArgs)
    End Sub

End Class
