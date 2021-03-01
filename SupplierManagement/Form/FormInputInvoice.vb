Imports System.Text
Imports SupplierManagement.PublicClass
Public Class FormInputInvoice
    Dim DS As DataSet
    Dim toolingBS As New BindingSource
    Dim toolingInvoiceBS As BindingSource
    WithEvents ToolingPaymentBS As New BindingSource
    Dim InvDrv As DataRowView
    Dim myparent As Object
    Dim CurrencyList As New List(Of Currency)

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ToolingPaymentBS.Count > 0 Then
            Try
                InvDrv.Item("totalamount") = getTotal()
                InvDrv.EndEdit()
                'toolingInvoiceBS.EndEdit()
                'ToolingPaymentBS.EndEdit()

                If Not Me.validate Then
                    commit()
                    Me.DialogResult = Windows.Forms.DialogResult.None
                End If
            Catch ex As Exception
                Me.DialogResult = Windows.Forms.DialogResult.None
                MessageBox.Show(ex.Message, "Get Total")
            End Try
            
        End If
        
    End Sub
    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()

        If TextBox33.Text = "" Then
            ErrorProvider1.SetError(TextBox33, "Supplier Invoice No cannot be blank.")
            myret = False
        Else
            ErrorProvider1.SetError(TextBox33, "")
        End If


        For Each drv As DataRowView In ToolingPaymentBS.List
            If Not ValidateRecord(drv) Then
                myret = False
            End If
        Next
        Return myret
    End Function
    Public Sub New(ByRef DS As DataSet, ByRef ToolinginvoiceBS As BindingSource, ByVal sender As Object)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.DS = DS

        toolingBS.DataSource = DS.Tables("ToolingListDT")
        ComboBox4.DataSource = toolingBS
        ComboBox4.DisplayMember = "displaymember"
        ComboBox4.ValueMember = "id"
        ComboBox4.DataBindings.Add("SelectedValue", ToolingPaymentBS, "toolinglistid", True, DataSourceUpdateMode.OnPropertyChanged)
        ComboBox4.SelectedIndex = -1




        Me.toolingInvoiceBS = ToolinginvoiceBS
        'ToolingPaymentBS.DataSource = DS.Tables("ToolingPayment")
        ToolingPaymentBS.DataSource = ToolinginvoiceBS
        ToolingPaymentBS.DataMember = "TITPRel"
        InvDrv = ToolinginvoiceBS.Current

        DataGridView3.AutoGenerateColumns = False
        DataGridView3.DataSource = ToolingPaymentBS

        TextBox36.DataBindings.Add(New Binding("Text", ToolingPaymentBS, "cost", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
        TextBox37.DataBindings.Add(New Binding("Text", ToolingPaymentBS, "balance", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
        TextBox38.DataBindings.Add(New Binding("Text", ToolingPaymentBS, "invoiceamount", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
        TextBox1.DataBindings.Add(New Binding("Text", ToolinginvoiceBS, "totalamount", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))

        TextBox4.DataBindings.Add(New Binding("Text", ToolinginvoiceBS, "proformainvoice", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox5.DataBindings.Add(New Binding("Text", ToolinginvoiceBS, "currency", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox6.DataBindings.Add(New Binding("Text", ToolinginvoiceBS, "amount", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))

        TextBox33.DataBindings.Add(New Binding("Text", ToolinginvoiceBS, "invoiceno", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox18.DataBindings.Add(New Binding("Text", ToolinginvoiceBS, "pct", True, DataSourceUpdateMode.OnPropertyChanged, "", "##0.00"))
        TextBox34.DataBindings.Add(New Binding("Text", ToolinginvoiceBS, "description", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker2.DataBindings.Add(New Binding("Text", ToolinginvoiceBS, "invoicedate", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        TextBox35.DataBindings.Add(New Binding("Text", ToolingPaymentBS, "pct", True, DataSourceUpdateMode.OnPropertyChanged, "", "##0.00"))
        TextBox2.DataBindings.Add(New Binding("Text", ToolingPaymentBS, "invoiceamount", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
        TextBox19.DataBindings.Add(New Binding("Text", ToolingPaymentBS, "exrate", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
        TextBox39.DataBindings.Add(New Binding("Text", ToolingPaymentBS, "description", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        Dim myCurr() As String = DS.Tables(15).Rows(0).Item("cvalue").ToString.Split(",")

        For i = 0 To myCurr.Count - 1
            CurrencyList.Add(New Currency(myCurr(i)))
        Next

        'CurrencyList.Add(New Currency("USD"))
        'CurrencyList.Add(New Currency("EUR"))
        'CurrencyList.Add(New Currency("CNY"))

        ComboBox2.DataSource = CurrencyList
        ComboBox2.DisplayMember = "currency"
        ComboBox2.ValueMember = "currency"
        ComboBox2.SelectedIndex = -1
        ComboBox2.DataBindings.Add("SelectedValue", ToolingPaymentBS, "currency", True, DataSourceUpdateMode.OnPropertyChanged)
        'ComboBox4.DataBindings.Add("SelectedValue", APBS, "applicantname", True, DataSourceUpdateMode.OnPropertyChanged)
        myparent = sender

        If HelperClass1.UserInfo.IsFinance Then
            'GroupBox1.Enabled = False
            'GroupBox6.Enabled = False
            'For Each myobj In TabControl1.Controls
            'If myobj.GetType() Is GetType(TabPage) Then
            For Each obj As Control In Me.Controls
                If obj.GetType() Is GetType(GroupBox) Then
                    For Each c As Control In obj.Controls
                        If c.GetType() Is GetType(TextBox) Then
                            Dim Txt As TextBox = CType(c, TextBox)
                            Txt.BackColor = Color.WhiteSmoke
                            Txt.ReadOnly = True
                        End If
                        If c.GetType() Is GetType(DateTimePicker) Then
                            Dim dt As DateTimePicker = CType(c, DateTimePicker)
                            dt.Enabled = False
                        End If
                        If c.GetType() Is GetType(Button) Then
                            Dim b As Button = CType(c, Button)
                            b.Enabled = False
                        End If
                        If c.GetType() Is GetType(ComboBox) Then
                            Dim cb As ComboBox = CType(c, ComboBox)
                            cb.Enabled = False
                        End If
                    Next
                End If
            Next
            'End If
            'Next
            DataGridView3.ContextMenuStrip = Nothing
        End If
    End Sub



    Private Sub BatchAddToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BatchAddToolStripMenuItem.Click
        'batch Add
        Try


            Dim drv As DataRowView = InvDrv
            ErrorProvider1.SetError(TextBox18, "")
            If IsDBNull(drv.Item("pct")) Then
                ErrorProvider1.SetError(TextBox18, "Invoice in % cannot be blank.")
                Exit Sub
            End If

            For Each mydrv As DataRowView In toolingBS.List


                Dim myrow As DataRowView = ToolingPaymentBS.AddNew
                'myrow.EndEdit()
                'myrow.BeginEdit()
                myrow.Row.Item("exrate") = 1
                myrow.Row.Item("invoiceamount") = (drv.Item("pct") / 100) * mydrv.Item("cost")
                myrow.Row.Item("toolinglistid") = mydrv.Item("id")

                myrow.Row.Item("currency") = "USD"

                myrow.Row.Item("pct") = drv.Item("pct")
                myrow.Item("displaymember") = mydrv.Item("toolinglistid")
                myrow.Item("suppliermoldno") = mydrv.Item("suppliermoldno")
                myrow.Item("toolsdescription") = mydrv.Item("toolsdescription")

                'Find Tooling List detail
                'Dim mypos = toolingBS.Find("id", myrow.Item("toolinglistid"))
                'toolingBS.Position = mypos
                'Dim drvTooling As DataRowView = toolingBS.Current
                'get data from 
                TextBox36.Text = Format(mydrv.Item("cost"), "#,##0.00")
                TextBox37.Text = Format(mydrv.Item("balance"), "#,##0.00")
                TextBox38.Text = Format(myrow.Item("invoiceamount") * myrow.Item("exrate"), "#,##0.00")
                myrow.EndEdit()
            Next
            TextBox1.Text = getTotal()
            TextBox3.Text = getTotal("CNY")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Batch Payment")
        End Try
    End Sub


    Private Sub AddToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddToolStripMenuItem.Click
        'Add NewRecord
        Dim drv As DataRowView = Nothing
        Try
            drv = ToolingPaymentBS.AddNew
            drv.Item("invoiceid") = InvDrv.Item("id")
            ComboBox2.Text = "USD"
            TextBox19.Text = 1
            drv.Item("currency") = "USD"
            drv.Item("exrate") = 1
            ComboBox4.SelectedIndex = -1
            drv.Item("pct") = InvDrv.Item("pct")
            TextBox35.Text = String.Format("{0:##0.00}", InvDrv.Item("pct"))
            TextBox36.Text = ""
            TextBox37.Text = ""
            TextBox38.Text = ""
            drv.EndEdit()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            If Not IsNothing(drv) Then
                drv.CancelEdit()
            End If

        End Try
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        'Delete Record
       
        If Not IsNothing(toolingInvoiceBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView3.SelectedRows
                    Try
                        ToolingPaymentBS.RemoveAt(drv.Index)
                    Catch ex As Exception

                    End Try

                Next
            End If
        End If
    End Sub


    Private Sub ComboBox4_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectionChangeCommitted
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next
        Dim pos = toolingBS.Find("id", ComboBox4.SelectedValue)
        toolingBS.Position = pos
        Dim drv As DataRowView = toolingBS.Current
        Dim tpdrv As DataRowView = ToolingPaymentBS.Current
        TextBox36.Text = Format(drv.Item("cost"), "#,##0.00")
        TextBox37.Text = Format(drv.Item("balance"), "#,##0.00")
        tpdrv.Item("displaymember") = drv.Item("toolinglistid")
        tpdrv.Item("suppliermoldno") = drv.Item("suppliermoldno")
        tpdrv.Item("toolsdescription") = drv.Item("toolsdescription")
        DataGridView3.Invalidate()
        TextBox35_Validated(Me.TextBox35, New System.EventArgs)

    End Sub

    Private Sub ToolingPaymentBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles ToolingPaymentBS.ListChanged
        'TextBox2.Enabled = Not IsNothing(ToolingPaymentBS.Current)
        'TextBox19.Enabled = Not IsNothing(ToolingPaymentBS.Current)
        'TextBox35.Enabled = Not IsNothing(ToolingPaymentBS.Current)
        'TextBox39.Enabled = Not IsNothing(ToolingPaymentBS.Current)
        'ComboBox4.Enabled = Not IsNothing(ToolingPaymentBS.Current)
        'ComboBox2.Enabled = Not IsNothing(ToolingPaymentBS.Current)

        'If Not IsNothing(ToolingPaymentBS.Current) Then
        '    'Provide the data from the latest ToolingBS
        '    Dim drvT As DataRowView = ToolingPaymentBS.Current
        '    Try
        '        Dim mypos = toolingBS.Find("id", drvT.Row.Item("toolinglistid"))
        '        toolingBS.Position = mypos
        '        Dim drvTooling As DataRowView = toolingBS.Current
        '        'get data from 
        '        TextBox36.Text = Format(drvTooling.Item("cost"), "#,##0.00")
        '        TextBox37.Text = Format(drvTooling.Item("balance"), "#,##0.00")
        '        TextBox38.Text = Format(drvT.Item("invoiceamount") * drvT.Item("exrate"), "#,##0.00")

        '    Catch ex As Exception

        '    End Try

        'End If
        CheckInitValue()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged, TextBox19.TextChanged, TextBox39.TextChanged
        Dim drv As DataRowView = ToolingPaymentBS.Current
        If TextBox2.Text = "" Then
            TextBox2.Text = 0
        End If
        Try          
            'TextBox38.Text = Format(drv.Item("invoiceamount") * drv.Item("exrate"), "#,##0.00")
            TextBox38.Text = Format(IIf(TextBox2.Text = "", 0, TextBox2.Text) * drv.Item("exrate"), "#,##0.00")
        Catch ex As Exception
            ' MessageBox.Show(ex.Message)
        End Try
        TextBox1.Text = getTotal()
        TextBox3.Text = getTotal("CNY")
        DataGridView3.Invalidate()

    End Sub



    Private Sub TextBox35_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox33.TextChanged, TextBox34.TextChanged
        myparent.datagridview3.invalidate()

    End Sub

    Private Sub TextBox35_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox35.Validated
        Dim drv As DataRowView = ToolingPaymentBS.Current

        Try
            TextBox2.Text = Format(TextBox36.Text * drv.Item("pct") / 100, "#,##0.00")
            TextBox38.Text = Format(drv.Item("invoiceamount") * drv.Item("exrate"), "#,##0.00")

        Catch ex As Exception
            ' MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub TextBox33_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox33.TextChanged
        DataGridView3.Invalidate()
    End Sub


    Private Sub TextBox38_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox38.TextChanged
        TextBox1.Text = getTotal()
        TextBox3.Text = getTotal("CNY")
        Try
            'Dim pos = toolingBS.Find("id", DirectCast(ToolingPaymentBS.Current, DataRowView).Row.Item("toolinglistid"))
            'toolingBS.Position = pos
            'DirectCast(ToolingPaymentBS.Current, DataRowView).Row.Item("pct") = DirectCast(toolingBS.Current, DataRowView).Row.Item("cost") / DirectCast(ToolingPaymentBS.Current, DataRowView).Row.Item("invoiceamount")
        Catch ex As Exception

        End Try
    End Sub

    Private Function getTotal() As String
        Dim myret As Decimal = 0
        Try
            For Each drv As DataRowView In ToolingPaymentBS.List
                myret = myret + drv.Item("invoiceamount") * drv.Item("exrate")
            Next           
        Catch ex As Exception

        End Try
        Return Format(myret, "#,##0.00")
    End Function

    Private Function getTotal(ByVal curr As String) As String
        Dim myret As Decimal = 0
        Try
            For Each drv As DataRowView In ToolingPaymentBS.List
                If drv.Item("currency") = curr Then
                    myret = myret + drv.Item("invoiceamount")
                End If

            Next
        Catch ex As Exception

        End Try
        Return Format(myret, "#,##0.00")
    End Function

    Private Function ValidateRecord(ByVal drv As DataRowView) As Boolean
        Dim myerror As New StringBuilder
        Dim myret As Boolean = True
        Try
            If IsDBNull(drv.Row.Item("toolinglistid")) Then
                myerror.Append("Tooling Id cannot be blank.")
                myret = False
            End If
            If IsDBNull(drv.Row.Item("invoiceamount")) Then
                If myerror.Length > 0 Then
                    myerror.Append(",")
                End If
                myerror.Append("Tooling amount cannot be blank.")
                myret = False
            End If
            If IsDBNull(drv.Row.Item("currency")) Then
                If myerror.Length > 0 Then
                    myerror.Append(",")
                End If
                myerror.Append("Currency cannot be blank.")
                myret = False
            End If
            If IsDBNull(drv.Row.Item("exrate")) Then
                If myerror.Length > 0 Then
                    myerror.Append(",")
                End If
                myerror.Append("Exchange Rate cannot be blank.")
                myret = False
            End If
        Catch ex As Exception
            myret = False
            'MessageBox.Show("Validate::" & ex.Message)
        End Try
        drv.Row.RowError = myerror.ToString
        drv.EndEdit() 'Need to Attach the detached record
        Return myret


    End Function

    Private Sub ToolingPaymentBS_PositionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolingPaymentBS.PositionChanged
        CheckInitValue()
    End Sub
    Public Sub CheckInitValue()
        Try
            If Not HelperClass1.UserInfo.IsFinance Then


                TextBox2.Enabled = Not IsNothing(ToolingPaymentBS.Current)
                TextBox19.Enabled = Not IsNothing(ToolingPaymentBS.Current)
                TextBox35.Enabled = Not IsNothing(ToolingPaymentBS.Current)
                TextBox39.Enabled = Not IsNothing(ToolingPaymentBS.Current)
                ComboBox4.Enabled = Not IsNothing(ToolingPaymentBS.Current)
                ComboBox2.Enabled = Not IsNothing(ToolingPaymentBS.Current)

                If Not IsNothing(ToolingPaymentBS.Current) Then
                    'Provide the data from the latest ToolingBS
                    Dim drvT As DataRowView = ToolingPaymentBS.Current
                    Try
                        'Dim mypos = toolingBS.Find("id", drvT.Row.Item("toolinglistid"))
                        'toolingBS.Position = mypos
                        'Dim drvTooling As DataRowView = toolingBS.Current
                        ''get data from 
                        'TextBox36.Text = Format(drvTooling.Item("cost"), "#,##0.00")
                        'TextBox37.Text = Format(drvTooling.Item("balance"), "#,##0.00")
                        'TextBox38.Text = Format(drvT.Item("invoiceamount") * drvT.Item("exrate"), "#,##0.00")
                    Catch ex As Exception

                    End Try

                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub FormInputInvoice_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If ToolingPaymentBS.Count > 0 Then
            If Not Me.validate Then
                'Me.DialogResult = Windows.Forms.DialogResult.None
                e.Cancel = True
            Else
                InvDrv.Item("totalamount") = getTotal()
                InvDrv.EndEdit()
                commit()
            End If
        Else
            If IsDBNull(InvDrv.Item("invoiceno")) Then
                toolingInvoiceBS.RemoveCurrent()
            End If
        End If        
    End Sub

    Private Sub commit()

    End Sub

    Private Sub DataGridView3_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView3.DataError

    End Sub

    Private Sub TextBox35_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox35.TextChanged

    End Sub


End Class

Public Class Currency
    Public Property currency As String
    Public Sub New(ByVal _currency As String)
        Me.currency = _currency
    End Sub
End Class