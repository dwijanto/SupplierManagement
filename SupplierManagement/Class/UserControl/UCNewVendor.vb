Public Class UCNewVendor
    Dim FamilyCodeBS As BindingSource
    Dim SubFamilyCodeBS As BindingSource
    Dim PaymentTermBS As BindingSource
    Dim VendorIndFamBS As BindingSource
    Dim GroupSBUBS As BindingSource
    Dim FamilyBS As BindingSource
    Dim CurrencyBS As BindingSource
    Dim _suppliermodificationid As String
    Dim DRV As DataRowView
    Dim TxType As TxEnum
    Public Property SupplierModificationID As String
        Get
            Return _suppliermodificationid
        End Get
        Set(ByVal value As String)
            _suppliermodificationid = value
            TextBox1.Text = _suppliermodificationid
        End Set
    End Property
    Public Sub HistoryMode()
        Button2.Enabled = False
        Button4.Enabled = False
    End Sub
    Sub BindingControls(ByVal txtype As TxEnum, ByVal drv As DataRowView, ByVal ApplicantNameBS As BindingSource, ByVal FamilyCodeBS As BindingSource, ByVal SubFamilyCodeBS As BindingSource, ByVal PaymentTermBS As BindingSource, ByVal VendorStatusBS As BindingSource, ByVal VendorTypeBS As BindingSource, ByVal VendorIndFamBS As BindingSource, ByVal GSMBS As BindingSource, ByVal GroupSBUBS As BindingSource, ByVal FamilyBS As BindingSource, ByVal PMBS As BindingSource, ByVal TechnologyBS1 As BindingSource, ByVal TechnologyBS2 As BindingSource, ByVal TechnologyBS3 As BindingSource, ByVal TechnologyBS4 As BindingSource, ByVal ApprovalDept1BS As BindingSource, ByVal ApprovalDept2BS As BindingSource, ByVal FinanceBS As BindingSource, ByVal CurrencyBS As BindingSource)
        Me.FamilyCodeBS = FamilyCodeBS
        Me.SubFamilyCodeBS = SubFamilyCodeBS
        Me.PaymentTermBS = PaymentTermBS
        Me.VendorIndFamBS = VendorIndFamBS
        Me.GroupSBUBS = GroupSBUBS
        Me.FamilyBS = FamilyBS
        Me.CurrencyBS = CurrencyBS
        Me.DRV = drv
        Me.TxType = txtype

        TextBox1.DataBindings.Clear() 'Suppliermodificationid M

        ComboBox1.DataBindings.Clear() 'applicantname M       
        'TextBox7.DataBindings.Clear() 'Currency M
        TextBox13.DataBindings.Clear() 'ecoqualitycontactemail
        TextBox14.DataBindings.Clear() 'ecoqualitycontactname

        ComboBox2.DataBindings.Clear() 'VendorStatus EX
        TextBox2.DataBindings.Clear() 'BR EX
        TextBox3.DataBindings.Clear() 'Vendor Name EX
        TextBox4.DataBindings.Clear() 'Short Name EX
        TextBox5.DataBindings.Clear() 'Vendor Name EX
        TextBox6.DataBindings.Clear() 'Short Name EX
        TextBox7.DataBindings.Clear() 'Long Name EX
        TextBox8.DataBindings.Clear() 'Ifamilycode Ex
        ComboBox3.DataBindings.Clear()
        ComboBox4.DataBindings.Clear()
        ComboBox5.DataBindings.Clear()
        ComboBox6.DataBindings.Clear()
        ComboBox7.DataBindings.Clear()
        ComboBox8.DataBindings.Clear()
        ComboBox9.DataBindings.Clear()
        ComboBox10.DataBindings.Clear()
        ComboBox11.DataBindings.Clear()
        ComboBox12.DataBindings.Clear()

        ComboBox13.DataBindings.Clear()
        ComboBox14.DataBindings.Clear()
        ComboBox15.DataBindings.Clear()
        ComboBox16.DataBindings.Clear() 'Currency

        TextBox15.DataBindings.Clear() 'email EX
        TextBox16.DataBindings.Clear() 'fax EX
        TextBox17.DataBindings.Clear() 'tel EX

        DateTimePicker1.DataBindings.Clear()

        ComboBox1.DataSource = ApplicantNameBS
        ComboBox1.DisplayMember = "username"
        ComboBox1.ValueMember = "username"

        ComboBox2.DataSource = VendorStatusBS
        ComboBox2.DisplayMember = "status"
        ComboBox2.ValueMember = "statusid"

        ComboBox3.DataSource = VendorTypeBS
        ComboBox3.DisplayMember = "producttype"
        ComboBox3.ValueMember = "producttype"

        ComboBox4.DataSource = GSMBS
        ComboBox4.DisplayMember = "gsm"
        ComboBox4.ValueMember = "gsm"

        ComboBox5.DataSource = GroupSBUBS
        ComboBox5.DisplayMember = "groupname"
        ComboBox5.ValueMember = "groupid"


        ComboBox6.DataSource = PaymentTermBS
        ComboBox6.DisplayMember = "paymenttermdesc"
        ComboBox6.ValueMember = "paymenttermid"

        ComboBox7.DataSource = PMBS
        ComboBox7.DisplayMember = "pm"
        ComboBox7.ValueMember = "pm"

        ComboBox8.DataSource = TechnologyBS1
        ComboBox8.DisplayMember = "technologyname"
        ComboBox8.ValueMember = "id"

        ComboBox9.DataSource = TechnologyBS2
        ComboBox9.DisplayMember = "technologyname"
        ComboBox9.ValueMember = "id"

        ComboBox10.DataSource = TechnologyBS3
        ComboBox10.DisplayMember = "technologyname"
        ComboBox10.ValueMember = "id"

        ComboBox11.DataSource = TechnologyBS4
        ComboBox11.DisplayMember = "technologyname"
        ComboBox11.ValueMember = "id"

        ComboBox12.DataSource = FamilyBS
        ComboBox12.DisplayMember = "familydesc"
        ComboBox12.ValueMember = "familyid"

        ComboBox13.DataSource = ApprovalDept1BS
        ComboBox13.DisplayMember = "username"
        ComboBox13.ValueMember = "userid"

        ComboBox14.DataSource = ApprovalDept2BS
        ComboBox14.DisplayMember = "username"
        ComboBox14.ValueMember = "userid"

        ComboBox15.DataSource = FinanceBS
        ComboBox15.DisplayMember = "username"
        ComboBox15.ValueMember = "userid"

        ComboBox16.DataSource = CurrencyBS
        ComboBox16.DisplayMember = "cvalue"
        ComboBox16.ValueMember = "cvalue"

        TextBox1.DataBindings.Add(New Binding("Text", drv, "suppliermodificationid", True))
        TextBox2.DataBindings.Add(New Binding("Text", drv, "br", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox3.DataBindings.Add(New Binding("Text", drv, "vendorname", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox4.DataBindings.Add(New Binding("Text", drv, "shortname", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox5.DataBindings.Add(New Binding("Text", drv, "familycode", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox6.DataBindings.Add(New Binding("Text", drv, "subfamilycode", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox7.DataBindings.Add(New Binding("Text", drv, "longname", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox8.DataBindings.Add(New Binding("Text", drv, "ifamilycode", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox13.DataBindings.Add(New Binding("Text", drv, "ecoqualitycontactemail", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox14.DataBindings.Add(New Binding("Text", drv, "ecoqualitycontactname", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox15.DataBindings.Add(New Binding("Text", drv, "email", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox16.DataBindings.Add(New Binding("Text", drv, "fax", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox17.DataBindings.Add(New Binding("Text", drv, "tel", True, DataSourceUpdateMode.OnPropertyChanged))



        ComboBox1.DataBindings.Add(New Binding("SelectedValue", drv, "applicantname", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", drv, "vendorstatus", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox3.DataBindings.Add(New Binding("SelectedValue", drv, "producttype", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox4.DataBindings.Add(New Binding("SelectedValue", drv, "vendorgsm", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox5.DataBindings.Add(New Binding("SelectedValue", drv, "groupsbuforfp", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox6.DataBindings.Add(New Binding("SelectedValue", drv, "paymentterm", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox7.DataBindings.Add(New Binding("SelectedValue", drv, "pm", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox8.DataBindings.Add(New Binding("SelectedValue", drv, "technology1", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox9.DataBindings.Add(New Binding("SelectedValue", drv, "technology2", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox10.DataBindings.Add(New Binding("SelectedValue", drv, "technology3", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox11.DataBindings.Add(New Binding("SelectedValue", drv, "technology4", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox12.DataBindings.Add(New Binding("SelectedValue", drv, "vendorfamilyfp", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox13.DataBindings.Add(New Binding("SelectedValue", drv, "approvaldept", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox14.DataBindings.Add(New Binding("SelectedValue", drv, "approvaldept2", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox15.DataBindings.Add(New Binding("SelectedValue", drv, "approvalfc", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox16.DataBindings.Add(New Binding("SelectedValue", drv, "currency", True, DataSourceUpdateMode.OnPropertyChanged))
        DateTimePicker1.DataBindings.Add(New Binding("Text", drv, "applicantdate", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox1.ReadOnly = True

        Select Case txtype
            Case TxEnum.UpdateRecord, TxEnum.ValidateRecord
                DisplayValidControl()
        End Select
        'If txtype = TxEnum.UpdateRecord Then
        '    DisplayValidControl()
        'End If

    End Sub

    Private Sub DisplayValidControl()
        EnabledCombobox()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim mydialog As New DialogFamilySubVendor(FamilyCodeBS, SubFamilyCodeBS)
        If mydialog.ShowDialog = DialogResult.OK Then
            Dim familydr As DataRowView = FamilyCodeBS.Current
            Dim subfamilydr As DataRowView = SubFamilyCodeBS.Current
            If TextBox5.Text.Length > 0 Then
                TextBox5.Text = String.Format("{0},{1}", TextBox5.Text, familydr.Item("familyvc"))
            Else
                TextBox5.Text = String.Format("{0}", familydr.Item("familyvc"))
            End If
            If TextBox6.Text.Length > 0 Then
                TextBox6.Text = String.Format("{0},{1}", TextBox6.Text, subfamilydr.Item("subfamilyvc"))
            Else
                TextBox6.Text = String.Format("{0}", subfamilydr.Item("subfamilyvc"))
            End If        
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If MessageBox.Show("Do you want to clear Family and SubFamily contents?", "Clear Content", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.OK Then
            TextBox5.Text = String.Empty
            TextBox6.Text = String.Empty
        End If
    End Sub

    'Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim myhelper As New FormHelper(PaymentTermBS)


    'End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim myform = New FormHelper(VendorIndFamBS)
        myform.DataGridView1.Columns(0).DataPropertyName = "helperdescription"
        If myform.ShowDialog = DialogResult.OK Then
            Dim drv As DataRowView = VendorIndFamBS.Current
            TextBox8.Text = "" & drv.Row.Item("familycode")
        End If
    End Sub

    Private Sub ComboBox3_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectionChangeCommitted       
        EnabledCombobox()
    End Sub

    Sub EnabledCombobox()
        Dim mydrv As DataRowView = ComboBox3.SelectedItem
        If Not IsNothing(mydrv) Then
            ComboBox5.Enabled = False
            ComboBox4.Enabled = False
            ComboBox12.Enabled = False
            ComboBox7.Enabled = False
            ComboBox8.Enabled = False
            ComboBox9.Enabled = False
            ComboBox10.Enabled = False
            ComboBox11.Enabled = False

            If mydrv.Row.Item(0).ToString.Contains("FP") Then
                ComboBox5.Enabled = True
                ComboBox4.Enabled = True
                ComboBox12.Enabled = True

                'ComboBox8.Enabled = False
                'ComboBox9.Enabled = False
                'ComboBox10.Enabled = False
                'ComboBox11.Enabled = False
            End If

            If mydrv.Row.Item(0).ToString.Contains("CP") Then
                'ComboBox5.Enabled = False
                'ComboBox4.Enabled = False
                'ComboBox12.Enabled = False
                ComboBox7.Enabled = True
                ComboBox8.Enabled = True
                ComboBox9.Enabled = True
                ComboBox10.Enabled = True
                ComboBox11.Enabled = True

            End If
        End If


    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim myform = New FormHelper(VendorIndFamBS)
        myform.DataGridView1.Columns(0).DataPropertyName = "helperdescription"
        If myform.ShowDialog = DialogResult.OK Then
            Dim drv As DataRowView = VendorIndFamBS.Current
            TextBox8.Text = drv.Row.Item("familycode")
        End If
    End Sub


    Public Overloads Function Validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(TextBox3, "")
        ErrorProvider1.SetError(TextBox4, "")
        ErrorProvider1.SetError(ComboBox13, "")
        ErrorProvider1.SetError(ComboBox14, "")
        If TextBox3.Text = "" Then
            myret = False
            ErrorProvider1.SetError(TextBox3, "Field cannot be blank.")
        End If

        If TextBox4.Text = "" Then
            myret = False
            ErrorProvider1.SetError(TextBox4, "Field cannot be blank.")
        End If

        If IsDBNull(DRV.Row.Item("approvaldept")) Then
            myret = False
            ErrorProvider1.SetError(ComboBox13, "Please select from the list.")
        End If
        If IsDBNull(DRV.Row.Item("approvaldept2")) Then
            myret = False
            ErrorProvider1.SetError(ComboBox14, "Please select from the list.")
        End If
        If Not IsDBNull(DRV.Row.Item("id")) Then
            DRV.Item("suppliermodificationid") = String.Format("{0}_{1:yyyyMMdd}_{2:0000}", TextBox4.Text, DateTimePicker1.Value.Date, DRV.Item("id"))
        End If
   
        Return myret
    End Function

    Private Sub ComboBox13_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox13.SelectionChangeCommitted, ComboBox14.SelectionChangeCommitted, ComboBox15.SelectionChangeCommitted
        Dim myCB As ComboBox = DirectCast(sender, ComboBox)
        Dim cbdrv As DataRowView = myCB.SelectedItem
        Select Case myCB.Name
            Case "ComboBox13"
                DRV.Item("spmusername") = cbdrv.Item("username")
            Case "ComboBox14"
                DRV.Item("pdusername") = cbdrv.Item("username")
            Case "ComboBox15"
                DRV.Item("fcusername") = cbdrv.Item("username")
        End Select


    End Sub



End Class
