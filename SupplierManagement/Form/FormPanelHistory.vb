Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class FormPanelHistory
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents PHBS As BindingSource
    Dim VendorBS As BindingSource
    Dim VendorBSHelper As BindingSource
    Dim CategoryBS As BindingSource
    Dim FPBS As BindingSource
    Dim CPBS As BindingSource

    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim PanelStatuslist As String() = {"FP", "CP"}
    Dim myFields As String() = {"vendorcode", "vendorname", "shortname", "category"}

    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        sb.Append("select ph.vendorcode::text,v.vendorname::text,v.shortname::text,ph.validfrom,sc.category,ps.paneldescription as fppanel,ps2.paneldescription as cppanel,ph.suppliercategoryid,ph.fp,ph.cp,ph.rank,ph.id,ph.userid from doc.panelhistory ph " &
                   " left join supplierscategory sc on sc.supplierscategoryid = ph.suppliercategoryid " &
                   " left join vendor v on v.vendorcode = ph.vendorcode " &
                   " left join doc.panelstatus ps on ps.id = ph.fp " &
                   " left join doc.panelstatus ps2 on ps2.id = ph.cp order by vendorname, validfrom desc;")
        sb.Append("select null::text as vendorcode,''::text as description,''::text as vendorname,''::text as shortname union all (select vendorcode::text, vendorcode::text || ' - ' || vendorname::text as description,vendorname::text,shortname::text from vendor order by vendorname);")
        sb.Append("select null as supplierscategoryid,''::text as category union all (select supplierscategoryid ,category::text from supplierscategory order by category);")
        sb.Append("select null::integer as id, ''::character varying as paneldescription union all (select id, paneldescription from doc.panelstatus where panelstatus = 'FP' order by panelstatus,paneldescription) ;")
        sb.Append("select null::integer as id, ''::character varying as paneldescription union all (select id, paneldescription from doc.panelstatus where panelstatus = 'CP' order by panelstatus,paneldescription);")


        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "PanelHistory"
                DS.Tables(1).TableName = "Vendor"
                DS.Tables(2).TableName = "Category"
                DS.Tables(3).TableName = "FP"
                DS.Tables(4).TableName = "CP"

            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(4, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel1.Text = message
                    Case 4
                        Try
                            PHBS = New BindingSource
                            VendorBS = New BindingSource
                            VendorBSHelper = New BindingSource
                            CategoryBS = New BindingSource
                            FPBS = New BindingSource
                            CPBS = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("id")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns("id").AutoIncrement = True
                            DS.Tables(0).Columns("id").AutoIncrementSeed = 0
                            DS.Tables(0).Columns("id").AutoIncrementStep = -1
                            DS.Tables(0).TableName = "PanelHistory"

                            PHBS.DataSource = DS.Tables(0)
                            VendorBS.DataSource = New DataView(DS.Tables(1))
                            VendorBSHelper.DataSource = New DataView(DS.Tables(1))

                            CategoryBS.DataSource = DS.Tables(2)
                            FPBS.DataSource = DS.Tables(3)
                            CPBS.DataSource = DS.Tables(4)


                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = PHBS
                            DataGridView1.RowTemplate.Height = 22

                            'Combobox Category
                            ComboBox1.DataBindings.Clear()
                            'combobox FP
                            ComboBox2.DataBindings.Clear()
                            'Combobox Vendor
                            ComboBox3.DataBindings.Clear()
                            'Combobox CP
                            ComboBox4.DataBindings.Clear()
                            TextBox1.DataBindings.Clear()
                            DateTimePicker1.DataBindings.Clear()

                            ComboBox1.DataSource = CategoryBS
                            ComboBox1.DisplayMember = "category"
                            ComboBox1.ValueMember = "supplierscategoryid"
                            ComboBox1.DataBindings.Add("SelectedValue", PHBS, "suppliercategoryid", True, DataSourceUpdateMode.OnPropertyChanged)

                            ComboBox3.DataSource = VendorBS
                            ComboBox3.DisplayMember = "description"
                            ComboBox3.ValueMember = "vendorcode"
                            ComboBox3.DataBindings.Add("SelectedValue", PHBS, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged)

                            ComboBox2.DataSource = FPBS
                            ComboBox2.DisplayMember = "paneldescription"
                            ComboBox2.ValueMember = "id"
                            ComboBox2.DataBindings.Add("SelectedValue", PHBS, "fp", True, DataSourceUpdateMode.OnPropertyChanged)

                            ComboBox4.DataSource = CPBS
                            ComboBox4.DisplayMember = "paneldescription"
                            ComboBox4.ValueMember = "id"
                            ComboBox4.DataBindings.Add("SelectedValue", PHBS, "cp", True, DataSourceUpdateMode.OnPropertyChanged)
                            TextBox1.DataBindings.Add("Text", PHBS, "rank", True, DataSourceUpdateMode.OnPropertyChanged, "")
                            DateTimePicker1.DataBindings.Add("Text", PHBS, "validfrom", True, DataSourceUpdateMode.OnPropertyChanged, "")
                            ToolStripComboBox1.SelectedIndex = 1

                            If IsNothing(PHBS.Current) Then
                                ComboBox1.SelectedIndex = -1
                                ComboBox2.SelectedIndex = -1
                                ComboBox3.SelectedIndex = -1
                                ComboBox4.SelectedIndex = -1
                            End If

                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub
    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles PHBS.ListChanged
        ComboBox1.Enabled = Not IsNothing(PHBS.Current)
        ComboBox2.Enabled = Not IsNothing(PHBS.Current)
        ComboBox3.Enabled = Not IsNothing(PHBS.Current)
        ComboBox4.Enabled = Not IsNothing(PHBS.Current)
        Button8.Enabled = Not IsNothing(PHBS.Current)
        DateTimePicker1.Enabled = Not IsNothing(PHBS.Current)
    End Sub



    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = PHBS.AddNew()
        drv.Row.BeginEdit()
        drv.Row.Item("validfrom") = Date.Today
        'ComboBox1.SelectedIndex = -1
        'ComboBox2.SelectedIndex = -1
        'ComboBox3.SelectedIndex = -1
        'ComboBox4.SelectedIndex = -1
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        PHBS.EndEdit()
        If Me.validate Then
            Try
                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                Dim ds2 As DataSet
                ds2 = DS.GetChanges

                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                    If Not DbAdapter1.PanelHistoryTx(Me, mye) Then
                        MessageBox.Show(mye.message)
                        Exit Sub
                    End If
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    DataGridView1.Invalidate()
                    MessageBox.Show("Saved.")
                End If
            Catch ex As Exception
                MessageBox.Show(" Error:: " & ex.Message)
            End Try
        End If
        DataGridView1.Invalidate()
    End Sub



    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()
        Dim ds2 = DS.GetChanges

        'For Each drv As DataRowView In PHBS.List
        '    If drv.Row.RowState = DataRowState.Modified Or DataRowState.Added Then
        '        If Not validaterow(drv) Then
        '            myret = False
        '        End If
        '    End If
        'Next
        'Return myret
        If Not IsNothing(ds2) Then
            For Each drv As DataRow In ds2.Tables(0).Rows
                If drv.RowState = DataRowState.Modified Or drv.RowState = DataRowState.Added Then
                    If Not validaterow(drv) Then
                        myret = False
                    End If
                End If
            Next

        End If
        Return myret
       
    End Function

    Private Function validaterow(ByVal drv As DataRow) As Boolean
        Dim myret As Boolean = True
        Dim sb As New StringBuilder
        If IsDBNull(drv.Item("vendorcode")) Then
            myret = False
            sb.Append("Vendor cannot be blank")
        End If
        If IsDBNull(drv.Item("suppliercategoryid")) Then
            myret = False
            If sb.Length > 0 Then
                sb.Append(", ")
            End If
            sb.Append("Category cannot be blank.")
        End If
        drv.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(PHBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    PHBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted, ComboBox2.SelectionChangeCommitted, ComboBox3.SelectionChangeCommitted, ComboBox4.SelectionChangeCommitted
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)

        '1. Force the Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        Dim drv As DataRowView = PHBS.Current
        Select Case myobj.Name
            Case "ComboBox1"
                'drv.Row.Item("category") = myobj.Text
                Dim myselected As DataRowView = ComboBox1.SelectedItem
                'drv.Row.Item("suppliercategoryid") = myselected.Row.Item("supplierscategoryid")
                drv.Row.Item("category") = myselected.Row.Item("category")

            Case "ComboBox2"
                'drv.Row.Item("fppanel") = myobj.Text
                Dim myselected As DataRowView = ComboBox2.SelectedItem
                'drv.Row.Item("fp") = myselected.Row.Item("id")
                drv.Row.Item("fppanel") = myselected.Row.Item("paneldescription")
            Case "ComboBox3"
                Dim myselected As DataRowView = ComboBox3.SelectedItem
                'drv.Row.Item("vendorcode") = myselected.Row.Item("vendorcode")
                drv.Row.Item("vendorname") = myselected.Row.Item("vendorname")
                drv.Row.Item("shortname") = myselected.Row.Item("shortname")
            Case "ComboBox4"
                'drv.Row.Item("cppanel") = myobj.Text
                Dim myselected As DataRowView = ComboBox4.SelectedItem
                'drv.Row.Item("cp") = myselected.Row.Item("id")
                drv.Row.Item("cppanel") = myselected.Row.Item("paneldescription")
        End Select

        '2. Force the BindingSource to commit the value 
        PHBS.EndEdit()

        DataGridView1.Invalidate()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim myobj As Button = CType(sender, Button)
        Select Case myobj.Name
            Case "Button8"
                Dim myform = New FormHelper(VendorBSHelper)
                myform.DataGridView1.Columns(0).DataPropertyName = "description"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = VendorBSHelper.Current
                    Dim mydrv As DataRowView = PHBS.Current
                    mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
                    mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")

                    ''Need bellow code to sync with combobox bindingsource
                    Dim myposition = VendorBS.Find("vendorcode", drv.Row.Item("vendorcode"))
                    VendorBS.Position = myposition
                End If

        End Select
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        'PHBS.EndEdit()
        PHBS.Filter = ""
        If ToolStripTextBox1.Text <> "" Then
            Select Case ToolStripComboBox1.SelectedIndex
                Case 0
                    If Not IsNumeric(ToolStripTextBox1.Text) Then
                        ToolStripButton1.Select()
                        SendKeys.Send("{BACKSPACE}")
                    End If
            End Select
            PHBS.Filter = myFields(ToolStripComboBox1.SelectedIndex).ToString & " like '%" & sender.ToString.Replace("'", "''") & "%'"
        End If
    End Sub


End Class