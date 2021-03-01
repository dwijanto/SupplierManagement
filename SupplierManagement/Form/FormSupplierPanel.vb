Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Imports System.ComponentModel

Public Class FormSupplierPanel
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents SPBS As BindingSource
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
        sb.Append("select sp.vendorcode::text,sp.vendorcode::text as vendorcodetext,v.vendorname::text,v.shortname3::text as shortname,sc.category::text,sp.supplierscategoryid,sp.fp,ps.paneldescription as fppanel,sp.cp,ps2.paneldescription as cppanel,sp.rank,sp.supplierspanelid from supplierspanel sp " &
                   " inner join supplierscategory sc on sc.supplierscategoryid = sp.supplierscategoryid " &
                   " left join vendor v on v.vendorcode = sp.vendorcode " &
                   " left join doc.panelstatus ps on ps.id = sp.fp " &
                   " left join doc.panelstatus ps2 on ps2.id = sp.cp order by vendorname;")
        sb.Append("select null::text as vendorcode,''::text as description,''::text as vendorname ,''::text as shortname union all (select vendorcode::text, vendorcode::text || ' - ' || vendorname::text as description,vendorname::text,shortname3::text as shortname from vendor order by vendorname);")
        sb.Append("select null as supplierscategoryid,''::text as category union all (select supplierscategoryid ,category::text from supplierscategory order by category);")
        sb.Append("select null::integer as id, ''::character varying as paneldescription union all (select id, paneldescription from doc.panelstatus where panelstatus = 'FP' order by panelstatus,paneldescription) ;")
        sb.Append("select null::integer as id, ''::character varying as paneldescription union all (select id, paneldescription from doc.panelstatus where panelstatus = 'CP' order by panelstatus,paneldescription);")


        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "SupplierPanel"
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
                            SPBS = New BindingSource
                            VendorBS = New BindingSource
                            VendorBSHelper = New BindingSource
                            CategoryBS = New BindingSource
                            FPBS = New BindingSource
                            CPBS = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("supplierspanelid")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns("supplierspanelid").AutoIncrement = True
                            DS.Tables(0).Columns("supplierspanelid").AutoIncrementSeed = 0
                            DS.Tables(0).Columns("supplierspanelid").AutoIncrementStep = -1
                            DS.Tables(0).TableName = "SupplierPanel"

                            SPBS.DataSource = DS.Tables(0)
                            VendorBS.DataSource = New DataView(DS.Tables(1))
                            VendorBSHelper.DataSource = New DataView(DS.Tables(1))

                            CategoryBS.DataSource = DS.Tables(2)
                            FPBS.DataSource = DS.Tables(3)
                            CPBS.DataSource = DS.Tables(4)


                            DataGridView1.AutoGenerateColumns = False
                            DirectCast(DataGridView1.Columns("ColumnVendorName"), DataGridViewComboBoxColumn).DataSource = VendorBS
                            DirectCast(DataGridView1.Columns("ColumnVendorName"), DataGridViewComboBoxColumn).DisplayMember = "vendorname"
                            DirectCast(DataGridView1.Columns("ColumnVendorName"), DataGridViewComboBoxColumn).ValueMember = "vendorcode"
                            DataGridView1.Columns("ColumnVendorName").DataPropertyName = "vendorcode"

                            DirectCast(DataGridView1.Columns("ColumnShortname"), DataGridViewComboBoxColumn).DataSource = VendorBS
                            DirectCast(DataGridView1.Columns("ColumnShortname"), DataGridViewComboBoxColumn).DisplayMember = "shortname"
                            DirectCast(DataGridView1.Columns("ColumnShortname"), DataGridViewComboBoxColumn).ValueMember = "vendorcode"
                            DataGridView1.Columns("ColumnShortName").DataPropertyName = "vendorcode"

                            DirectCast(DataGridView1.Columns("ColumnCategory"), DataGridViewComboBoxColumn).DataSource = CategoryBS
                            DirectCast(DataGridView1.Columns("ColumnCategory"), DataGridViewComboBoxColumn).DisplayMember = "category"
                            DirectCast(DataGridView1.Columns("ColumnCategory"), DataGridViewComboBoxColumn).ValueMember = "supplierscategoryid"
                            DataGridView1.Columns("ColumnCategory").DataPropertyName = "supplierscategoryid"

                            DirectCast(DataGridView1.Columns("ColumnFP"), DataGridViewComboBoxColumn).DataSource = FPBS
                            DirectCast(DataGridView1.Columns("ColumnFP"), DataGridViewComboBoxColumn).DisplayMember = "paneldescription"
                            DirectCast(DataGridView1.Columns("ColumnFP"), DataGridViewComboBoxColumn).ValueMember = "id"
                            DataGridView1.Columns("ColumnFP").DataPropertyName = "fp"

                            DirectCast(DataGridView1.Columns("ColumnCP"), DataGridViewComboBoxColumn).DataSource = CPBS
                            DirectCast(DataGridView1.Columns("ColumnCP"), DataGridViewComboBoxColumn).DisplayMember = "paneldescription"
                            DirectCast(DataGridView1.Columns("ColumnCP"), DataGridViewComboBoxColumn).ValueMember = "id"
                            DataGridView1.Columns("ColumnCP").DataPropertyName = "cp"

                            DataGridView1.DataSource = SPBS
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

                            ComboBox1.DataSource = CategoryBS
                            ComboBox1.DisplayMember = "category"
                            ComboBox1.ValueMember = "supplierscategoryid"
                            ComboBox1.DataBindings.Add("SelectedValue", SPBS, "supplierscategoryid", True, DataSourceUpdateMode.OnPropertyChanged)

                            ComboBox3.DataSource = VendorBS
                            ComboBox3.DisplayMember = "description"
                            ComboBox3.ValueMember = "vendorcode"
                            ComboBox3.DataBindings.Add("SelectedValue", SPBS, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged)

                            ComboBox2.DataSource = FPBS
                            ComboBox2.DisplayMember = "paneldescription"
                            ComboBox2.ValueMember = "id"
                            ComboBox2.DataBindings.Add("SelectedValue", SPBS, "fp", True, DataSourceUpdateMode.OnPropertyChanged)

                            ComboBox4.DataSource = CPBS
                            ComboBox4.DisplayMember = "paneldescription"
                            ComboBox4.ValueMember = "id"
                            ComboBox4.DataBindings.Add("SelectedValue", SPBS, "cp", True, DataSourceUpdateMode.OnPropertyChanged)
                            TextBox1.DataBindings.Add("Text", SPBS, "rank", True, DataSourceUpdateMode.OnPropertyChanged, "")

                            ToolStripComboBox1.SelectedIndex = 1

                            If IsNothing(SPBS.Current) Then
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

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles SPBS.ListChanged
        ComboBox1.Enabled = Not IsNothing(SPBS.Current)
        ComboBox2.Enabled = Not IsNothing(SPBS.Current)
        ComboBox3.Enabled = Not IsNothing(SPBS.Current)
        ComboBox4.Enabled = Not IsNothing(SPBS.Current)
        TextBox1.Enabled = Not IsNothing(SPBS.Current)
        Button8.Enabled = Not IsNothing(SPBS.Current)

    End Sub



    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = SPBS.AddNew()
        drv.Row.BeginEdit()
        ' drv.Row.Item("validfrom") = Date.Today
        'ComboBox1.SelectedIndex = -1
        'ComboBox2.SelectedIndex = -1
        'ComboBox3.SelectedIndex = -1
        'ComboBox4.SelectedIndex = -1
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click

        If Me.validate Then
            Try
                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                SPBS.EndEdit()
                Dim ds2 As DataSet
                ds2 = DS.GetChanges

                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                    If Not DbAdapter1.SupplierPanelTx(Me, mye) Then
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
        Else
          
        End If
        DataGridView1.Invalidate()
    End Sub



    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()
        'SPBS.EndEdit()

        'Dim ds2 = DS.GetChanges
        'If Not IsNothing(ds2) Then
        '    For Each drv As DataRow In ds2.Tables(0).Rows
        '        If drv.RowState = DataRowState.Modified Or drv.RowState = DataRowState.Added Then
        '            If Not validaterow(drv) Then
        '                myret = False
        '            End If
        '        End If
        '    Next
        '    'Get The error 
        '    If Not myret Then
        '        DS.Merge(ds2)
        '    End If
        'End If
        'Return myret
        For Each drv As DataRowView In SPBS.List
            If drv.Row.RowState = DataRowState.Modified Or drv.Row.RowState = DataRowState.Added Then
                If Not validaterow(drv.Row) Then
                    myret = False
                End If
            End If
        Next
        Return myret


    End Function

    Private Function validaterow(ByRef drv As DataRow) As Boolean
        Dim myret As Boolean = True
        Dim sb As New StringBuilder
        Try
            If IsDBNull(drv.Item("vendorcode")) Then
                myret = False
                sb.Append("Vendor cannot be blank")
            End If
            If IsDBNull(drv.Item("supplierscategoryid")) Then
                myret = False
                If sb.Length > 0 Then
                    sb.Append(", ")
                End If
                sb.Append("Category cannot be blank.")
            End If
            drv.RowError = sb.ToString
        Catch ex As Exception

        End Try

        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(SPBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    SPBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged

    End Sub
    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted, ComboBox2.SelectionChangeCommitted, ComboBox3.SelectionChangeCommitted, ComboBox4.SelectionChangeCommitted
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)
        ''Dim bindings = myobj.DataBindings.Cast(Of Binding)().Where(Function(x) x.PropertyName = "SelectedItem" AndAlso x.DataSourceUpdateMode = DataSourceUpdateMode.OnPropertyChanged)

        '1. Force the Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        Dim drv As DataRowView = SPBS.Current

        Select Case myobj.Name
            Case "ComboBox1"
                Dim myselected As DataRowView = ComboBox1.SelectedItem
                'drv.Row.Item("supplierscategoryid") = myselected.Row.Item("supplierscategoryid")
                drv.Row.Item("category") = myselected.Row.Item("category")
            Case "ComboBox2"
                Dim myselected As DataRowView = ComboBox2.SelectedItem
                'drv.Row.Item("fp") = myselected.Row.Item("id")
                drv.Row.Item("fppanel") = myselected.Row.Item("paneldescription")
            Case "ComboBox3"
                Dim myselected As DataRowView = ComboBox3.SelectedItem
                'drv.Row.Item("vendorcode") = myselected.Row.Item("vendorcode")
                drv.Row.Item("vendorname") = myselected.Row.Item("vendorname")
            Case "ComboBox4"
                Dim myselected As DataRowView = ComboBox4.SelectedItem
                'drv.Row.Item("cp") = myselected.Row.Item("id")
                drv.Row.Item("cppanel") = myselected.Row.Item("paneldescription")
        End Select
        '2. Force the BindingSource to commit the value 
        SPBS.EndEdit()

        DataGridView1.Invalidate()
    End Sub


    'Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted, ComboBox2.SelectionChangeCommitted, ComboBox3.SelectionChangeCommitted, ComboBox4.SelectionChangeCommitted
    '    Dim myobj As ComboBox = DirectCast(sender, ComboBox)
    '    Dim drv As DataRowView = SPBS.Current
    '    'SPBS.RaiseListChangedEvents = True
    '    Select Case myobj.Name
    '        Case "ComboBox1"
    '            'drv.Row.Item("category") = myobj.Text
    '            Dim myselected As DataRowView = ComboBox1.SelectedItem
    '            drv.Row.Item("supplierscategoryid") = myselected.Row.Item("supplierscategoryid")
    '            drv.Row.Item("category") = myselected.Row.Item("category")

    '        Case "ComboBox2"
    '            'drv.Row.Item("fppanel") = myobj.Text
    '            Dim myselected As DataRowView = ComboBox2.SelectedItem
    '            drv.Row.Item("fp") = myselected.Row.Item("id")
    '            drv.Row.Item("fppanel") = myselected.Row.Item("paneldescription")
    '        Case "ComboBox3"
    '            Dim myselected As DataRowView = ComboBox3.SelectedItem
    '            drv.Row.Item("vendorcode") = myselected.Row.Item("vendorcode")
    '            drv.Row.Item("vendorname") = myselected.Row.Item("vendorname")
    '        Case "ComboBox4"
    '            'drv.Row.Item("cppanel") = myobj.Text
    '            Dim myselected As DataRowView = ComboBox4.SelectedItem
    '            drv.Row.Item("cp") = myselected.Row.Item("id")
    '            drv.Row.Item("cppanel") = myselected.Row.Item("paneldescription")
    '    End Select

    '    DataGridView1.Invalidate()
    'End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim myobj As Button = CType(sender, Button)
        'SPBS.RaiseListChangedEvents = True
        Select Case myobj.Name
            Case "Button8"
                Dim myform = New FormHelper(VendorBSHelper)
                myform.DataGridView1.Columns(0).DataPropertyName = "description"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = VendorBSHelper.Current
                    Dim mydrv As DataRowView = SPBS.Current
                    'mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
                    'mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
                    'mydrv.Row.Item("shortname") = drv.Row.Item("shortname")

                    ''Need bellow code to sync with combobox bindingsource
                    Dim myposition = VendorBS.Find("vendorcode", drv.Row.Item("vendorcode"))
                    VendorBS.Position = myposition
                End If

        End Select
        DataGridView1.Invalidate()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        FormPanelHistory.Show()
    End Sub


    Private Sub ToolStripTextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ToolStripTextBox1.KeyUp
        Select Case e.KeyValue
            Case Keys.Enter
                MessageBox.Show("Enter")
        End Select
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        ' SPBS.EndEdit()
        SPBS.Filter = ""
        If ToolStripTextBox1.Text <> "" Then
            Select Case ToolStripComboBox1.SelectedIndex
                Case 0
                    If Not IsNumeric(ToolStripTextBox1.Text) Then
                        ToolStripTextBox1.Select()
                        SendKeys.Send("{BACKSPACE}")

                    End If
            End Select
            SPBS.Filter = myFields(ToolStripComboBox1.SelectedIndex).ToString & " like '%" & sender.ToString.Replace("'", "''") & "%'"
        End If
    End Sub


    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        Dim newcolumn As New DataGridViewColumn
        newcolumn = Nothing
        Dim oricolumn As New DataGridViewColumn
        oricolumn = DataGridView1.Columns(e.ColumnIndex)
        Select Case e.ColumnIndex
            Case 0
                DataGridView1.Columns(1).HeaderCell.SortGlyphDirection = SortOrder.None
                DataGridView1.Columns(2).HeaderCell.SortGlyphDirection = SortOrder.None

            Case 1
                newcolumn = DataGridView1.Columns("Column2")              
            Case 2
                newcolumn = DataGridView1.Columns("Column3")
            Case 3
                newcolumn = DataGridView1.Columns("Column4")
            Case 4
                newcolumn = DataGridView1.Columns("Column5")
            Case 5
                newcolumn = DataGridView1.Columns("Column6")
                'Case Else
                '    newcolumn = Nothing

        End Select
        If Not IsNothing(newcolumn) Then
            Dim oldcolumn As DataGridViewColumn = DataGridView1.SortedColumn
            Dim direction As ListSortDirection
            If oldcolumn IsNot Nothing Then
                If oldcolumn Is newcolumn AndAlso DataGridView1.SortOrder = SortOrder.Ascending Then
                    direction = ListSortDirection.Descending
                Else
                    ' Sort a new column and remove the old SortGlyph.
                    direction = ListSortDirection.Ascending
                    oldcolumn.HeaderCell.SortGlyphDirection = SortOrder.None
                End If
            Else
                direction = ListSortDirection.Ascending
            End If
            DataGridView1.Sort(newcolumn, direction)
            If direction = ListSortDirection.Ascending Then
                oricolumn.HeaderCell.SortGlyphDirection = SortOrder.Ascending
            Else
                oricolumn.HeaderCell.SortGlyphDirection = SortOrder.Descending
            End If
        End If
       
    End Sub



End Class