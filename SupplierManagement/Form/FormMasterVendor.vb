Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class FormMasterVendor

    Dim SelectedFolder As String
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)

    Dim WithEvents VBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim bsSSM As BindingSource
    Dim bsPM As BindingSource
    Dim myDict As Dictionary(Of String, Integer)
    Dim myFields As String() = {"vendorcode", "vendorname", "shortname", "ssm", "pm"}


    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        sb.Append(" select v.vendorcode::text,v.vendorname::text,v.shortname::text,ssm.officersebname as ssm,v.ssmidpl as ssmid,v.ssmeffectivedate,pm.officersebname as pm,v.pmid as pmid,v.pmeffectivedate,o.officername::text,v.shortname2" &
                  " from vendor v" &
                  " left join officerseb ssm on ssm.ofsebid = v.ssmidpl" &
                  " left join officerseb pm on pm.ofsebid = v.pmid" &
                  " left join officer o on o.officerid = v.officerid" &
                  " order by vendorcode;")
        'sb.Append("select null as ssm,null as id union all (Select o.officersebname,o.ofsebid from officerseb o  where parent  = 0 and isactive Order by o.officersebname);")
        'sb.Append("select null as pm,null as id union all (Select o.officersebname,o.ofsebid from officerseb o  where levelid  = 4 and isactive Order by o.officersebname);")
        'sb.Append("select null as ssm,null as id union all (select o.officersebname,o.ofsebid from doc.user u left join officerseb o on o.userid = u.userid  where parent = 0 and  o.isactive order by officersebname);")
        sb.Append("select null as ssm,null as id union all (select mu.username as officersebname,o.ofsebid from doc.user u left join officerseb o on o.userid = u.userid  left join masteruser mu on mu.id = o.muid where parent = 0 and  o.isactive order by officersebname);")
        'sb.Append("select null as pm,null as id union all (select o.officersebname,o.ofsebid from doc.user u left join officerseb o on o.userid = u.userid  where levelid = 4 and  o.isactive order by officersebname);")
        sb.Append("select null as pm,null as id union all (select mu.username as officersebname,o.ofsebid from doc.user u left join officerseb o on o.userid = u.userid left join masteruser mu on mu.id = o.muid  where levelid = 4 and  o.isactive order by officersebname);")
        'sb.Append("select o.officersebname,o.ofsebid from doc.user u left join officerseb o on o.userid = u.userid  where parent = 0 and  o.isactive order by officersebname;")
        sb.Append("select mu.username as officersebname,o.ofsebid from doc.user u left join officerseb o on o.userid = u.userid  " &
                  " left join masteruser mu on mu.id = o.muid where parent = 0 and  o.isactive order by officersebname;")
        'sb.Append("select o.officersebname,o.ofsebid from doc.user u left join officerseb o on o.userid = u.userid  where levelid = 4 and  o.isactive order by officersebname;")
        sb.Append("select mu.username as officersebname,o.ofsebid from doc.user u left join officerseb o on o.userid = u.userid  " &
                  " left join masteruser mu on mu.id = o.muid where levelid = 4 and  o.isactive order by officersebname;")
        sb.Append("select * from vendorspmpm where id = 0;")
        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                DS.Tables(0).TableName = "Vendor"
                DS.Tables(1).TableName = "SSM"
                DS.Tables(2).TableName = "PM"
                DS.Tables(3).TableName = "SSM-Master"
                DS.Tables(4).TableName = "PM-Master"
                DS.Tables(5).TableName = "Vendor_SPM_PM_History"
                
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
                            VBS = New BindingSource
                            bsSSM = New BindingSource
                            bsPM = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("vendorcode")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).TableName = "VendorStatus"

                            Dim pk3(0) As DataColumn
                            pk3(0) = DS.Tables(3).Columns("officersebname")
                            DS.Tables(3).PrimaryKey = pk3


                            Dim pk4(0) As DataColumn
                            pk4(0) = DS.Tables(4).Columns("officersebname")
                            DS.Tables(4).PrimaryKey = pk4


                            VBS.DataSource = DS.Tables(0)                            
                            bsSSM.DataSource = DS.Tables(1)
                            bsPM.DataSource = DS.Tables(2)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = VBS
                            DataGridView1.RowTemplate.Height = 22

                            TextBox1.DataBindings.Clear()
                            TextBox2.DataBindings.Clear()
                            TextBox3.DataBindings.Clear()
                            TextBox4.DataBindings.Clear()
                            TextBox5.DataBindings.Clear()
                            ComboBox1.DataBindings.Clear()
                            ComboBox2.DataBindings.Clear()
                            DateTimePicker1.DataBindings.Clear()
                            DateTimePicker2.DataBindings.Clear()
                            

                            ComboBox1.DataSource = bsSSM
                            ComboBox1.DisplayMember = "ssm"
                            ComboBox1.ValueMember = "id"
                            ComboBox1.DataBindings.Add("SelectedValue", VBS, "ssmid", True, DataSourceUpdateMode.OnPropertyChanged)

                            ComboBox2.DataSource = bsPM
                            ComboBox2.DisplayMember = "pm"
                            ComboBox2.ValueMember = "id"
                            ComboBox2.DataBindings.Add("SelectedValue", VBS, "pmid", True, DataSourceUpdateMode.OnPropertyChanged)

                            ToolStripComboBox1.SelectedIndex = 1


                            TextBox1.DataBindings.Add(New Binding("Text", VBS, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox2.DataBindings.Add(New Binding("Text", VBS, "vendorname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox3.DataBindings.Add(New Binding("Text", VBS, "shortname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox4.DataBindings.Add(New Binding("Text", VBS, "officername", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox5.DataBindings.Add(New Binding("Text", VBS, "shortname2", True, DataSourceUpdateMode.OnPropertyChanged, ""))

                            DateTimePicker1.DataBindings.Add(New Binding("Text", VBS, "ssmeffectivedate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            DateTimePicker2.DataBindings.Add(New Binding("Text", VBS, "pmeffectivedate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            'If IsNothing(VBS.Current) Then
                            '    ComboBox1.SelectedIndex = -1
                            '    ComboBox2.SelectedIndex = -1
                            'End If

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

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles VBS.ListChanged
        ComboBox1.Enabled = Not IsNothing(VBS.Current)
        ComboBox2.Enabled = Not IsNothing(VBS.Current)
        TextBox1.Enabled = Not IsNothing(VBS.Current)
        TextBox2.Enabled = Not IsNothing(VBS.Current)
        TextBox3.Enabled = Not IsNothing(VBS.Current)

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = VBS.AddNew()
        drv.Row.BeginEdit()
    End Sub


    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Try
            'Me.validate()
            TextBox1.Focus()
            VBS.EndEdit()
            If Me.validate Then
                Try
                    'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                    Dim ds2 As DataSet
                    ds2 = DS.GetChanges

                    If Not IsNothing(ds2) Then
                        Dim mymessage As String = String.Empty
                        Dim ra As Integer
                        Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                        If Not DbAdapter1.VendorTx(Me, mye) Then
                            DS.Merge(ds2)
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
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()

        For Each drv As DataRowView In VBS.List
            If drv.Row.RowState = DataRowState.Modified Or drv.Row.RowState = DataRowState.Added Then
                If Not validaterow(drv) Then
                    myret = False
                End If
            End If
        Next
        Return myret
    End Function

    Private Function validaterow(ByVal drv As DataRowView) As Boolean
        Dim myret As Boolean = True
        Dim sb As New StringBuilder
        If IsDBNull(drv.Row.Item("vendorcode")) Then
            myret = False
            sb.Append("Vendor Code cannot be blank.")
        Else
            If Not IsNumeric(drv.Row.Item("vendorcode")) Then
                myret = False
                sb.Append("Vendor Code should be numeric.")
            End If
        End If
        If IsDBNull(drv.Row.Item("vendorname")) Then
            myret = False
            If sb.Length > 0 Then
                sb.Append(", ")
            End If
            sb.Append("Vendor Name cannot be blank.")
        End If
        'If Not IsDBNull(drv.Row.Item("ssm")) Then
        '    If IsDBNull(drv.Row.Item("ssmeffectivedate")) Then
        '        myret = False
        '        If sb.Length > 0 Then
        '            sb.Append(", ")
        '        End If
        '        sb.Append("SPM Effective Date cannot be blank.")
        '    End If

        'End If
        'If Not IsDBNull(drv.Row.Item("pm")) Then
        '    If IsDBNull(drv.Row.Item("pmeffectivedate")) Then
        '        myret = False
        '        If sb.Length > 0 Then
        '            sb.Append(", ")
        '        End If
        '        sb.Append("PM Effective Date cannot be blank.")
        '    End If

        'End If
        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(VBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    VBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub


    'Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
    '    Dim myobj As Button = CType(sender, Button)
    '    Try
    '        Select Case myobj.Name
    '            Case "Button8"
    '                Dim myform = New FormHelper(bsVendorNameHelper)
    '                myform.DataGridView1.Columns(0).DataPropertyName = "description"
    '                If myform.ShowDialog = DialogResult.OK Then
    '                    Dim drv As DataRowView = bsVendorNameHelper.Current
    '                    Dim mydrv As DataRowView = VBS.Current
    '                    mydrv.BeginEdit()
    '                    mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
    '                    mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
    '                    mydrv.EndEdit()
    '                    'Need bellow code to sync with combobox
    '                    Dim myposition = bsVendorName.Find("vendorcode", drv.Row.Item("vendorcode"))
    '                    bsVendorName.Position = myposition
    '                End If
    '        End Select
    '    Catch ex As Exception

    '        MessageBox.Show(ex.Message)
    '    End Try

    '    DataGridView1.Invalidate()
    'End Sub


    Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted, ComboBox2.SelectionChangeCommitted
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)
        '1. Force Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        If Not IsNothing(VBS.Current) Then
            'Dim myselected1 As DataRowView = ComboBox1.SelectedItem
            Dim myselected1 As DataRowView = ComboBox1.SelectedItem
            Dim myselected2 As DataRowView = ComboBox2.SelectedItem
            Dim drv As DataRowView = VBS.Current
            Try

                drv.Row.BeginEdit()
                'drv.Row.Item("vendorname") = myselected1.Row.Item("vendorname")
                drv.Row.Item("ssm") = myselected1.Item("ssm")
                drv.Row.Item("pm") = myselected2.Item("pm")
                VBS.EndEdit()
            Catch ex As Exception
                'ComboBox1.SelectedValue = drv.Row.Item("vendorcode", DataRowVersion.Original)
                drv.Row.CancelEdit()
                MessageBox.Show(ex.Message)
            End Try

        End If
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged, ToolStripComboBox1.SelectedIndexChanged
        VBS.Filter = ""
        ToolStripStatusLabel1.Text = ""
        If ToolStripTextBox1.Text <> "" Then
            Select Case ToolStripComboBox1.SelectedIndex
                Case 0
                    If Not IsNumeric(ToolStripTextBox1.Text) Then
                        ToolStripTextBox1.Select()
                        SendKeys.Send("{BACKSPACE}")
                        Exit Sub
                    End If
            End Select
            VBS.Filter = myFields(ToolStripComboBox1.SelectedIndex).ToString & " like '%" & sender.ToString.Replace("'", "''") & "%'"
            ToolStripStatusLabel1.Text = "Record Count " & VBS.Count
        End If
    End Sub


    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message.ToString)
    End Sub

    Private Sub DateTimePicker1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker1.Validated, DateTimePicker2.Validated
        Dim myobj As DateTimePicker = TryCast(sender, DateTimePicker)
        Dim drv As DataRowView = VBS.Current
        If myobj.Checked Then
            drv.Row.Item(myobj.Tag) = myobj.Value
        Else
            drv.Row.Item(myobj.Tag) = DBNull.Value
        End If
    End Sub

    'Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
    '    Dim myobj As DateTimePicker = DirectCast(sender, DateTimePicker)
    '    Debug.Print(myobj.Text)
    'End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim myform = New FormImportVendorSPMPM
        myform.ds = DS
        myform.ShowDialog()
    End Sub

   

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Dim myform = New FormHistorySPMPMAssignment
        myform.ShowDialog()
    End Sub
End Class