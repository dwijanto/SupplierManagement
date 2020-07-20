Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Imports System.ComponentModel

Public Class FormVendorStatus001
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents VSBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim bsVendorName As BindingSource
    Dim bsVendorNameHelper As BindingSource
    Dim bsStatus As BindingSource
    Dim bsProductType As BindingSource
    Dim myDict As Dictionary(Of String, Integer)
    Dim myFields As String() = {"vendorcode", "vendorname", "shortname", "statusname"}

    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Public Function getStatusBS() As BindingSource
        Dim sqlstr As String = String.Format("select p.paramname as status, p.ivalue as statusid from doc.paramdt p" &
                   " left join doc.paramhd ph on ph.paramhdid = p.paramhdid" &
                   " where ph.paramname = 'vendorstatus'" &
                   " order by p.ivalue;")
        Dim DS As New DataSet
        Dim bs As New BindingSource
        If DbAdapter1.TbgetDataSet(sqlstr, DS) Then
            bs.DataSource = DS.Tables(0)
        End If
        Return bs
    End Function

    Public Function getProductTypeBS() As BindingSource
        Dim sqlstr As String = String.Format("select ''::text as producttype,0::integer as producttypeid union all (select p.paramname as producttype, p.ivalue as producttypeid from doc.paramdt p" &
                   " left join doc.paramhd ph on ph.paramhdid = p.paramhdid" &
                   " where ph.paramname = 'producttype'" &
                   " order by p.ivalue);")
        Dim DS As New DataSet
        Dim bs As New BindingSource
        If DbAdapter1.TbgetDataSet(sqlstr, DS) Then
            bs.DataSource = DS.Tables(0)
        End If
        Return bs
    End Function

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        sb.Append("with ppt as (select p.paramname as producttype, p.ivalue as producttypeid from doc.paramdt p " &
                  "             left join doc.paramhd ph on ph.paramhdid = p.paramhdid where ph.paramname = 'producttype'), " &
                  "     pvs as (select p.paramname as statusname, p.ivalue as status from doc.paramdt p " &
                  "             left join doc.paramhd ph on ph.paramhdid = p.paramhdid where  ph.paramname = 'vendorstatus')" &
                  " select vs.vendorcode::text,v.vendorname::text,v.shortname::text,vs.status,pvs.statusname,ppt.producttype, rank,vs.latestupdate,vs.usermodified, vs.producttypeid" &
                  " from doc.vendorstatus vs " &
                  " left join vendor v on v.vendorcode = vs.vendorcode " &
                  " left join pvs on pvs.status = vs.status" &
                  " left join ppt on ppt.producttypeid = vs.producttypeid" &
                  " order by vendorcode;")
        sb.Append("select null as vendorcode,''::text as description,''::text as vendorname union all (select vendorcode, vendorcode::text || ' - ' || vendorname::text as description,vendorname::text from vendor order by vendorname);")
        sb.Append("select p.paramname as status, p.ivalue as statusid from doc.paramdt p" &
                   " left join doc.paramhd ph on ph.paramhdid = p.paramhdid" &
                   " where ph.paramname = 'vendorstatus'" &
                   " order by p.ivalue;")
        sb.Append("select ''::text as producttype,null::integer as producttypeid union all (select p.paramname as producttype, p.ivalue as producttypeid from doc.paramdt p" &
                   " left join doc.paramhd ph on ph.paramhdid = p.paramhdid" &
                   " where ph.paramname = 'producttype'" &
                   " order by p.ivalue);")
        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                DS.Tables(0).TableName = "VendorStatus"
                DS.Tables(2).TableName = "Status"
                DS.Tables(3).TableName = "ProductType"
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
                            VSBS = New BindingSource
                            bsVendorName = New BindingSource
                            bsVendorNameHelper = New BindingSource
                            bsStatus = New BindingSource
                            bsProductType = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("vendorcode")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).TableName = "VendorStatus"

                            VSBS.DataSource = DS.Tables(0)
                            bsVendorName.DataSource = New DataView(DS.Tables(1))
                            bsVendorNameHelper.DataSource = New DataView(DS.Tables(1))
                            bsStatus.DataSource = DS.Tables(2)
                            bsProductType.DataSource = DS.Tables(3)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = VSBS
                            DataGridView1.RowTemplate.Height = 22


                            ComboBox1.DataBindings.Clear()
                            ComboBox2.DataBindings.Clear()
                            ComboBox3.DataBindings.Clear()

                            ComboBox1.DataSource = bsVendorName
                            ComboBox1.DisplayMember = "description"
                            ComboBox1.ValueMember = "vendorcode"
                            ComboBox1.DataBindings.Add("SelectedValue", VSBS, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged)


                            ComboBox2.DataSource = bsStatus
                            ComboBox2.DisplayMember = "status"
                            ComboBox2.ValueMember = "statusid"
                            ComboBox2.DataBindings.Add("SelectedValue", VSBS, "status", True, DataSourceUpdateMode.OnPropertyChanged)

                            ComboBox3.DataSource = bsProductType
                            ComboBox3.DisplayMember = "producttype"
                            ComboBox3.ValueMember = "producttypeid"
                            ComboBox3.DataBindings.Add("SelectedValue", VSBS, "producttypeid", True, DataSourceUpdateMode.OnPropertyChanged)

                            ToolStripComboBox1.SelectedIndex = 1
                            TextBox1.DataBindings.Clear()

                            TextBox1.DataBindings.Add(New Binding("Text", VSBS, "rank", True, DataSourceUpdateMode.OnPropertyChanged, ""))

                            If IsNothing(VSBS.Current) Then
                                ComboBox1.SelectedIndex = -1
                                ComboBox2.SelectedIndex = -1
                                ComboBox3.SelectedIndex = -1
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

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles VSBS.ListChanged
        ComboBox1.Enabled = Not IsNothing(VSBS.Current)
        ComboBox2.Enabled = Not IsNothing(VSBS.Current)
        ComboBox3.Enabled = Not IsNothing(VSBS.Current)
        TextBox1.Enabled = Not IsNothing(VSBS.Current)
        Button8.Enabled = Not IsNothing(VSBS.Current)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = VSBS.AddNew()
        drv.Row.BeginEdit()       
    End Sub


    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Try
            VSBS.EndEdit()
            If Me.validate Then
                Try
                    'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                    Dim ds2 As DataSet
                    ds2 = DS.GetChanges

                    If Not IsNothing(ds2) Then
                        Dim mymessage As String = String.Empty
                        Dim ra As Integer
                        Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                        If Not DbAdapter1.VendorStatusTx(Me, mye) Then
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

        For Each drv As DataRowView In VSBS.List
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
            sb.Append("Vendor Code cannot be blank")
        End If
        If IsDBNull(drv.Row.Item("status")) Then
            myret = False
            If sb.Length > 0 Then
                sb.Append(", ")
            End If
            sb.Append("Status cannot be blank.")
        End If
        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(VSBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    VSBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    'Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
    '    Dim myobj As ComboBox = DirectCast(sender, ComboBox)
    '    ''Dim bindings = myobj.DataBindings.Cast(Of Binding)().Where(Function(x) x.PropertyName = "SelectedItem" AndAlso x.DataSourceUpdateMode = DataSourceUpdateMode.OnPropertyChanged)

    '    '1. Force the Combobox to commit the value 
    '    For Each binding As Binding In myobj.DataBindings
    '        binding.WriteValue()
    '        binding.ReadValue()
    '    Next

    '    If Not IsNothing(VSBS.Current) Then
    '        Dim myselected As DataRowView = ComboBox1.SelectedItem
    '        Dim drv As DataRowView = VSBS.Current
    '        'drv.Row.Item("vendorcode") = myselected.Row.Item("vendorcode")
    '        drv.Row.Item("vendorname") = myselected.Row.Item("vendorname")

    '    End If
    '    DataGridView1.Invalidate()
    'End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim myobj As Button = CType(sender, Button)
        Try
            Select Case myobj.Name
                Case "Button8"
                    Dim myform = New FormHelper(bsVendorNameHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "description"
                    If myform.ShowDialog = DialogResult.OK Then
                        Dim drv As DataRowView = bsVendorNameHelper.Current
                        Dim mydrv As DataRowView = VSBS.Current
                        mydrv.BeginEdit()
                        mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
                        mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
                        mydrv.EndEdit()
                        'Need bellow code to sync with combobox
                        Dim myposition = bsVendorName.Find("vendorcode", drv.Row.Item("vendorcode"))
                        bsVendorName.Position = myposition
                    End If
            End Select
        Catch ex As Exception

            MessageBox.Show(ex.Message)
        End Try
        
        DataGridView1.Invalidate()
    End Sub


    Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted, ComboBox2.SelectionChangeCommitted, ComboBox3.SelectionChangeCommitted
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)
        '1. Force Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        If Not IsNothing(VSBS.Current) Then
            Dim myselected1 As DataRowView = ComboBox1.SelectedItem
            Dim myselected2 As DataRowView = ComboBox2.SelectedItem
            Dim myselected3 As DataRowView = ComboBox3.SelectedItem
            Dim drv As DataRowView = VSBS.Current
            Try
               
                drv.Row.BeginEdit()
                drv.Row.Item("vendorname") = myselected1.Row.Item("vendorname")
                drv.Row.Item("statusname") = myselected2.Item("status")
                drv.Row.Item("producttype") = myselected3.Item("producttype")
                VSBS.EndEdit()
            Catch ex As Exception
                ComboBox1.SelectedValue = drv.Row.Item("vendorcode", DataRowVersion.Original)
                drv.Row.CancelEdit()
                MessageBox.Show(ex.Message)
            End Try

        End If
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged, ToolStripComboBox1.SelectedIndexChanged
        VSBS.Filter = ""
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
            VSBS.Filter = myFields(ToolStripComboBox1.SelectedIndex).ToString & " like '%" & sender.ToString.Replace("'", "''") & "%'"
            ToolStripStatusLabel1.Text = "Record Count " & VSBS.Count
        End If
    End Sub


    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message.ToString)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim mydrv As DataRowView = VSBS.Current
        If Not IsNothing(mydrv) Then
            If IsDBNull(mydrv.Row.Item("producttype")) Then
                MessageBox.Show("Technology only for ""CP"" Vendor.")
                Exit Sub
            End If
            If Not mydrv.Row.Item("producttype").ToString.Contains("CP") Then
                MessageBox.Show("Technology only for ""CP"" Vendor.")
                Exit Sub
            End If

            Dim myform As New FormVendorTechnology(mydrv.Row.Item("vendorcode"), mydrv.Row.Item("vendorname"))
            myform.Show()
        End If
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Dim mydrv As DataRowView = VSBS.Current
        Dim myform As New FormVendorFamilySubFamilyVC(mydrv.Row.Item("vendorcode"), mydrv.Row.Item("vendorname"))
        myform.Show()
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        Dim mydrv As DataRowView = VSBS.Current
        Dim myform As New FormVendorIndirectFamily(mydrv.Row.Item("vendorcode"), mydrv.Row.Item("vendorname"))
        myform.Show()
    End Sub
End Class