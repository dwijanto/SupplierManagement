Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Imports System.ComponentModel

Public Class FormVendorStatus

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents VSBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim bsVendorName As BindingSource
    Dim bsVendorNameHelper As BindingSource
    Dim myDict As Dictionary(Of String, Integer)

    Public Enum SupEnum
        <EnumDescription("Active")> Active = 1
        <EnumDescription("2nd Tier")> SecondTier = 2
        <EnumDescription("No Business")> No_Business = 3
        <EnumDescription("EOL")> EOL = 4
    End Enum

    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()

        sb.Append("select vs.vendorcode,v.vendorname::text,status,case status " &
                            " when 1 then 'Active' " &
                            " when 2 then '2nd Tier' " &
                            " when 3 then 'No Business' " &
                            " when 4 then 'EOL'  end as statusname," &
                            " rank from doc.vendorstatus vs left join vendor v on v.vendorcode = vs.vendorcode order by vendorcode;")
        sb.Append("select null as vendorcode,''::text as description,''::text as vendorname union all (select vendorcode, vendorcode::text || ' - ' || vendorname::text as description,vendorname::text from vendor order by vendorname);")

        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "VendorStatus"

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

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("vendorcode")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).TableName = "VendorStatus"

                            VSBS.DataSource = DS.Tables(0)
                            bsVendorName.DataSource = New DataView(DS.Tables(1))
                            bsVendorNameHelper.DataSource = New DataView(DS.Tables(1))

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = VSBS
                            DataGridView1.RowTemplate.Height = 22

                            'myDict = New Dictionary(Of String, Integer)
                            'myDict.Add("Active", 1)
                            'myDict.Add("EOL", 2)
                            'myDict.Add("No Business", 3)

                            ComboBox1.DataBindings.Clear()
                            ComboBox2.DataBindings.Clear()
                            'ComboBox3.DataBindings.Clear()


                            ComboBox1.DataSource = bsVendorName
                            ComboBox1.DisplayMember = "description"
                            ComboBox1.ValueMember = "vendorcode"
                            ComboBox1.DataBindings.Add("SelectedValue", VSBS, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged)


                            'ComboBox2.DataSource = New BindingSource(GetType(SupEnum).EnumToDictionary, Nothing)

                            ComboBox2.DataSource = New BindingSource(GetType(SupEnum).EnumToDictionary, Nothing)
                            ComboBox2.DisplayMember = "Key"
                            ComboBox2.ValueMember = "Value"
                            ComboBox2.DataBindings.Add("SelectedValue", VSBS, "status", True, DataSourceUpdateMode.OnPropertyChanged)

                            TextBox1.DataBindings.Clear()

                            ComboBox3.DataSource = New BindingSource(myDict, Nothing)
                            ComboBox3.DisplayMember = "Key"
                            ComboBox3.ValueMember = "Value"
                            ComboBox3.DataBindings.Add("SelectedValue", VSBS, "status", True, DataSourceUpdateMode.OnPropertyChanged)

                            TextBox1.DataBindings.Add(New Binding("Text", VSBS, "rank", True, DataSourceUpdateMode.OnPropertyChanged, ""))

                            If IsNothing(VSBS.Current) Then
                                ComboBox1.SelectedIndex = -1
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
    Private Sub ProgressReportOld(ByVal id As Integer, ByVal message As String)
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

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("vendorcode")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).TableName = "VendorStatus"

                            VSBS.DataSource = DS.Tables(0)
                            bsVendorName.DataSource = DS.Tables(1)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = VSBS
                            DataGridView1.RowTemplate.Height = 22

                            myDict = New Dictionary(Of String, Integer)
                            myDict.Add("Active", 1)
                            myDict.Add("EOL", 2)
                            myDict.Add("No Business", 3)

                            ComboBox1.DataBindings.Clear()
                            ComboBox2.DataBindings.Clear()
                            ComboBox3.DataBindings.Clear()


                            ComboBox1.DataSource = bsVendorName
                            ComboBox1.DisplayMember = "description"
                            ComboBox1.ValueMember = "vendorcode"
                            ComboBox1.DataBindings.Add("SelectedValue", VSBS, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged)


                            ComboBox2.DataSource = New BindingSource(myDict, Nothing)
                            ComboBox2.DisplayMember = "Key"
                            ComboBox2.ValueMember = "Value"
                            ComboBox2.DataBindings.Add("SelectedValue", VSBS, "status", True, DataSourceUpdateMode.OnPropertyChanged)

                            TextBox1.DataBindings.Clear()

                            'ComboBox3.DataSource = New BindingSource(GetType(SupEnum).EnumToDictionary, Nothing)
                            'ComboBox3.DisplayMember = "Key"
                            'ComboBox3.ValueMember = "Value"
                            'ComboBox3.DataBindings.Add("SelectedValue", VSBS, "status", True, DataSourceUpdateMode.OnPropertyChanged)

                            TextBox1.DataBindings.Add(New Binding("Text", VSBS, "rank", True, DataSourceUpdateMode.OnPropertyChanged, ""))

                            If IsNothing(VSBS.Current) Then
                                ComboBox1.SelectedIndex = -1


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

    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        '
        If Not IsNothing(VSBS.Current) Then
            Dim myselected As DataRowView = ComboBox1.SelectedItem            
            Dim drv As DataRowView = VSBS.Current
            drv.Row.Item("vendorcode") = myselected.Row.Item("vendorcode")
            drv.Row.Item("vendorname") = myselected.Row.Item("vendorname")

            'bsVendorName.Position = ComboBox1.SelectedIndex
            'drv.Row.Item("vendorname") = myselected.Row.Item("vendorname")
            ' drv.Row.Item("vendorcode") = myselected.Row.Item("vendorcode")

            'drv.Row.Item("vendorcode") = myselected.Row.Item("vendorcode") 'Sometime the combobox doesn't work correctly. just to make sure we assign with the correct one.

            ' drv.Row.Item("vendorcode") = myselected.Row.Item("vendorcode") -- No need to assign, until this point the vendorcode still blank, it'll automatically populated after propertychanged
        End If
        DataGridView1.Invalidate()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim myobj As Button = CType(sender, Button)
        Select Case myobj.Name
            Case "Button8"
                Dim myform = New FormHelper(bsVendorNamehelper)
                myform.DataGridView1.Columns(0).DataPropertyName = "description"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = bsVendorNameHelper.Current
                    Dim mydrv As DataRowView = VSBS.Current
                    mydrv.BeginEdit()
                    mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
                    mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")

                    'Need bellow code to sync with combobox
                    Dim myposition = bsVendorName.Find("vendorcode", drv.Row.Item("vendorcode"))
                    bsVendorName.Position = myposition

                End If

        End Select
        DataGridView1.Invalidate()
    End Sub


    Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectionChangeCommitted
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)
        ''Dim bindings = myobj.DataBindings.Cast(Of Binding)().Where(Function(x) x.PropertyName = "SelectedItem" AndAlso x.DataSourceUpdateMode = DataSourceUpdateMode.OnPropertyChanged)

        '1. Force the Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        If Not IsNothing(VSBS.Current) Then
            Dim myselected As System.Collections.Generic.KeyValuePair(Of String, Integer) = ComboBox2.SelectedItem
            Dim drv As DataRowView = VSBS.Current
            'drv.Row.BeginEdit()
            'drv.Row.Item("status") = myselected.Value
            drv.Row.Item("statusname") = myselected.Key
            VSBS.EndEdit()
        End If
        DataGridView1.Invalidate()
    End Sub
End Class