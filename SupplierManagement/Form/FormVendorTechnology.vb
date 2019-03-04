Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass

Public Class FormVendorTechnology
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents VTBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim bsTechnology As BindingSource   
    Dim myVendorcode As Long
    Dim myvendorname As String
    Dim mylineno As Integer
    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Label1.Text = String.Format("{0} - {1}", myVendorcode, myvendorname)
        loaddata()
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()

        sb.Append(String.Format("select vt.*,t.technologyname from doc.vendortechnology vt left join doc.technology t on t.id = vt.technologyid where vendorcode = {0} order by lineno;", myVendorcode))
        sb.Append("select * from doc.technology order by technologyname;")        


        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                DS.Tables(0).TableName = "VendorTechnology"
                DS.Tables(1).TableName = "Technology"
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
                            VTBS = New BindingSource
                            bsTechnology = New BindingSource
                          

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("id")
                            DS.Tables(0).Columns("id").AutoIncrement = True
                            DS.Tables(0).Columns("id").AutoIncrementSeed = 0
                            DS.Tables(0).Columns("id").AutoIncrementStep = -1
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).TableName = "VendorTechnology"
                       
                            VTBS.DataSource = DS.Tables(0)
                            bsTechnology.DataSource = New DataView(DS.Tables(1))
                          
                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = VTBS
                            'DataGridView1.RowTemplate.Height = 22

                            ComboBox1.DataBindings.Clear()

                            ComboBox1.DataSource = bsTechnology
                            ComboBox1.DisplayMember = "technologyname"
                            ComboBox1.ValueMember = "id"
                            ComboBox1.DataBindings.Add("SelectedValue", VTBS, "technologyid", True, DataSourceUpdateMode.OnPropertyChanged)


                            TextBox1.DataBindings.Clear()

                            TextBox1.DataBindings.Add(New Binding("Text", VTBS, "lineno", True, DataSourceUpdateMode.OnPropertyChanged, ""))

                            If IsNothing(VTBS.Current) Then
                                ComboBox1.SelectedIndex = -1
                            Else
                                mylineno = DS.Tables(0).Rows(DS.Tables(0).Rows.Count).Item("lineno")
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

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles VTBS.ListChanged
        ComboBox1.Enabled = Not IsNothing(VTBS.Current)
        TextBox1.Enabled = Not IsNothing(VTBS.Current)

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = VTBS.AddNew()
        mylineno = mylineno + 10
        drv.Row.Item("lineno") = mylineno
        drv.Row.Item("vendorcode") = myVendorcode
        TextBox1.Text = mylineno
        drv.Row.BeginEdit()
    End Sub


    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Try
            VTBS.EndEdit()
            If Me.validate Then
                Try
                    'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                    Dim ds2 As DataSet
                    ds2 = DS.GetChanges

                    If Not IsNothing(ds2) Then
                        Dim mymessage As String = String.Empty
                        Dim ra As Integer
                        Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                        If Not DbAdapter1.VendorTechnologyTx(Me, mye) Then
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

        For Each drv As DataRowView In VTBS.List
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
        If IsDBNull(drv.Row.Item("lineno")) Then
            myret = False
            If sb.Length > 0 Then
                sb.Append(", ")
            End If
            sb.Append("Line No cannot be blank.")
        End If
        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(VTBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    VTBS.RemoveAt(drv.Index)
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

    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        'Dim myobj As ComboBox = DirectCast(sender, ComboBox)
        ''1. Force Combobox to commit the value 
        'For Each binding As Binding In myobj.DataBindings
        '    binding.WriteValue()
        '    binding.ReadValue()
        'Next

        If Not IsNothing(VTBS.Current) Then
            Dim myselected1 As DataRowView = ComboBox1.SelectedItem


            Dim drv As DataRowView = VTBS.Current
            Try

                drv.Row.BeginEdit()
                drv.Row.Item("technologyid") = myselected1.Row.Item("id")
                drv.Row.Item("technologyname") = myselected1.Row.Item("technologyname")

                VTBS.EndEdit()
            Catch ex As Exception
                ComboBox1.SelectedValue = drv.Row.Item("vendorcode", DataRowVersion.Original)
                drv.Row.CancelEdit()
                MessageBox.Show(ex.Message)
            End Try

        End If
        DataGridView1.Invalidate()
    End Sub

  


    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message.ToString)
    End Sub

    Public Sub New(ByVal vendorcode As Long, ByVal vendorname As String)

        ' This call is required by the designer.
        InitializeComponent()
        myVendorcode = vendorcode
        myvendorname = vendorname
        ' Add any initialization after the InitializeComponent() call.

    End Sub




End Class