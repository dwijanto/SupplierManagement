Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class FormPanelStatus
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents PSBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim PanelStatuslist As String() = {"FP", "CP"}

    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()

        sb.Append("select panelstatus::character varying,paneldescription,id from doc.panelstatus order by panelstatus,paneldescription")


        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "PanelStatus"

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
                            PSBS = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("id")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns("id").AutoIncrement = True
                            DS.Tables(0).Columns("id").AutoIncrementSeed = 0
                            DS.Tables(0).Columns("id").AutoIncrementStep = -1
                            DS.Tables(0).TableName = "PanelStatus"

                            PSBS.DataSource = DS.Tables(0)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = PSBS
                            DataGridView1.RowTemplate.Height = 22
                            ComboBox1.DataSource = PanelStatuslist

                            ComboBox1.DataBindings.Clear()
                            TextBox2.DataBindings.Clear()


                            ComboBox1.DataBindings.Add("Text", PSBS, "panelstatus", True, DataSourceUpdateMode.OnPropertyChanged, "")
                            TextBox2.DataBindings.Add(New Binding("Text", PSBS, "paneldescription", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            If IsNothing(PSBS.Current) Then
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

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles PSBS.ListChanged
        ComboBox1.Enabled = Not IsNothing(PSBS.Current)
        TextBox2.Enabled = Not IsNothing(PSBS.Current)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = PSBS.AddNew()
        drv.Row.BeginEdit()
        drv.Row.Item("panelstatus") = ComboBox1.Text
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        PSBS.EndEdit()
        If Me.validate Then
            Try
                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                Dim ds2 As DataSet
                ds2 = DS.GetChanges

                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                    If Not DbAdapter1.PanelStatusTx(Me, mye) Then
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



    'Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
    '    PSBS.EndEdit()
    '    If Me.validate Then
    '        Try
    '            'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
    '            Dim ds2 As DataSet
    '            ds2 = DS.GetChanges

    '            If Not IsNothing(ds2) Then
    '                Dim mymessage As String = String.Empty
    '                Dim ra As Integer
    '                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
    '                If DbAdapter1.PanelStatusTx(Me, mye) Then
    '                    'delete original Dataset (DS) for those table having added record -> Merged with modified Dataset (DS2)
    '                    'For update record, no need to delete the original dataset (DS) because the id is the same. 
    '                    'Why need to delete the added one, because when we create new record, the id started with 0,-1,-2 and so on.
    '                    'when we update to database, we put the real id from database.
    '                    'so we have different value id for DS and DS2. if we do merged without deleting the original one, we will have 2 records.
    '                    Dim modifiedRows = From row In DS.Tables(0)
    '                        Where row.RowState = DataRowState.Added
    '                    For Each row In modifiedRows.ToArray
    '                        row.Delete()
    '                    Next

    '                Else
    '                    MessageBox.Show(mye.message)
    '                    Exit Sub
    '                End If
    '                DS.Merge(ds2)
    '                DS.AcceptChanges()
    '                DataGridView1.Invalidate()
    '                MessageBox.Show("Saved.")
    '            End If
    '        Catch ex As Exception
    '            MessageBox.Show(" Error:: " & ex.Message)
    '        End Try
    '    End If
    '    DataGridView1.Invalidate()
    'End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()

        For Each drv As DataRowView In PSBS.List
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
        If IsDBNull(drv.Row.Item("panelstatus")) Then
            myret = False
            sb.Append("Panel Status cannot be blank")
        End If
        If IsDBNull(drv.Row.Item("paneldescription")) Then
            myret = False
            If sb.Length > 0 Then
                sb.Append(", ")
            End If
            sb.Append("Type cannot be blank.")
        End If
        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(PSBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    PSBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        DataGridView1.Invalidate()
    End Sub


End Class