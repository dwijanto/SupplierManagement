Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass

Public Class FormMasterGroup
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents GroupBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder

    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Public Function getGroupSBUBS() As BindingSource
        Dim sqlstr As String = String.Format("select null::character varying as groupname, 0 as groupid union all (select groupname,groupid from doc.groupauth where newvendorcreation order by groupname);")
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

        'sb.Append("select p.paramname as status, p.ivalue as statusid from doc.paramdt p" &
        '          " left join doc.paramhd ph on ph.paramname = 'vendorstatus'" &
        '          " order by p.ivalue ")
        sb.Append("select groupname,groupid,newvendorcreation from doc.groupauth;")


        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "Group"

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
                            GroupBS = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("groupid")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns("groupid").AutoIncrement = True
                            DS.Tables(0).Columns("groupid").AutoIncrementSeed = 0
                            DS.Tables(0).Columns("groupid").AutoIncrementStep = -1
                            DS.Tables(0).TableName = "Group"

                            GroupBS.DataSource = DS.Tables(0)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = GroupBS
                            DataGridView1.RowTemplate.Height = 22

                            TextBox1.DataBindings.Clear()
                            CheckBox1.DataBindings.Clear()

                            TextBox1.DataBindings.Add(New Binding("Text", GroupBS, "groupname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            CheckBox1.DataBindings.Add(New Binding("checked", GroupBS, "newvendorcreation", True, DataSourceUpdateMode.OnPropertyChanged, ""))


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

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles GroupBS.ListChanged
        TextBox1.Enabled = Not IsNothing(GroupBS.Current)

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = GroupBS.AddNew()
        drv.Row.Item("newvendorcreation") = False
        drv.EndEdit()
    End Sub
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        GroupBS.EndEdit()
        If Me.validate Then
            Try
                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                Dim ds2 As DataSet
                ds2 = DS.GetChanges

                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)

                    If Not DbAdapter1.GroupTx(Me, mye) Then
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

        For Each drv As DataRowView In GroupBS.List
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
        If IsDBNull(drv.Row.Item("groupname")) Then
            myret = False
            sb.Append("Group Name cannot be blank")
        End If

        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(GroupBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    GroupBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GroupBS.CancelEdit()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        If Not IsNothing(GroupBS) Then
            Dim drv As DataRowView = GroupBS.Current
            Dim myform As New FormGroupVendor
            myform.groupid = drv.Row.Item("groupid")
            myform.groupname = drv.Row.Item("groupname")
            If myform.groupid <> 0 Then
                myform.Show()
            Else
                MessageBox.Show("Please commit the current record.")
            End If

        End If
        
    End Sub

    Private Sub ToolStripButton2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(GroupBS) Then
            Dim drv As DataRowView = GroupBS.Current
            Dim myform As New FormGroupUser
            myform.groupid = drv.Row.Item("groupid")
            myform.groupname = drv.Row.Item("groupname")
            myform.Show()
        End If
    End Sub
End Class