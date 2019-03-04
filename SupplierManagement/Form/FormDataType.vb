Imports System.Threading
Imports System.Text

Public Class FormDataType
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents BS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim myAdapter As New DataTypeAdapter
    Dim groupAdapter As New GroupDataTypeAdapter
    Dim GroupBS As BindingSource

    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Try
            GroupBS = groupAdapter.getComboBoxBS()       
            If myAdapter.loaddata Then
                ProgressReport(4, "InitData")
                ProgressReport(1, "Loading Data.Done!")
                ProgressReport(5, "Continuous")
            End If
        Catch ex As Exception
            ProgressReport(1, "Loading Data. Error::" & ex.Message)
            ProgressReport(5, "Continuous")
        End Try
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

                        ComboBox1.DataBindings.Clear()
                        ComboBox1.ValueMember = "id"
                        ComboBox1.DisplayMember = "groupname"
                        ComboBox1.DataSource = GroupBS

                        BS = New BindingSource
                        BS.DataSource = myAdapter.BS
                        DataGridView1.AutoGenerateColumns = False
                        DataGridView1.DataSource = BS
                        DataGridView1.RowTemplate.Height = 22

                        TextBox1.DataBindings.Clear()
                        TextBox2.DataBindings.Clear()
                        TextBox3.DataBindings.Clear()
                        TextBox4.DataBindings.Clear()

                        TextBox1.DataBindings.Add(New Binding("Text", BS, "datatypename", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                        TextBox2.DataBindings.Add(New Binding("Text", BS, "lineorder", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                        TextBox3.DataBindings.Add(New Binding("Text", BS, "amountlineorder", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                        TextBox4.DataBindings.Add(New Binding("Text", BS, "qtylineorder", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                        ComboBox1.DataBindings.Add(New Binding("SelectedValue", BS, "groupid", True, DataSourceUpdateMode.OnPropertyChanged, ""))

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


    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged, ComboBox1.SelectedValueChanged
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = BS.AddNew()
        ComboBox1.SelectedIndex = 0
        Dim cdrv As DataRowView = ComboBox1.SelectedItem

        drv.Item("groupname") = cdrv.Item("groupname")
        drv.Item("groupid") = cdrv.Item("id")
        TextBox1.Focus()
    End Sub
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        BS.EndEdit()
        If Me.validate Then
            myAdapter.save()
        End If
        DataGridView1.Invalidate()
    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()

        For Each drv As DataRowView In BS.List
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
        If IsDBNull(drv.Row.Item("datatypename")) Then
            myret = False
            sb.Append("Data Type Name cannot be blank.")
        End If

        If IsDBNull(drv.Row.Item("groupid")) Then
            myret = False
            sb.Append("Group Name cannot be blank.")
        End If

        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(BS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    BS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub



    Private Sub BS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles BS.ListChanged
        TextBox1.Enabled = Not IsNothing(BS.Current)
        TextBox2.Enabled = Not IsNothing(BS.Current)
        TextBox3.Enabled = Not IsNothing(BS.Current)
        TextBox4.Enabled = Not IsNothing(BS.Current)
        ComboBox1.Enabled = Not IsNothing(BS.Current)
        If IsNothing(BS.Current) Then
            ComboBox1.SelectedIndex = -1
        End If
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        BS.CancelEdit()
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        assigngroupname()
       

    End Sub

    Private Sub assigngroupname()
        Dim cdrv As DataRowView = ComboBox1.SelectedItem
        Dim drv As DataRowView = BS.Current
        drv.Item("groupname") = cdrv.Item("groupname")
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Dim myform As New FormDataTypeLabel
        myform.ShowDialog()
    End Sub
End Class