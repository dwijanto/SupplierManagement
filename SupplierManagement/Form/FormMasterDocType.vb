Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class FormMasterDocType
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents DocTypeBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder

    Dim DocTypeGroupList As List(Of DocTypeGroup)
    Dim myFields As String() = {"doctypename"}
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        DocTypeGroupList = New List(Of DocTypeGroup)
        DocTypeGroupList.Add(New DocTypeGroup(1, "Document"))
        DocTypeGroupList.Add(New DocTypeGroup(2, "Contract"))

    End Sub
    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        loaddata()
    End Sub
    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        DS = New DataSet
        Dim mymessage As String = String.Empty
        sb.Clear()
        sb.Append("select d.doctypename,d.groupid,d.id,dr.reminder,dr.reminderupdatedate,dr.reminderupdateby,dt.template,dt.templateupdatedate,dt.templateupdateby,case when d.groupid = 1 then 'Document' when d.groupid = 2 then 'Contract' end as groupname" &
                  " from doc.doctype d" &
                  " left join doc.doctypetemplate dt on dt.id = d.id" &
                  " left join doc.doctypereminder dr on dr.id = d.id" &
                  " order by doctypename;")
        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                DS.Tables(0).TableName = "DocType"
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
    Sub DoWork1()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        DS = New DataSet
        Dim mymessage As String = String.Empty
        sb.Clear()
        sb.Append("select doctypename,groupid,template,id,templateupdatedate,templateupdateby,case when groupid = 1 then 'Document' when groupid = 2 then 'Contract' end as groupname from doc.doctype order by doctypename;")
        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                DS.Tables(0).TableName = "DocType"
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
                            DocTypeBS = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("id")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns("id").AutoIncrement = True
                            DS.Tables(0).Columns("id").AutoIncrementSeed = 0
                            DS.Tables(0).Columns("id").AutoIncrementStep = -1                            

                            DocTypeBS.DataSource = DS.Tables(0)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = DocTypeBS
                            'DataGridView1.RowTemplate.Height = 22

                            ComboBox1.DataBindings.Clear()
                            TextBox1.DataBindings.Clear()
                            TextBox2.DataBindings.Clear()
                            TextBox3.DataBindings.Clear()

                            ComboBox1.DataSource = DocTypeGroupList 'PanelStatuslist
                            ComboBox1.DisplayMember = "GroupName"
                            ComboBox1.ValueMember = "GroupId"
                            ComboBox1.DataBindings.Add(New Binding("SelectedValue", DocTypeBS, "groupid", True, DataSourceUpdateMode.OnPropertyChanged, ""))

                            TextBox1.DataBindings.Add(New Binding("Text", DocTypeBS, "template", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox3.DataBindings.Add(New Binding("Text", DocTypeBS, "reminder", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox2.DataBindings.Add(New Binding("Text", DocTypeBS, "doctypename", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            If IsNothing(DocTypeBS.Current) Then
                                ComboBox1.SelectedIndex = -1
                            End If
                            ToolStripComboBox1.SelectedIndex = 0
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

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles DocTypeBS.ListChanged
        ComboBox1.Enabled = Not IsNothing(DocTypeBS.Current)
        TextBox1.Enabled = Not IsNothing(DocTypeBS.Current)
        TextBox2.Enabled = Not IsNothing(DocTypeBS.Current)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged, TextBox1.TextChanged
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = DocTypeBS.AddNew()
        drv.Row.Item("groupid") = 1
        drv.Row.BeginEdit()        
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        DocTypeBS.EndEdit()
        If Me.validate Then
            Try
                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                Dim ds2 As DataSet
                ds2 = DS.GetChanges

                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                    If Not DbAdapter1.MasterDocTypeTx(Me, mye) Then
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

        For Each drv As DataRowView In DocTypeBS.List
            If drv.Row.RowState = DataRowState.Modified Or drv.Row.RowState = DataRowState.Added Then
                drv.Row.Item("templateupdateby") = HelperClass1.UserInfo.userid
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
        If IsDBNull(drv.Row.Item("doctypename")) Then
            myret = False
            sb.Append("Doc Type Name cannot be blank")
        End If       
        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(DocTypeBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    DocTypeBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)
        ''Dim bindings = myobj.DataBindings.Cast(Of Binding)().Where(Function(x) x.PropertyName = "SelectedItem" AndAlso x.DataSourceUpdateMode = DataSourceUpdateMode.OnPropertyChanged)

        '1. Force the Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        If Not IsNothing(DocTypeBS.Current) Then
            Dim myselected As SupplierManagement.DocTypeGroup = ComboBox1.SelectedItem
            Dim drv As DataRowView = DocTypeBS.Current            
            drv.Row.Item("groupname") = myselected.GroupName

        End If
        DataGridView1.Invalidate()

    End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged, ToolStripComboBox1.SelectedIndexChanged
        'PHBS.EndEdit()
        DocTypeBS.Filter = ""
        If ToolStripTextBox1.Text <> "" And ToolStripComboBox1.SelectedIndex <> -1 Then
            Select Case ToolStripComboBox1.SelectedIndex
                Case 1
                    If Not IsNumeric(ToolStripTextBox1.Text) Then
                        ToolStripButton1.Select()
                        SendKeys.Send("{BACKSPACE}")
                    End If
            End Select
            DocTypeBS.Filter = myFields(ToolStripComboBox1.SelectedIndex).ToString & " like '%" & sender.ToString.Replace("'", "''") & "%'"
        End If
    End Sub

   

End Class

Class DocTypeGroup
    Property GroupId As Integer
    Property GroupName As String

    Public Sub New(ByVal GroupId As Integer, ByVal GroupName As String)
        Me.GroupId = GroupId
        Me.GroupName = GroupName
    End Sub

End Class
