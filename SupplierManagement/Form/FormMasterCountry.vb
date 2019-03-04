Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass

Delegate Function DBAdapterTable(ByVal obj As Object, ByVal arguments As Object)

Public Class FormMasterCountry



    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    'Dim myDBadapterTableFunction As New DBAdapterTable(AddressOf myfunction)

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)

    Dim WithEvents SCBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder

    Private myfunct As DBAdapterTable
    Public Property TableName As String = "Country"
    Public Property Sqlstr As String = "select p.paramname as countryname, p.paramdtid as id from doc.paramdt p" &
                  " left join doc.paramhd ph on ph.paramhdid = p.paramhdid" &
                  " where ph.paramname = 'country'" &
                  "order by p.paramname"

    Public Property ColumnId As String = "id"
    Public Property ColumnName As String = "countryname"
    'Private _lableColumnName As String

    Public Property LabelColumnName As String = "Country Name"

    'Public Property labelColumnName As String
    '    Get
    '        Return _lableColumnName
    '    End Get
    '    Set(ByVal value As String)
    '        _lableColumnName = value
    '        Label1.Text = value
    '    End Set
    'End Property


    Public Sub New(ByVal myfunct As Object)

        ' This call is required by the designer.
        InitializeComponent()
        Me.myfunct = myfunct

        'Label1.Text = LabelColumnName
        ' Add any initialization after the InitializeComponent() call.

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

       
        sb.Append(Sqlstr)


        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = TableName

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
                            SCBS = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns(ColumnId)
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns(ColumnId).AutoIncrement = True
                            DS.Tables(0).Columns(ColumnId).AutoIncrementSeed = 0
                            DS.Tables(0).Columns(ColumnId).AutoIncrementStep = -1
                            DS.Tables(0).TableName = TableName

                            SCBS.DataSource = DS.Tables(0)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.Columns(0).DataPropertyName = ColumnName
                            DataGridView1.Columns(0).HeaderText = LabelColumnName
                     

                            DataGridView1.DataSource = SCBS
                            DataGridView1.RowTemplate.Height = 22

                            TextBox1.DataBindings.Clear()

                            Label1.Text = LabelColumnName
                            TextBox1.DataBindings.Add(New Binding("Text", SCBS, ColumnName, True, DataSourceUpdateMode.OnPropertyChanged, ""))


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


    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles SCBS.ListChanged
        TextBox1.Enabled = Not IsNothing(SCBS.Current)

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        SCBS.AddNew()
    End Sub
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        SCBS.EndEdit()
        If Me.validate Then
            Try
                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                Dim ds2 As DataSet
                ds2 = DS.GetChanges

                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)

                    'If Not DbAdapter1.CountryTx(Me, mye) Then
                    '    MessageBox.Show(mye.message)
                    '    Exit Sub
                    'End If

                    If Not myfunct.Invoke(Me, mye) Then
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

        For Each drv As DataRowView In SCBS.List
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
        If IsDBNull(drv.Row.Item(ColumnName)) Then
            myret = False
            sb.Append(String.Format("{0} cannot be blank", LabelColumnName))
        End If

        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(SCBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    SCBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SCBS.CancelEdit()
    End Sub



    Public Function myfunction(ByVal obj As Object, ByVal obj2 As Object)
        Throw New NotImplementedException
    End Function


    Private Sub ToolStripButton4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        loaddata()
    End Sub
End Class