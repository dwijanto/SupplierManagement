Imports System.Threading
Imports System.Text

Public Class FormToolingSupplier
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myController As ToolingSupplierAdapter
    Dim dbadapter1 As DbAdapter = DbAdapter.getInstance
    Dim drv As DataRowView = Nothing
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        RemoveHandler DialogVendorCurr.RefreshDataGridView, AddressOf RefreshDataGridView
        AddHandler DialogVendorCurr.RefreshDataGridView, AddressOf RefreshDataGridView
    End Sub


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loadData()
    End Sub

    Private Sub loadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        myController = New ToolingSupplierAdapter

        Try
            ProgressReport(1, "Loading..")
            If myController.LoadData() Then
                ProgressReport(4, "Init Data")
            End If

            ProgressReport(1, "Done.")
        Catch ex As Exception

            ProgressReport(1, ex.Message)
        End Try
    End Sub


    Public Sub showTx(ByVal tx As TxEnum)
        If Not myThread.IsAlive Then
            Select Case tx
                Case TxEnum.NewRecord
                    drv = myController.GetNewRecord
                Case TxEnum.UpdateRecord
                    drv = myController.GetCurrentRecord
            End Select

            Me.drv.BeginEdit()
            'Dim myform = New DialogVendorCurr(drv, VendorBS, VendorHelperBS)
            'myform.ShowDialog()
        End If

    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myController.BS
            End Select
        End If
    End Sub


    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        Dim obj As ToolStripTextBox = DirectCast(sender, ToolStripTextBox)
        myController.ApplyFilter = obj.Text
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        loadData()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        showTx(TxEnum.NewRecord)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        showTx(TxEnum.UpdateRecord)
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(myController.GetCurrentRecord) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myController.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ToolStripButton2.PerformClick()
    End Sub

    Private Sub RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.Validate()
        myController.save()
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        'Start Thread
        If Not myThread.IsAlive Then
            'Get file
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                myThread = New Thread(AddressOf doImport)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub doImport()
        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String


        Dim myret As Boolean = False
        
        Dim errormsg As String = String.Empty

        Dim Filename As String = String.Empty
        Dim mySB As New StringBuilder
        Try
            Filename = OpenFileDialog1.FileName
            Using objTFParser = New FileIO.TextFieldParser(Filename)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    ProgressReport(1, "Read Data")

                    Do Until .EndOfData
                        myrecord = .ReadFields

                        If count >= 1 Then
                            Dim mymodel As New ToolingSupplierModel With {.toolingsupplierid = Validstr(myrecord(1)),
                                                                      .toolingsuppliername = Validstr(myrecord(2)),
                                                                      .address = Validstr(myrecord(10)),
                                                                      .deliveryaddress = Validstr(myrecord(10)),
                                                                      .tel = Validstr(myrecord(4)),
                                                                      .fax = Validstr(myrecord(6))
                                                                      }
                            mySB.Append(mymodel.toolingsupplierid & vbTab &
                                             mymodel.toolingsuppliername & vbTab &
                                              mymodel.address & vbTab &
                                              mymodel.deliveryaddress & vbTab &
                                              mymodel.tel & vbTab &
                                              mymodel.fax & vbCrLf)

                        End If
                        count += 1
                    Loop

                    If mySB.Length > 0 Then
                        'Delete, copy
                        Dim sqlstr = "delete from doc.toolingsupplier;copy doc.toolingsupplier(toolingsupplierid,toolingsuppliername,address,deliveryaddress,tel,fax) from stdin with null as 'Null';"
                        errormsg = dbadapter1.copy(sqlstr, mySB.ToString, myret)
                        If Not myret Then
                            ProgressReport(1, errormsg)
                            Exit Sub
                        End If
                        myController.LoadData()
                        ProgressReport(4, "Init Data")
                    End If

                End With
            End Using

            ProgressReport(1, "Add Records Done.")
            ProgressReport(3, "Set Continuous Again")
            myret = True
            Me.DialogResult = Windows.Forms.DialogResult.OK
        Catch ex As Exception

            errormsg = ex.Message

            ProgressReport(1, errormsg)
            Me.DialogResult = Windows.Forms.DialogResult.None
            ProgressReport(3, "Set Continuous Again")
        Finally
           
        End Try
    End Sub

    Private Function Validstr(ByVal myrecord As String) As String
        If myrecord.Length = 0 Then
            myrecord = "Null"
        End If
        Return myrecord
    End Function

End Class