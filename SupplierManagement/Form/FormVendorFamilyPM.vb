Imports System.Threading
Imports System.Text

Public Class FormVendorFamilyPM
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myController As VendorFamilyPMAdapter
    Dim VendorController As VendorTxAdapter
    Dim FamilyPMController As FamilyPMAdapter
    Dim drv As DataRowView = Nothing
    Dim dbadapter1 = DbAdapter.getInstance
    Dim openfiledialog1 As New OpenFileDialog
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        RemoveHandler DialogVendorFamilyPM.RefreshDataGridView, AddressOf RefreshDataGridView
        AddHandler DialogVendorFamilyPM.RefreshDataGridView, AddressOf RefreshDataGridView
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
        myController = New VendorFamilyPMAdapter
        VendorController = New VendorTxAdapter
        FamilyPMController = New FamilyPMAdapter
        Try
            ProgressReport(1, "Loading..")
            If myController.LoadData() Then
                ProgressReport(4, "Init Data")
            End If
            VendorController.loaddata()
            FamilyPMController.LoadData()

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
            Dim VendorBS = VendorController.GetBindingSource
            Dim VendorHelperBS = VendorController.GetBindingSource
            Dim FamilyPMBS = FamilyPMController.GetBindingSource
            Dim FamilyPMHelperBS = FamilyPMController.GetBindingSource
            Me.drv.BeginEdit()
            Dim myform = New DialogVendorFamilyPM(drv, VendorBS, VendorHelperBS, FamilyPMBS, FamilyPMHelperBS)
            myform.ShowDialog()
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
        myController.save()
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""            
            If openfiledialog1.ShowDialog = DialogResult.OK Then                
                myThread = New Thread(AddressOf DoImport)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoImport()
        Try
            Dim mystr As New StringBuilder
            Dim myInsert As New System.Text.StringBuilder
            Dim myrecord() As String

            Using objTFParser = New FileIO.TextFieldParser(openfiledialog1.FileName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    ProgressReport(1, "Read Data")
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count >= 1 Then                            
                            myInsert.Append(myrecord(0) & vbTab &
                                            myrecord(1) & vbTab &
                                            validdate(myrecord(2)) & vbTab &
                                            validdate(myrecord(3)) & vbCrLf)
                                
                        End If
                        count += 1
                    Loop
                End With
            End Using
            'update record
            If myInsert.Length > 0 Then
                ProgressReport(1, "Start Add New Records")

                Dim sqlstr As String = "copy doc.vendorfamily(vendorcode,familyid,pmeffectivedate,spmeffectivedate) from stdin with null as 'Null';"
                Dim ra As Long = 0
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False

                Try
                    ProgressReport(1, "Add Record Please wait!")
                    errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                    If myret Then
                        ProgressReport(1, "Add Records Done.")

                    Else
                        ProgressReport(1, errmessage)

                    End If
                Catch ex As Exception
                    ProgressReport(1, ex.Message)

                End Try

            End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try

        ProgressReport(3, "Set Continuous Again")

    End Sub

    Private Function validdate(ByVal value As String) As String
        Dim myret As String
        myret = "Null"
        If value = "" Then Return myret
        myret = String.Format("'{0:yyyy-MM-dd}'", CDate(value))
        Return myret
    End Function

End Class