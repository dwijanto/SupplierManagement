Imports System.Threading
Public Class FormFamilyPM
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myAdapter As FamilyPMAdapter
    Dim PMAdapter1 As PMAdapter
    Dim FamilyAdapter1 As FamilyAdapter

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        RemoveHandler DialogFamilyPM.RefreshDataGridView, AddressOf DialogUserAU_RefreshDataGridView
        AddHandler DialogFamilyPM.RefreshDataGridView, AddressOf DialogUserAU_RefreshDataGridView
    End Sub

    Private Sub FormUserN_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
        myAdapter = New FamilyPMAdapter
        PMAdapter1 = New PMAdapter
        FamilyAdapter1 = New FamilyAdapter
        Try
            ProgressReport(1, "Loading..")
            If myAdapter.loaddata() Then
                ProgressReport(4, "Init Data")
            End If
            PMAdapter1.loaddata()
            FamilyAdapter1.loaddata()
            
            ProgressReport(1, "Done.")
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
    End Sub

    Public Sub showTx(ByVal tx As TxEnum)
        If Not myThread.IsAlive Then
            Dim drv As DataRowView = Nothing
            Select Case tx
                Case TxEnum.NewRecord
                    drv = myAdapter.GetNewRecord
                Case TxEnum.UpdateRecord
                    drv = myAdapter.GetCurrentRecord
            End Select
            'Dim GSMDT As DataTable = myAdapter.GetTableGSM
            Dim PMBS As New BindingSource
            PMBS.DataSource = PMAdapter1.GetBindingSource

            Dim FamilyBS As New BindingSource
            FamilyBS.DataSource = FamilyAdapter1.GetBindingSource

            Dim FamilyHelperBS As New BindingSource
            FamilyHelperBS.DataSource = FamilyAdapter1.GetBindingSource

            Dim myform = New DialogFamilyPM(drv, FamilyBS, FamilyHelperBS, PMBS)
            myform.Show()
            myform.Focus()

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
                    DataGridView1.DataSource = myAdapter.BS               
            End Select
        End If
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        Dim obj As ToolStripTextBox = DirectCast(sender, ToolStripTextBox)
        myAdapter.ApplyFilter = obj.Text
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
        If Not IsNothing(myAdapter.GetCurrentRecord) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myAdapter.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ToolStripButton2.PerformClick()
    End Sub

    Private Sub DialogUserAU_RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        myAdapter.save()
    End Sub

    Private Sub ToolStripTextBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.Click

    End Sub
End Class