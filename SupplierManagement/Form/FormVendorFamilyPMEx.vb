﻿Imports System.Threading

Public Class FormVendorFamilyPMEx
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myController As VendorFamilyPMExAdapter
    Dim VendorController As VendorTxAdapter
    Dim FamilyPMController As FamilyPMAdapter
    Dim PMController As PMAdapter
    Dim drv As DataRowView = Nothing
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        RemoveHandler DialogVendorFamilyPMEx.RefreshDataGridView, AddressOf RefreshDataGridView
        AddHandler DialogVendorFamilyPMEx.RefreshDataGridView, AddressOf RefreshDataGridView
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
        myController = New VendorFamilyPMExAdapter
        VendorController = New VendorTxAdapter
        FamilyPMController = New FamilyPMAdapter
        PMController = New PMAdapter
        Try
            ProgressReport(1, "Loading..")
            If myController.LoadData() Then
                ProgressReport(4, "Init Data")
            End If
            VendorController.loaddata()
            FamilyPMController.LoadData()
            PMController.loaddata()
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
            Dim VendorBS = VendorController.GetBindingSourceVendorFamily
            Dim VendorHelperBS = VendorController.GetBindingSourceVendorFamily
            Dim FamilyPMBS = FamilyPMController.GetBindingSource
            Dim FamilyPMHelperBS = FamilyPMController.GetBindingSource
            Dim PMBS = PMController.GetBindingSource
            Dim PMHelperBS = PMController.GetBindingSource
            Me.drv.BeginEdit()
            Dim myform = New DialogVendorFamilyPMEx(drv, VendorBS, VendorHelperBS, FamilyPMBS, FamilyPMHelperBS, PMBS, PMHelperBS)
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
End Class