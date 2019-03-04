﻿Imports System.Threading
Public Class FormIndirectFamily
    Dim mySelectedPath As String

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myThreadImport As New ThreadStart(AddressOf DoImport)
    Dim myController As IndirectFamilyAdapter
    Dim FolderBrowserDialog1 As New FolderBrowserDialog
    Dim OpenFileDialog1 As New OpenFileDialog

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        'AddHandler FormCategoryIssue.myRefresh, AddressOf loadData
        'AddHandler FormManagementMessage.myRefresh, AddressOf loadData
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Sub DoWork()
        myController = New IndirectFamilyAdapter

        Try
            ProgressReport(1, "Loading..")
            If myController.loaddata() Then
                ProgressReport(4, "Init Data")
            End If

            ProgressReport(1, String.Format("Done. Records Count({0}).", myController.BS.Count))
        Catch ex As Exception

            ProgressReport(1, ex.Message)
        End Try
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
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
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
            End Select
        End If
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loadData()
    End Sub

    Public Sub loadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Me.loadData()
    End Sub

    Private Sub ToolStripTextBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        myController.ApplyFilter = ToolStripTextBox1.Text
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)

    End Sub

    Private Sub ToolStripButton7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        If Not myThread.IsAlive Then
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                Try
                    mySelectedPath = OpenFileDialog1.FileName
                    myThread = New System.Threading.Thread(myThreadImport)
                    myThread.SetApartmentState(ApartmentState.MTA)
                    myThread.Start()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Private Sub ToolStripButton6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        myController.Model.ExportToExcel(Me)
    End Sub

    Sub DoImport()
        myController.Model.ImportFromText(Me, mySelectedPath)
    End Sub

    Private Sub ToolStripTextBox1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.Click
        myController.ApplyFilter = ToolStripTextBox1.Text
    End Sub
End Class