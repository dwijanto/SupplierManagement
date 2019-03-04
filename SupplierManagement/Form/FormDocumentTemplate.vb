Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass

Public Class FormDocumentTemplate
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myThreadDownload As New System.Threading.Thread(AddressOf DoDownload)
    Dim WithEvents DocTypeBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim SelectedFolder As String = String.Empty
    Dim myFields As String() = {"doctypename"}
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

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
        sb.Append("select false::boolean as download, doctypename,groupid,template,id from doc.doctype where not template isnull order by doctypename ;")
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
                            ToolStripComboBox1.ComboBox.SelectedIndex = 0
                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = DocTypeBS
                            'DataGridView1.RowTemplate.Height = 22

                           
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



    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Me.Validate()
        If Not myThreadDownload.IsAlive Then
            'get Folder

            Dim myfolder = New SaveFileDialog
            myfolder.FileName = "SaveFile"
            If myfolder.ShowDialog = Windows.Forms.DialogResult.OK Then
                ToolStripStatusLabel1.Text = ""
                SelectedFolder = IO.Path.GetDirectoryName(myfolder.FileName)
                myThreadDownload = New Thread(AddressOf DoDownload)
                myThreadDownload.Start()
            End If

        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoDownload()
        Dim i As Integer = 0
        For Each drv As DataRowView In DocTypeBS.List
            drv.Row.RowError = ""
            If drv.Row.Item("download") Then
                i = i + 1
                ProgressReport(1, String.Format("Downloading ::{0} of {1} {2}", i, DocTypeBS.Count, drv.Row.Item("template")))
                'Dim filesource As String = String.Format("\\172.22.10.44\SharedFolder\PriceCmmf\New\documents\{0}{1}", drv.Row.Item("id"), drv.Row.Item("docext"))
                Dim filesource As String = String.Format("{0}\{1}", HelperClass1.template, drv.Row.Item("template"))
                Dim FileTarget As String = String.Format("{0}\{1}", SelectedFolder, drv.Row.Item("template"))
                Try
                    FileIO.FileSystem.CopyFile(filesource, FileTarget, True)
                Catch ex As Exception
                    drv.Row.RowError = ex.Message
                End Try

            End If
        Next
        ProgressReport(1, "Done. Please check your folder ::" & SelectedFolder)
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If Not IsNothing(DocTypeBS) Then
            For Each drv As DataRowView In DocTypeBS.List
                drv.Row.Item("download") = CheckBox1.Checked
            Next
            DocTypeBS.EndEdit()
        End If
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged, ToolStripComboBox1.SelectedIndexChanged
        'PHBS.EndEdit()
        DocTypeBS.Filter = ""
        If ToolStripTextBox1.Text <> "" And ToolStripComboBox1.SelectedIndex <> -1 Then
            DocTypeBS.Filter = myFields(ToolStripComboBox1.SelectedIndex).ToString & " like '%" & sender.ToString.Replace("'", "''") & "%'"
        End If
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        loaddata()
    End Sub
End Class