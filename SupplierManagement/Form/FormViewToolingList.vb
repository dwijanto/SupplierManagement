Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class FormViewToolingList
    Dim sqlstr As String
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim DS As DataSet
    Dim BS As BindingSource
    Public Sub New(ByVal sqlstr)

        ' This call is required by the designer.
        InitializeComponent()
        Me.sqlstr = sqlstr
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub FormViewToolingList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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

    Sub DoWork()
        DS = New DataSet
        Dim mymessage As String = String.Empty
        ProgressReport(6, "InitData")
        If DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(4, "InitData")
        End If
        ProgressReport(5, "InitData")
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Try
                Me.Invoke(d, New Object() {id, message})
            Catch ex As Exception

            End Try
        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel1.Text = message
                    Case 4
                        Try                           
                            BS = New BindingSource
                            BS.DataSource = DS.Tables(0)
                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = BS
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
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
End Class