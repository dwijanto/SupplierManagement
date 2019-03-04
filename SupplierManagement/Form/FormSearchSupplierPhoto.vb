Imports System.Threading
Imports System.Text
Public Class FormSearchSupplierPhoto
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myadapter As New VendorAdapter


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
        ProgressReport(6, "Marquee")
        If myadapter.LoadSupplierSearchPhoto() Then
            ProgressReport(4, "Init Data")
        End If
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

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = myadapter.BS
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

    Private Sub FormSearchSupplierPhoto_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripButton2.Click
        loaddata()
    End Sub

    Private Sub ToolStripTextBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        Dim obj As ToolStripTextBox = DirectCast(sender, ToolStripTextBox)
        Dim myfilter As String
        If obj.Text = "" Then
            'reset filter
            myfilter = ""
        Else
            myfilter = String.Format("[vendorcode] like '*{0}*' or [vendorname] like '*{0}*' or [shortname] like '*{0}*'", obj.Text)

        End If
        myadapter.BS.Filter = myfilter
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click, DataGridView1.DoubleClick
        Dim drv As DataRowView = myadapter.BS.Current
        If IsDBNull(drv.Row.Item("shortname")) Then
            MessageBox.Show("Shortname is blank. cannot add photo.")
            Exit Sub
        End If
        Dim shortname As String = drv.Row.Item("shortname")

        Dim myform As New FormSupplierPhotos(drv)
        myform.Show()

    End Sub


End Class