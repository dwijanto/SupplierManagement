Imports System.Threading
Imports System.Text

Public Class FormSupplierPhotos
    Private shortname As String
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myadapter As SupplierPhotosAdapter
    Private DRV As DataRowView
    Public Sub New(ByVal drv As DataRowView)
        ' This call is required by the designer.
        InitializeComponent()
        Me.DRV = drv
        myadapter = New SupplierPhotosAdapter(drv)
        AddHandler SupplierManagement.DialogSupplierPhotos.FinishTx, AddressOf FinishTX
        ' Add any initialization after the InitializeComponent() call.
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
        ProgressReport(6, "Marquee")
        If myadapter.LoadSupplierPhoto() Then
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

    Private Sub FormSupplierPhotos_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Validate()
        If Not IsNothing(myadapter.DS) Then
            Dim abc = myadapter.DS.GetChanges()
            If Not IsNothing(abc) Then
                Select Case MessageBox.Show("Save unsaved records?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                    Case Windows.Forms.DialogResult.Yes
                        If Me.Validate Then
                            ToolStripButton4.PerformClick()
                        Else
                            e.Cancel = True
                        End If

                    Case Windows.Forms.DialogResult.Cancel
                        e.Cancel = True
                End Select
            End If
        End If
    End Sub

    Private Sub FormSearchSupplierPhoto_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click, AddToolStripMenuItem.Click
        showTx(TxRecord.TxAdd)
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        If e.ColumnIndex = 9 Then
            If e.RowIndex > -1 Then
                Dim drv As DataRowView = myadapter.BS.Current
                Dim myfilename = String.Format("{0}\{1}{2}", myadapter.PhotoFolder, drv.Row.Item("id"), IO.Path.GetExtension(drv.Row.Item("filename")))
                Dim p As New System.Diagnostics.Process
                p.StartInfo.FileName = myfilename
                Try
                    p.Start()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

            End If
        End If
        
    End Sub



    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click, DataGridView1.CellDoubleClick, UpdateToolStripMenuItem.Click
        showTx(TxRecord.TxUpdate)
    End Sub

    Private Sub showTx(ByVal txRecord As TxRecord)
        Dim drv As DataRowView
        If txRecord = SupplierManagement.TxRecord.TxAdd Then
            drv = myadapter.BS.AddNew()
            drv.Row.Item("vendorcode") = Me.DRV.Row.Item("vendorcode")
            drv.Row.Item("createddate") = Date.Today
        Else
            drv = myadapter.BS.Current
        End If
        If Not IsNothing(drv) Then
            Dim myform As New DialogSupplierPhotos(drv)
            myform.Show()
        End If        
    End Sub

    Private Sub FinishTX()
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click, DeleteToolStripMenuItem.Click
        If Not IsNothing(myadapter.BS.Current) Then
            If MessageBox.Show("Delete selected record?", "Delete Record(s)", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myadapter.BS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If myadapter.save() Then
            'copy file
            'LINQ
            Dim query = From fullpath In myadapter.BS.List
                        Where Not IsDBNull(fullpath.item("fullpath"))

            For Each DRV As DataRowView In query
                'copy file to server            
                Dim mytarget = myadapter.PhotoFolder & "\" & DRV.Row.Item("id") & IO.Path.GetExtension(DRV.Row.Item("filename"))
                Try
                    FileIO.FileSystem.CopyFile(DRV.Row.Item("fullpath"), mytarget, True)
                Catch ex As Exception                    
                    DRV.Row.RowError = "Copy photo failed."
                End Try
            Next
            DataGridView1.Invalidate()
        End If
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        loaddata()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim myobj As ToolStripButton = DirectCast(sender, ToolStripButton)
        Dim myfilter As String = String.Empty
        If myobj.Text = "Show Less" Then
            myobj.Text = "Show All"
            myfilter = String.Format("[lineorderfp]>=1 or [lineordercp]>=1 or [lineordergeneral]>=1")
        Else
            myobj.Text = "Show Less"
            myfilter = ""
        End If
        myadapter.BS.Filter = myfilter
    End Sub

   
End Class