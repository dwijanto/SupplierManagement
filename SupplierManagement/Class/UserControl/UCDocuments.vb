Imports System.Threading
Imports SupplierManagement.PublicClass

Public Class UCDocuments
    Dim myform As Object
    Private WithEvents BS As BindingSource
    Private DataTypeBS As BindingSource
    Private drv As DataRowView
    Dim myThreadDownload As New System.Threading.Thread(AddressOf DoDownload)
    Private SelectedFolder As String

    Dim FileSourceFullPath As String

    Public Sub DisableDataGridViewMenu()
        DataGridView1.ContextMenuStrip = Nothing
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub HistoryMode()
        If Not HelperClass1.UserInfo.IsAdmin Then
            DataGridView1.ContextMenuStrip = Nothing
        End If
    End Sub


    Public Sub BindingControl(ByRef myform As Object, ByRef bs As BindingSource, ByVal DataTypeBS As BindingSource, ByVal FileSourceFullPath As String)
        Me.myform = myform
        Me.BS = bs
        Me.FileSourceFullPath = FileSourceFullPath
        Me.DataTypeBS = DataTypeBS
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = bs

        ComboBox1.DataSource = DataTypeBS
        ComboBox1.DisplayMember = "doctype"
        ComboBox1.ValueMember = "id"

        ComboBox1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", bs, "doctypeid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox1.DataBindings.Add(New Binding("Text", bs, "docname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("Text", bs, "remarks", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        Check_Enabled_Object()
    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        BS.EndEdit()

        For Each drv As DataRowView In BS.List
            drv.Row.RowError = ""
            If IsDBNull(drv.Row.Item("docfullname")) Then
                drv.Row.RowError = "File name cannot be blank."
                myret = False
                'Else
                '    If (drv.Row.Item("docname") = "") Then
                '        drv.Row.RowError = "File name cannot be blank."
                '        myret = False
                '    End If
            End If
        Next

        Return myret
    End Function

    Private Sub NewRecordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewRecordToolStripMenuItem.Click
        Dim drv As DataRowView = BS.AddNew
        ComboBox1.SelectedIndex = 0
        drv.Item("doctypeid") = DirectCast(ComboBox1.SelectedItem, DataRowView).Row.Item("id")

        Combobox1SelectionChangeCommitted()
        drv.EndEdit()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If Not IsNothing(BS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    BS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim openfiledialog1 As New OpenFileDialog
        If openfiledialog1.ShowDialog = DialogResult.OK Then
            Dim mydrv As DataRowView = BS.Current
            mydrv.Row.Item("docfullname") = IO.Path.GetFullPath(openfiledialog1.FileName)
            mydrv.Row.Item("docname") = IO.Path.GetFileName(openfiledialog1.FileName)
            mydrv.Row.Item("docext") = IO.Path.GetExtension(openfiledialog1.FileName)
            TextBox1.Text = IO.Path.GetFileName(openfiledialog1.FileName)
            DataGridView1.Invalidate()
        End If
    End Sub

    Private Sub BS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles BS.ListChanged
        Check_Enabled_Object()
    End Sub


    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        'Me.validate()
        Combobox1SelectionChangeCommitted()
    End Sub

    Private Sub Check_Enabled_Object()
        If Not IsNothing(BS) Then
            Button2.Enabled = Not IsNothing(BS.Current)
            ComboBox1.Enabled = Not IsNothing(BS.Current)
            TextBox2.Enabled = Not IsNothing(BS.Current)
            If IsNothing(BS.Current) Then
                ComboBox1.SelectedIndex = -1
            End If
        End If
    End Sub

    Private Sub Combobox1SelectionChangeCommitted()
        Dim drv = ComboBox1.SelectedItem
        Dim mydrv As DataRowView = BS.Current
        mydrv.Row.Item("doctype") = drv.row.item("doctype")
        mydrv.Row.Item("doctypeid") = drv.row.item("id")
        mydrv.EndEdit()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Validate()
        If Not myThreadDownload.IsAlive Then
            'get Folder
            Dim myfolder = New SaveFileDialog
            myfolder.FileName = "SaveFile"
            If myfolder.ShowDialog = Windows.Forms.DialogResult.OK Then
                myform.ToolStripStatusLabel1.Text = ""
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
        For Each drv As DataRowView In BS.List
            If drv.Row.Item("download") Then
                i = i + 1
                ProgressReport(1, String.Format("Downloading ::{0} of {1} {2}", i, BS.Count, drv.Row.Item("docname")))
                Dim filesource As String = String.Format("{0}\{1}{2}", FileSourceFullPath, drv.Row.Item("id"), drv.Row.Item("docext"))
                Dim FileTarget As String = String.Format("{0}\{1}", SelectedFolder, drv.Row.Item("docname"))
                Try
                    FileIO.FileSystem.CopyFile(filesource, FileTarget, True)
                Catch ex As Exception
                End Try
            End If
        Next
        ProgressReport(1, "Done. Please check your folder ::" & SelectedFolder)
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
                        myform.ToolStripStatusLabel1.Text = message
                    Case 2
                        myform.ToolStripStatusLabel1.Text = message
                    Case 4

                    Case 5
                        myform.ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        myform.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        HelperClass1.previewdoc(BS, FileSourceFullPath)
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If Not IsNothing(BS) Then
            For Each drv As DataRowView In BS.List
                drv.Row.Item("download") = CheckBox1.Checked
            Next
            BS.EndEdit()
        End If
    End Sub


End Class
