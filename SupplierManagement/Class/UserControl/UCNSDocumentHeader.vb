Imports System.Threading

Public Class UCNSDocumentHeader
    Dim myform As Object
    Private Const SIF As Integer = 21
    Private WithEvents BS As BindingSource 'For Attachment List
    Private WithEvents SITxBS As BindingSource
    Private DataTypeBS As BindingSource
    Private DataTypeHelperBS As BindingSource
    Private BSZetol As BindingSource
    Private bsZetolMaster As BindingSource
    Private drv As DataRowView
    Dim myThreadDownload As New System.Threading.Thread(AddressOf DoDownload)
    Private SelectedFolder As String
    Dim HelperClass1 = HelperClass.getInstance

    Dim FileSourceFullPath As String

    Dim auditbylist As String() = {"", "SGS", "SEB Asia", "Intertek", "Waived Audit"}
    Dim auditTypelist As String() = {"", "Initial", "1st follow up", "2nd follow up", "3rd follow up", "4th follow up", "5th follow up"}
    Dim auditGradelist As String() = {"", "Minor", "Major", "Critical", "Zero Tolerance NC", "Red Level", "Orange Level", "Yellow Level", "Green Level"}
    Dim overallAuditList As String() = {"", "Zero Tolerance", "Red Level", "Orange Level", "Yellow Level", "Green Level"}

    Public Sub DisableDataGridViewMenu()
        DataGridView1.ContextMenuStrip = Nothing
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub HistoryMode()
        DataGridView1.ContextMenuStrip = Nothing
    End Sub


    Public Sub BindingControl(ByRef myform As Object, ByRef bs As BindingSource, ByRef sitxbs As BindingSource, ByRef zetolbs As BindingSource, ByRef bsZetolMaster As BindingSource, ByVal DataTypeBS As BindingSource, ByVal FileSourceFullPath As String, ByVal PaymenttermBS As BindingSource, ByVal LevelBS As BindingSource)
        'Me.DataTypeHelperBS = 
        Dim dv As DataView = New DataView(DataTypeBS.DataSource)
        Me.DataTypeHelperBS = New BindingSource
        Me.DataTypeHelperBS.DataSource = dv.ToTable
        Me.myform = myform
        Me.BS = bs
        Me.SITxBS = sitxbs
        Me.FileSourceFullPath = FileSourceFullPath
        Me.DataTypeBS = DataTypeBS
        Me.BSZetol = zetolbs
        Me.bsZetolMaster = bsZetolMaster

        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = bs

        DataGridView2.AutoGenerateColumns = False
        DataGridView2.DataSource = sitxbs

        DataGridView3.AutoGenerateColumns = False
        DataGridView3.DataSource = zetolbs

        ComboBox3.DataBindings.Clear()
        ComboBox3.DataSource = DataTypeBS
        ComboBox3.DisplayMember = "doctype"
        ComboBox3.ValueMember = "id"

        ComboBox4.DataBindings.Clear()
        ComboBox4.DataSource = LevelBS
        ComboBox4.DisplayMember = "levelname"
        ComboBox4.ValueMember = "id"


        ComboBox5.DataBindings.Clear()
        ComboBox5.DataSource = PaymenttermBS
        ComboBox5.DisplayMember = "payt"
        ComboBox5.ValueMember = "paymenttermid"


        ComboBox6.DataBindings.Clear()
        ComboBox6.DataSource = auditbylist
        ComboBox6.DataBindings.Add("Text", bs, "auditby", True, DataSourceUpdateMode.OnPropertyChanged)

        ComboBox7.DataBindings.Clear()
        ComboBox7.DataSource = auditTypelist
        ComboBox7.DataBindings.Add("Text", bs, "audittype", True, DataSourceUpdateMode.OnPropertyChanged)

        ComboBox8.DataBindings.Clear()
        ComboBox8.DataSource = auditGradelist
        ComboBox8.DataBindings.Add("Text", bs, "auditgrade", True, DataSourceUpdateMode.OnPropertyChanged)

        ComboBox9.DataBindings.Clear()
        ComboBox9.DataSource = overallAuditList
        ComboBox9.DataBindings.Add("Text", bs, "overallauditresult", True, DataSourceUpdateMode.OnPropertyChanged)


        DateTimePicker2.DataBindings.Clear()
        DateTimePicker3.DataBindings.Clear()
        TextBox8.DataBindings.Clear()
        TextBox9.DataBindings.Clear()
        TextBox10.DataBindings.Clear()
        TextBox11.DataBindings.Clear()
        TextBox12.DataBindings.Clear()
        TextBox13.DataBindings.Clear()
        TextBox14.DataBindings.Clear()
        TextBox15.DataBindings.Clear()
        TextBox16.DataBindings.Clear()
        TextBox18.DataBindings.Clear()
        TextBox19.DataBindings.Clear()
        TextBox20.DataBindings.Clear()
        TextBox21.DataBindings.Clear()
        TextBox22.DataBindings.Clear()
        TextBox23.DataBindings.Clear()
        TextBox24.DataBindings.Clear()
        TextBox25.DataBindings.Clear()
        TextBox26.DataBindings.Clear()
        TextBox27.DataBindings.Clear()
        TextBox28.DataBindings.Clear()


        ComboBox3.DataBindings.Add(New Binding("SelectedValue", bs, "doctypeid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox4.DataBindings.Add(New Binding("SelectedValue", bs, "levelid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox5.DataBindings.Add(New Binding("SelectedValue", bs, "paymentcode", True, DataSourceUpdateMode.OnPropertyChanged))
        DateTimePicker2.DataBindings.Add(New Binding("Text", bs, "docdate", True, DataSourceUpdateMode.OnPropertyChanged))
        DateTimePicker3.DataBindings.Add(New Binding("Text", bs, "expireddate", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox8.DataBindings.Add(New Binding("Text", bs, "remarks", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox9.DataBindings.Add(New Binding("Text", bs, "version", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox10.DataBindings.Add(New Binding("Text", bs, "docfullname", True, DataSourceUpdateMode.OnPropertyChanged)) 'if no value means update
        TextBox11.DataBindings.Add(New Binding("Text", bs, "leadtime", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox12.DataBindings.Add(New Binding("Text", bs, "sasl", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox13.DataBindings.Add(New Binding("Text", bs, "nqsu", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox14.DataBindings.Add(New Binding("Text", bs, "projectname", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox15.DataBindings.Add(New Binding("Text", bs, "sascore", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox16.DataBindings.Add(New Binding("Text", bs, "returnrate", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox18.DataBindings.Add(New Binding("Text", bs, "sefscore", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox19.DataBindings.Add(New Binding("Text", bs, "turnovery", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox20.DataBindings.Add(New Binding("Text", bs, "turnovery1", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox21.DataBindings.Add(New Binding("Text", bs, "turnovery2", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox22.DataBindings.Add(New Binding("Text", bs, "turnovery3", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox23.DataBindings.Add(New Binding("Text", bs, "turnovery4", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox24.DataBindings.Add(New Binding("Text", bs, "ratioy", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox25.DataBindings.Add(New Binding("Text", bs, "ratioy1", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox26.DataBindings.Add(New Binding("Text", bs, "ratioy2", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox27.DataBindings.Add(New Binding("Text", bs, "ratioy3", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox28.DataBindings.Add(New Binding("Text", bs, "ratioy4", True, DataSourceUpdateMode.OnPropertyChanged))

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
        drv.Row.Item("docdate") = Date.Today

        'ComboBox1.SelectedIndex = 0
        'drv.Item("doctypeid") = DirectCast(ComboBox1.SelectedItem, DataRowView).Row.Item("id")
        'Combobox1SelectionChangeCommitted()
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

    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
    '    Dim openfiledialog1 As New OpenFileDialog
    '    If openfiledialog1.ShowDialog = DialogResult.OK Then
    '        Dim mydrv As DataRowView = BS.Current
    '        mydrv.Row.Item("docfullname") = IO.Path.GetFullPath(openfiledialog1.FileName)
    '        mydrv.Row.Item("docname") = IO.Path.GetFileName(openfiledialog1.FileName)
    '        mydrv.Row.Item("docext") = IO.Path.GetExtension(openfiledialog1.FileName)
    '        'TextBox1.Text = IO.Path.GetFileName(openfiledialog1.FileName)
    '        DataGridView1.Invalidate()
    '    End If
    'End Sub



    Private Sub BS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles BS.ListChanged
        Check_Enabled_Object()
    End Sub


    'Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
    '    'Me.validate()
    '    Combobox1SelectionChangeCommitted()
    'End Sub

    Private Sub Check_Enabled_Object()
        If Not IsNothing(BS) Then
            'Button2.Enabled = Not IsNothing(BS.Current)
            'ComboBox1.Enabled = Not IsNothing(BS.Current)
            'TextBox2.Enabled = Not IsNothing(BS.Current)
            GroupBox8.Enabled = True
            TabControl1.Enabled = True
            If IsNothing(BS.Current) Then
                ComboBox3.SelectedIndex = -1
                ComboBox5.SelectedIndex = -1
                GroupBox8.Enabled = False
                TabControl1.Enabled = False
            End If
        End If
    End Sub

    Private Sub Combobox1SelectionChangeCommitted()
        'Dim drv = ComboBox1.SelectedItem
        'Dim mydrv As DataRowView = BS.Current
        'mydrv.Row.Item("doctype") = drv.row.item("doctype")
        'mydrv.Row.Item("doctypeid") = drv.row.item("id")
        'mydrv.EndEdit()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.validate()
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



    Private Sub AddToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddToolStripMenuItem.Click
        Dim myItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim cm As ContextMenuStrip = DirectCast(myItem.Owner, ContextMenuStrip)

        Select Case cm.SourceControl.Name
            Case "DataGridView3"
                Dim myform As New FormHelperMulti(bsZetolMaster)
                If myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    Dim bsdetailcurrent As DataRowView = BS.Current
                    For Each mydrv As DataRowView In myform.SelectedDrv
                        Dim drv As DataRowView = bsZetol.AddNew
                        Try
                            drv.Item("attachmentid") = bsdetailcurrent.Item("id")
                            drv.Item("zetolid") = mydrv.Item("id")
                            drv.Item("zename") = mydrv.Item("name")
                            drv.EndEdit()
                        Catch ex As Exception
                            drv.CancelEdit()                       
                        End Try
                    Next
                End If

        End Select
    End Sub

    Private Sub DeleteItemToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteItemToolStripMenuItem.Click
        Dim myItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim cm As ContextMenuStrip = DirectCast(myItem.Owner, ContextMenuStrip)

        Select Case cm.SourceControl.Name
            Case "DataGridView3"
                If Not IsNothing(BSZetol.Current) Then
                    If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                        For Each drv As DataGridViewRow In DataGridView3.SelectedRows
                            BSZetol.RemoveAt(drv.Index)
                        Next
                    End If
                    DataGridView3.Invalidate()
                End If            
        End Select
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim openfiledialog1 As New OpenFileDialog
        If openfiledialog1.ShowDialog = DialogResult.OK Then
            Dim mydrv As DataRowView = BS.Current
            mydrv.Row.Item("docname") = IO.Path.GetFileName(openfiledialog1.FileName)
            mydrv.Row.Item("docext") = IO.Path.GetExtension(openfiledialog1.FileName)
            TextBox10.Text = openfiledialog1.FileName
        End If
    End Sub

   
    Private Sub ComboBox3_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectionChangeCommitted
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)

        '1. Force the Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next
        If Not IsNothing(ComboBox3.SelectedItem) Then
            Dim drv As DataRowView = ComboBox3.SelectedItem
            If Not IsNothing(BS.Current) Then
                Dim mydrv As DataRowView = BS.Current
                mydrv.Row.Item("doctype") = drv.Row.Item("doctype")
            End If
        End If
        DataGridView1.Invalidate()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim myform = New FormHelper(DataTypeHelperBS)
        myform.DataGridView1.Columns(0).DataPropertyName = "doctype"
        If myform.ShowDialog = DialogResult.OK Then
            'checkcombobox3()
            Dim drv As DataRowView = DataTypeHelperBS.Current
            Dim mydrv As DataRowView = BS.Current
            mydrv.Row.Item("doctype") = drv.Row.Item("doctype")
            mydrv.Row.Item("doctypeid") = drv.Row.Item("id")

            'Sync with bsDocType
            Dim itemfound As Integer = DataTypeBS.Find("id", drv.Row.Item("id"))
            DataTypeBS.Position = itemfound

        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim ImportFile = New OpenFileDialog
        ImportFile.FileName = "Select File"
        If ImportFile.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim mydrv As DataRowView = BS.Current
            Dim myobj As Object
            If mydrv.Row.Item("doctypeid") = SIF Then
                myobj = New ClassImportSIFAttachment(SITxBS, BS, ImportFile.FileName)
            Else
                myobj = New ClassImportIdentityAttachment(SITxBS, BS, ImportFile.FileName)
            End If
            If Not myobj.getrecord() Then
                MessageBox.Show(myobj.errormessage)
            Else

                DataGridView2.Invalidate()
                SITxBS.Position = 0
            End If
        End If
    End Sub
End Class
