Imports System.Threading
Public Class FormVendorInformationTask
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myAdapter As VendorInformationTaskAdapter = New VendorInformationTaskAdapter
    Dim helper1 As HelperClass = HelperClass.getInstance
    Dim DS As DataSet
    WithEvents EPCheckBox1 As New CheckBox



    Dim dtpicker1 As New DateTimePicker
    Dim fieldList() As String = {"statusname", "applicantname", "vendorcode", "vendorname", "shortname", "modifiedfield", "suppliermodificationid"}
    Dim criteria As String
    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripTextBox1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        criteria = ""

        'If ToolStripTextBox1.Text <> "" Then
        '    criteria = String.Format("and (lower({0}) like '%{1}%')", fieldList(ToolStripComboBox1.SelectedIndex), ToolStripTextBox1.Text.ToLower)
        'End If

        If EPCheckBox1.Checked Then
            criteria = String.Format(" and vim.applicantdate = '{0:yyyy-MM-dd}'", dtpicker1.Value)
        End If
        DS = New DataSet
        Dim myUser As String = helper1.UserId
        Try
            If Not helper1.UserInfo.IsAdmin Then
                If myAdapter.loadData(DS, myUser, criteria) Then
                    ProgressReport(4, "Initialize Data")
                End If
            Else
                If myAdapter.loadData(DS, criteria) Then
                    ProgressReport(4, "Initialize Data")
                End If
            End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
    End Sub

    Private Sub FormVendorInformationTask_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
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
                        DataGridView1.AutoGenerateColumns = False
                        'DataGridView1.DataSource = DS.Tables(0)
                        DataGridView1.DataSource = myAdapter.bs1
                        DataGridView2.AutoGenerateColumns = False
                        'DataGridView2.DataSource = DS.Tables(1)
                        DataGridView2.DataSource = myAdapter.bs2
                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        loaddata()
    End Sub


    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim drv As DataRowView = myAdapter.getCurrentRecordTx
        Dim myform As Object
        If IsDBNull(drv.Row.Item("extid")) Then
            myform = New FormVendorInformationModification(TxEnum.ValidateRecord, drv.Item("id"))
        Else
            myform = New FormNewVendor(TxEnum.ValidateRecord, drv.Item("id"))           
        End If

        myform.Show()
        'If IsDBNull(drv.Row.Item("extid")) Then
        '    Dim myform = New FormVendorInformationModification(TxEnum.ValidateRecord, drv.Item("id"))
        'Else
        '    Dim myform = New FormNewVendor(TxEnum.ValidateRecord, drv.Item("id"))
        '    AddHandler FormNewVendor.myRefreshData, AddressOf loaddata

        'End If
    End Sub

    Private Sub DataGridView2_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        Dim drv As DataRowView = myAdapter.getCurrentRecordHistory
        Dim myform As Object
        If IsDBNull(drv.Row.Item("extid")) Then
            myform = New FormVendorInformationModification(TxEnum.HistoryRecord, drv.Item("id"))
        Else
            myform = New FormNewVendor(TxEnum.HistoryRecord, drv.Item("id"))
        End If

        myform.Show()
    End Sub

    Private Sub ToolStripComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        RemoveHandler FormNewVendor.RefreshDataGridView, AddressOf RefreshDataGridView
        AddHandler FormNewVendor.RefreshDataGridView, AddressOf RefreshDataGridView

        ' Add any initialization after the InitializeComponent() call.
        ToolStripComboBox1.SelectedIndex = 0
        EPCheckBox1.Text = "Applicant Date"
        EPCheckBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        ToolStrip1.Items.Insert(5, New ToolStripControlHost(EPCheckBox1))
        dtpicker1.CustomFormat = "dd-MMM-yyyy"
        dtpicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        dtpicker1.Size = New Point(100, 20)
        ToolStrip1.Items.Insert(6, New ToolStripControlHost(dtpicker1))

        'RemoveHandler FormNewVendor.RefreshData, AddressOf loaddata
        'AddHandler FormNewVendor.RefreshData, AddressOf loaddata
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        applyFilter()
    End Sub

    Private Sub applyFilter()


        If ToolStripComboBox1.SelectedIndex >= 0 Then
            Try
                Dim myfilter = String.Format("{0} like '*{1}*'", fieldList(ToolStripComboBox1.SelectedIndex), ToolStripTextBox1.Text)
                myAdapter.bs1.Filter = myfilter
                myAdapter.bs2.Filter = myfilter
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try

        End If

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not IsNothing(myAdapter.bs2.Current) Then
            Dim mydrv As DataRowView = myAdapter.bs2.Current

            Dim myform1 As FormVIMEmailModification = New FormVIMEmailModification(mydrv.Row.Item("id"), mydrv.Row.Item("vendorcode"))

            RemoveHandler myform1.RefreshData, AddressOf loaddata
            AddHandler myform1.RefreshData, AddressOf loaddata

            myform1.Show()

        End If
       

    End Sub

    Private Sub RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)
        loaddata()
    End Sub

    
End Class