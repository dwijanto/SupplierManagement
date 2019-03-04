Imports System.Threading
Public Class FormFindVendorInformationModification
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Private myAdapter As New VendorInformationModificationAdapter
    Private Criteria As String = String.Empty
    Dim BS As BindingSource
    WithEvents EPCheckBox1 As New CheckBox
    Dim dtpicker1 As New DateTimePicker
    'Dim fieldList() As String = {"v.vendorname", "v.shortname", "u.applicantname", "v.vendorcode::character varying", "lateststatus", "cr.username"}
    Dim fieldList() As String = {"vendorname", "shortname", "applicantname", "vendorcode::character varying", "lateststatus", "creatorname"}

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        getCriteria()
       

        'If TextBox1.Text <> "" Then
        '    If IsNumeric(TextBox1.Text) Then
        '        Criteria = String.Format("where u.vendorcode = {0}", TextBox1.Text)
        '    Else
        '        ProgressReport(1, "Incorrect data type.")
        '    End If

        'End If


        If Not myThread.IsAlive Then
            'ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Try

            If myAdapter.loaddata(Criteria) Then
                ProgressReport(4, "InitData")
                ProgressReport(1, "Loading Data.Done!")
                ProgressReport(5, "Continuous")
            End If
        Catch ex As Exception
            ProgressReport(1, "Loading Data. Error::" & ex.Message)
            ProgressReport(5, "Continuous")
        End Try
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
                        bs = New BindingSource
                        bs.DataSource = myAdapter.getBindingSource
                        DataGridView1.DataSource = bs
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

    Private Sub FormFindVendorInformationModification_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ProgressReport(1, "Ready.")

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        showdetail()
    End Sub

    Private Sub showdetail()
        If Not IsNothing(BS) Then


            Dim drv As DataRowView = BS.Current

            Dim myFormVendor = New FormVendorInformationModification(TxEnum.UpdateRecord, drv.Row.Item("id"))
            myFormVendor.ShowDialog()
        End If
    End Sub


    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        showdetail()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'Add DateTimePicker
        ToolStripComboBox1.SelectedIndex = 0
        EPCheckBox1.Text = "Applicant Date"
        EPCheckBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        ToolStrip1.Items.Insert(4, New ToolStripControlHost(EPCheckBox1))
        dtpicker1.CustomFormat = "dd-MMM-yyyy"
        dtpicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        dtpicker1.Size = New Point(100, 20)
        ToolStrip1.Items.Insert(5, New ToolStripControlHost(dtpicker1))
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
       

        Dim DS = New DataSet
        Dim mymessage As String = String.Empty
        getCriteria()
        Dim sqlstr = myAdapter.myModel.GetReportQuery(Criteria)

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("VIMReport{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
            Dim myreport As New ExportToExcelFile(Me, Sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            myreport.Run(Me, New System.EventArgs)
        End If

        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub

    Private Sub FormattingReport()
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable()
        'Throw New NotImplementedException
    End Sub

    Private Sub getCriteria()
        Criteria = ""
        If ToolStripTextBox1.Text <> "" Then
            Criteria = String.Format("where lower({0}) like '%{1}%'", fieldList(ToolStripComboBox1.SelectedIndex), ToolStripTextBox1.Text.ToLower)
        End If
        If EPCheckBox1.Checked Then
            Criteria = Criteria + IIf(Criteria = "", "where", " and") + String.Format(" q.applicantdate = '{0:yyyy-MM-dd}'", dtpicker1.Value)
        End If
    End Sub

End Class