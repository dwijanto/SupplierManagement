Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class FormHistorySPMPMAssignment

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents ToolingListBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder


    Dim myarray() = {"", "v.vendorcode::text", "lower(vendorname::text)", "lower(shortname)", "lower(o.officersebname)", "lower(op.officersebname)"}
    Private Sqlstr As String
    Dim SqlstrReport As String
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        SqlstrReport = "select vsp.vendorcode,v.vendorname::text, v.shortname::text,o.officersebname::text as spmname,vsp.spmeffectivedate,op.officersebname::text as pmname,vsp.pmeffectivedate from vendorspmpm vsp" &
                                   " left join vendor v on v.vendorcode = vsp.vendorcode" &
                                   " left join officerseb o on o.ofsebid = v.ssmidpl" &
                                   " left join officerseb op on op.ofsebid = v.pmid" &
                                   " {0} order by vendorname"
    End Sub
    Public Sub New(ByVal sqlstr, ByVal Report, ByVal FileName)

        ' This call is required by the designer.
        InitializeComponent()
        Me.Sqlstr = sqlstr

        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        ProgressReport(7, "InitFilter")

        Sqlstr = String.Format(SqlstrReport, sb.ToString.ToLower)

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("ReportHistorySPM-PM{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\\172.22.10.44\Users_I\Logistic Dept\KPI & Reporting\templates\KPI Graph Final.xltx")
            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\SupplierDocument.xltm")
            Dim myreport As New ExportToExcelFile(Me, Sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            myreport.Run(Me, New System.EventArgs)
        End If

        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)

     
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)

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
                            ToolingListBS = New BindingSource
                            DS.Tables(0).TableName = "ToolingList"

                            ToolingListBS.DataSource = DS.Tables(0)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = ToolingListBS
                            DataGridView1.RowTemplate.Height = 22

                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                    Case 7
                        If ComboBox1.SelectedIndex > 0 And TextBox1.Text <> "" Then

                            sb.Append(String.Format("where {0} like '%{1}%'", myarray(ComboBox1.SelectedIndex), TextBox1.Text))
                        End If
                        If ComboBox2.SelectedIndex > 0 And TextBox2.Text <> "" Then
                            If sb.Length > 0 Then
                                sb.Append(String.Format(" and {0} like '%{1}%'", myarray(ComboBox2.SelectedIndex), TextBox2.Text))
                            Else
                                sb.Append(String.Format(" where {0} like '%{1}%'", myarray(ComboBox2.SelectedIndex), TextBox2.Text))
                            End If

                        End If
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf GetToWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub


    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loadReport()
    End Sub
    
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        loaddata()
    End Sub

    Private Sub GetToWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        ProgressReport(7, "InitFilter")

        Dim Sqlstr = String.Format(SqlstrReport, sb.ToString.ToLower)
        If DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "History"

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

    Private Sub loadReport()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""

            myThread = New Thread(AddressOf DoWork)
            myThread.SetApartmentState(ApartmentState.STA)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
End Class