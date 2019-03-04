Imports System.Threading
Imports Microsoft.Office.Interop
Imports SupplierManagement.SharedClass
Imports System.Text
Imports SupplierManagement.PublicClass

Public Class FormReportDocumentCount

    Dim myThread As New Threading.Thread(AddressOf DoWork)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        If CheckBox1.Checked Then
            ' FileName = Application.StartupPath & "\PrintOut"
            If Not myThread.IsAlive Then
                Try
                    myThread = New System.Threading.Thread(New ThreadStart(AddressOf DoWork))
                    myThread.SetApartmentState(ApartmentState.MTA)
                    myThread.Start()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Else
                MsgBox("Please wait until the current process is finished")
            End If

        Else
            runreport(sender, e)
        End If

    End Sub
    Private Sub runreport(ByVal sender As System.Object, ByVal e As System.EventArgs)
       

        Dim mymessage As String = String.Empty

        Dim sqlstr As String = "select foo.*,vs.sbuname from (select * from doc.countdocument union all select * from doc.vendormissingdocument) foo" &
                                " left join doc.vendorsbu vs on vs.vendorcode = foo.vendorcode;"

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("SupplierDocumentReport{0:yyyyMMdd}.xlsm", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 8 'because hidden

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\\172.22.10.44\Users_I\Logistic Dept\KPI & Reporting\templates\KPI Graph Final.xltx")
            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\SupplierDocument.xltm")
            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\SupplierDocumentTemplate.xltm")
            myreport.Run(Me, e)
        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
        Dim oXl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        oXl = owb.Parent


        owb.Worksheets(8).select()

        Dim osheet = owb.Worksheets(8)
        osheet.name = "RawData"

        owb.Names.Add("rawdata", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C2),COUNTA('RawData'!R1))")

        owb.Worksheets(1).select()
        osheet = owb.Worksheets(1)
        'oXl.Run("ShowFG")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        'oXl.Run("ShowCP")
        osheet.PivotTables("PivotTable2").PivotCache.Refresh()
        'oXl.Run("OpenTabs")

        osheet = owb.Worksheets(2)
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        
        osheet = owb.Worksheets(3)
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        osheet = owb.Worksheets(4)
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        osheet = owb.Worksheets(5)
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        osheet = owb.Worksheets(6)
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        osheet = owb.Worksheets(1)
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        osheet = owb.Worksheets(7)
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        
        'oXl.Run("BACK")
    End Sub

    Private Sub DoWork()
        ProgressReport(1, "Updating SBU... Please wait")
        ProgressReport(6, "")
        Dim myvendorsbu = New VendorSBU
        myvendorsbu.UpdateVendorSBU()
        myvendorsbu.UpdateShortnameInfo()
        ProgressReport(1, "")
        ProgressReport(4, "")
        ProgressReport(5, "")
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
                        ToolStripStatusLabel2.Text = message
                    Case 4
                        runreport(Me, New System.EventArgs)
                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub


    Private Sub FormReportDocumentCount_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class