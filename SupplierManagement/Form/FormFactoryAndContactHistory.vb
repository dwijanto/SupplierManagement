Imports System.Threading
Imports Microsoft.Office.Interop
Imports SupplierManagement.SharedClass
Imports System.Text
Imports SupplierManagement.PublicClass
Public Class FormFactoryAndContactHistory
    Dim myQueryWorksheetList As List(Of QueryWorksheet)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""


        runreport(sender, e)


    End Sub
    Private Sub runreport(ByVal sender As System.Object, ByVal e As System.EventArgs)
        myQueryWorksheetList = New List(Of QueryWorksheet)

        Dim mymessage As String = String.Empty

        Dim sqlstr As String = String.Format("select * from doc.rptcontacthistory('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date) ", DateTimePicker1.Value, DateTimePicker2.Value)
        Dim sqlstr2 As String = String.Format("select * from doc.rptfactoryhistorypm('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date) ", DateTimePicker1.Value, DateTimePicker2.Value)
        Dim sqlstr3 As String = String.Format("select * from doc.rptvendorfactorycontactdeletedpm('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date) ", DateTimePicker1.Value, DateTimePicker2.Value)
        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("FactoryAndContactHistory{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
            Dim myqueryworksheet = New QueryWorksheet With {.DataSheet = 1,
                                                           .SheetName = "Contact",
                                                           .Sqlstr = sqlstr}
            myQueryWorksheetList.Add(myqueryworksheet)
            myqueryworksheet = New QueryWorksheet With {.DataSheet = 2,
                                                           .SheetName = "Factory",
                                                           .Sqlstr = sqlstr2}
            myQueryWorksheetList.Add(myqueryworksheet)

            myqueryworksheet = New QueryWorksheet With {.DataSheet = 3,
                                                          .SheetName = "Deleted",
                                                          .Sqlstr = sqlstr3}
            myQueryWorksheetList.Add(myqueryworksheet)

            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            Dim myreport As New ExportToExcelFile(Me, myQueryWorksheetList, filename, reportname, mycallback, PivotCallback)
            myreport.Run(Me, e)
        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)

    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
    End Sub

    'Private Sub DoWork()
    '    ProgressReport(1, "Updating SBU... Please wait")
    '    ProgressReport(6, "")
    '    Dim myvendorsbu = New VendorSBU
    '    myvendorsbu.UpdateVendorSBU()
    '    myvendorsbu.UpdateShortnameInfo()
    '    ProgressReport(1, "")
    '    ProgressReport(4, "")
    '    ProgressReport(5, "")
    'End Sub
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

End Class