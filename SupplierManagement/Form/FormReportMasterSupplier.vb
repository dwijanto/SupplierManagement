Imports System.Threading
Imports Microsoft.Office.Interop
Imports SupplierManagement.SharedClass
Imports System.Text
Imports SupplierManagement.PublicClass
Public Class FormReportMasterSupplier

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

        Dim sqlstr As String = "select * from doc.vendormasterreportpm"

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("VendorMasterReport{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1 'because hidden

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            myreport.Run(Me, e)
        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)

    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
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

End Class