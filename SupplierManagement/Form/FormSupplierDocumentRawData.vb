Imports System.Threading
Imports Microsoft.Office.Interop
Imports SupplierManagement.SharedClass
Imports System.Text
Imports SupplierManagement.PublicClass
Public Class FormSupplierDocumentRawData
    Dim myThread As New Threading.Thread(AddressOf DoWork)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim dr As DataRow = CType(bsHeader.Current, DataRowView).Row
        Dim myQueryWorksheetList As New List(Of QueryWorksheet)
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty



        Dim sqlstr As String
        'sqlstr = "select vd.vendorcode,v.vendorname::text,v.shortname::text,d.id as documentid,dt.doctypename,dl.levelname,de.expireddate,d.docdate,d.docname,d.docext,d.uploaddate,d.userid,d.remarks,vr.version,pt.payt || ' - ' || pt.details as paymentterm,sc.leadtime,sc.sasl,q.nqsu,p.projectname,sa.auditby,sa.audittype,sa.auditgrade,sef.score,sif.myyear,sif.turnovery,sif.turnovery1,sif.turnovery2,sif.turnovery3,sif.turnovery4,sif.ratioy,sif.ratioy1,sif.ratioy2,sif.ratioy3,sif.ratioy4,spp.*,sc.*,ps1.*,ps2* from doc.vendordoc vd" &
        '         " left join doc.header h on h.id = vd.headerid" &
        '         " left join vendor v on v.vendorcode = vd.vendorcode" &
        '         " left join doc.document d on d.id = vd.documentid" &
        '         " left join doc.version vr on vr.documentid = d.id" &
        '         " left join doc.generalcontract gt on gt.documentid = d.id" &
        '         " left join paymentterm pt on pt.paymenttermid = gt.paymentcode" &
        '         " left join doc.supplychain sc on sc.documentid = d.id" &
        '         " left join doc.qualityappendix q on q.documentid = d.id" &
        '         " left join doc.project p on p.documentid = d.id" &
        '         " left join doc.socialaudit sa on sa.documentid = d.id" &
        '         " left join doc.sef sef on sef.documentid = d.id" &
        '         " left join doc.sif sif on sif.documentid = d.id" &
        '         " left join doc.doctype dt on dt.id = d.doctypeid" &
        '         " left join doc.doclevel dl on dl.id = d.doclevelid" &
        '         " left join doc.docexpired de on de.documentid = d.id" &
        '         " left join supplierpanel spp on spp.vendorcode = vd.vendorcode" &
        '         " left joinn suppliercategory sc on sc.supplierscategoryid = spp.suppliercategoryid" &
        '         " left join doc.panelstatus ps1 on ps1.id = spp.fp" &
        '         " left join doc.panelstatus ps2 on ps2.id = spp.cp;"

        Dim criteria As String = String.Empty
        If Not CheckBox1.Checked Then
            criteria = " where not vd.vendorcode isnull"
        End If
        sqlstr = "select vd.vendorcode,v.vendorname::text,v.shortname2::text,d.id as documentid,dt.doctypename,dl.levelname,de.expireddate,d.docdate,d.docname,d.docext,d.uploaddate,d.userid,d.remarks,vr.version,pt.payt || ' - ' || pt.details as paymentterm,sc.leadtime,sc.sasl,q.nqsu,ps.returnrate,p.projectname,sa.auditby,sa.audittype,sa.auditgrade,sef.score,sif.myyear,sif.turnovery,sif.turnovery1,sif.turnovery2,sif.turnovery3,sif.turnovery4,sif.ratioy,sif.ratioy1,sif.ratioy2,sif.ratioy3,sif.ratioy4,spp.*,scy.*,ps1.*,ps2.*, doc.getstatusdoc(vd.status) from doc.document d" &
                 " left join doc.vendordoc vd on vd.documentid = d.id " &
                 " left join doc.header h on h.id = vd.headerid" &
                 " left join vendor v on v.vendorcode = vd.vendorcode" &
                 " left join doc.version vr on vr.documentid = d.id" &
                 " left join doc.generalcontract gt on gt.documentid = d.id" &
                 " left join paymentterm pt on pt.paymenttermid = gt.paymentcode" &
                 " left join doc.supplychain sc on sc.documentid = d.id" &
                 " left join doc.qualityappendix q on q.documentid = d.id" &
                 " left join doc.project p on p.documentid = d.id" &
                 " left join doc.projectspecification ps on ps.documentid = d.id" &
                 " left join doc.socialaudit sa on sa.documentid = d.id" &
                 " left join doc.sef sef on sef.documentid = d.id" &
                 " left join doc.sif sif on sif.documentid = d.id" &
                 " left join doc.doctype dt on dt.id = d.doctypeid" &
                 " left join doc.doclevel dl on dl.id = d.doclevelid" &
                 " left join doc.docexpired de on de.documentid = d.id " &
                 " left join supplierspanel spp on spp.vendorcode = vd.vendorcode" &
                 " left join supplierscategory scy on scy.supplierscategoryid = spp.supplierscategoryid" &
                 " left join doc.panelstatus ps1 on ps1.id = spp.fp" &
                 " left join doc.panelstatus ps2 on ps2.id = spp.cp " & criteria & ";"
        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("SupplierDocumentRawData{0:yyyyMMdd}.xlsx", Date.Today)

        'Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        'DirectoryBrowser.Description = "Which directory do you want to use?"
        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            'Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)

            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myqueryworksheet = New QueryWorksheet With {.DataSheet = 1,
                                                            .SheetName = "RawData",
                                                            .Sqlstr = sqlstr
                                                            }
            myQueryWorksheetList.Add(myqueryworksheet)


            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback)
            Dim myreport As New ExportToExcelFile(Me, myQueryWorksheetList, filename, reportname, mycallback, PivotCallback)

            myreport.Run(Me, e)

        End If
    End Sub

    Private Sub FormattingReport()
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable()
        'Throw New NotImplementedException
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
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

    Private Sub runreport(ByVal sender As System.Object, ByVal e As System.EventArgs)


        Dim mymessage As String = String.Empty

        Dim sqlstr As String = "select foo.*,vs.sbuname,case (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) when true then 'Expired' else doc.getstatusdoc(vd.status) end as statusname from (select * from doc.countdocumentpm1 union all select * from doc.vendormissingdocumentpm1) foo" &
                                " left join doc.vendorsbu vs on vs.vendorcode = foo.vendorcode left join doc.vendordoc vd on vd.id = foo.id left join doc.document d on d.id = foo.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid;"

        If Not HelperClass1.UserInfo.IsAdmin Then
            sqlstr = String.Format("with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname2::text" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{0}$'   order by vendorname)  " &
                 "select foo.*,vs.sbuname,case (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) when true then 'Expired' else doc.getstatusdoc(vd.status) end as statusname from (select * from doc.countdocumentpm1 union all select * from doc.vendormissingdocumentpm1) foo" &
                 " inner join va on va.vendorcode = foo.vendorcode " &
                                " left join doc.vendorsbu vs on vs.vendorcode = va.vendorcode left join doc.vendordoc vd on vd.id = foo.id left join doc.document d on d.id = foo.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid;", HelperClass1.UserInfo.userid)
        End If

        Dim mysaveform As New SaveFileDialog
        'mysaveform.FileName = String.Format("SupplierDocumentReport{0:yyyyMMdd}.xlsm", Date.Today)
        mysaveform.FileName = String.Format("SupplierDocumentReport{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1 'because hidden

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\\172.22.10.44\Users_I\Logistic Dept\KPI & Reporting\templates\KPI Graph Final.xltx")
            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\SupplierDocument.xltm")
            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\SupplierDocumentTemplate.xltm")
            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            myreport.Run(Me, e)
        End If
    End Sub
End Class