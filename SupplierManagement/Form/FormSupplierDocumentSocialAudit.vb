Imports System.Threading
Imports Microsoft.Office.Interop
Imports SupplierManagement.SharedClass
Imports System.Text
Imports SupplierManagement.PublicClass
Public Class FormSupplierDocumentSocialAudit

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)


    Private Sub FormattingReport()
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable()
        'Throw New NotImplementedException
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""


        runreport(sender, e)

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

        Dim sqlstr As String = "with foo as (select * from doc.countdocumentpm union all  select * from doc.vendormissingdocumentpm)" &
            " select foo.id,foo.vendorcode,foo.vendorname,foo.shortname,foo.gsm,foo.spm,foo.pm,foo.documentid,foo.doctypename,foo.countbydoc,foo.levelname,foo.expireddate,foo.docdate, foo.docname,foo.docext,foo.uploaddate,foo.userid,foo.remarks,foo.version,foo.projectname," &
            " foo.auditby,foo.audittype,foo.auditgrade,sa.overallauditresult,sa.score,doc.getzetolname(foo.documentid) as zerotolerance," &
            " foo.category,foo.panelstatus1,foo.paneldescription1,foo.panelstatus2,foo.paneldescription2,foo.fp,foo.cp,foo.sp,foo.mould,foo.groupact,foo.producttype,foo.statusname,foo.rank," &
            " vs.sbuname,case (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) when true then 'Expired' else doc.getstatusdoc(vd.status) end as statusname from foo left join doc.vendorsbu vs on vs.vendorcode = foo.vendorcode  left join doc.vendordoc vd on vd.id = foo.id  left join doc.document d on d.id = foo.documentid  " &
            " left join doc.doctypereminder dr on dr.id = d.doctypeid " &
            " left join doc.socialaudit sa on sa.documentid = foo.documentid" &
            " where doctypename in ('Social Audit');"

        If Not HelperClass1.UserInfo.IsAdmin Then
            sqlstr = String.Format("with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname3::text as shortname" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{0}$'   order by vendorname)," &
                                   " foo as (select * from doc.countdocumentpm union all  select * from doc.vendormissingdocumentpm)" &
           " select foo.id,foo.vendorcode,foo.vendorname,foo.shortname,foo.gsm,foo.spm,foo.pm,foo.documentid,foo.doctypename,foo.countbydoc,foo.levelname,foo.expireddate,foo.docdate, foo.docname,foo.docext,foo.uploaddate,foo.userid,foo.remarks,foo.version,foo.projectname," &
            " foo.auditby,foo.audittype,foo.auditgrade,sa.overallauditresult,sa.score,doc.getzetolname(foo.documentid) as zerotolerance," &
            " foo.category,foo.panelstatus1,foo.paneldescription1,foo.panelstatus2,foo.paneldescription2,foo.fp,foo.cp,foo.sp,foo.mould,foo.groupact,foo.producttype,foo.statusname,foo.rank," &
            " vs.sbuname,case (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) when true then 'Expired' else doc.getstatusdoc(vd.status) end as statusname from foo" &
            " inner join va on va.vendorcode = foo.vendorcode" &
            " left join doc.vendorsbu vs on vs.vendorcode = foo.vendorcode  left join doc.vendordoc vd on vd.id = foo.id  left join doc.document d on d.id = foo.documentid  " &
            " left join doc.doctypereminder dr on dr.id = d.doctypeid " &
            " left join doc.socialaudit sa on sa.documentid = foo.documentid" &
            " where doctypename in ('Social Audit');", HelperClass1.UserInfo.userid)

            '"select foo.*,vs.sbuname,case (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) when true then 'Expired' else doc.getstatusdoc(vd.status) end as statusname from (select * from doc.countdocument union all select * from doc.vendormissingdocument) foo" &
            '" inner join va on va.vendorcode = foo.vendorcode " &
            '               " left join doc.vendorsbu vs on vs.vendorcode = va.vendorcode left join doc.vendordoc vd on vd.id = foo.id left join doc.document d on d.id = foo.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid;", HelperClass1.UserInfo.userid)
        End If

        Dim mysaveform As New SaveFileDialog

        mysaveform.FileName = String.Format("SupplierDocumentReportSocialAudit{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            myreport.Run(Me, e)
        End If
    End Sub

    Private Sub StatusStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub
End Class