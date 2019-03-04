Imports System.Threading
Imports Microsoft.Office.Interop
Imports SupplierManagement.SharedClass
Imports System.Text
Imports SupplierManagement.PublicClass
Public Class FormSupplierDocumentSIFID

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

        Dim sqlstr As String = "with foo as (select * from doc.countdocumentpm union all  select * from doc.vendormissingdocumentpm)," &
            " sifid as (select * from crosstab('with dt as (select *,row_number() over(partition by documentid,labelid order by id,documentid,labelid desc) ," &
            " labelid::text || row_number() over(partition by documentid,labelid order by id,documentid,labelid desc) as newlabel" &
            " from doc.sitx order by id,documentid,labelid) select documentid,newlabel,value from dt','select ivalue from doc.paramdt where paramhdid = 10 order by nvalue')" &
            " as(documentid integer,companyname text,website text,customer1 text,customer2 text,customer3 text,brand1 text,brand2 text,brand3 text," &
            " mainshareholder1share text,mainshareholder2share text,mainshareholder3share text,oem text,odm text,mainproductsold1 text,mainproductsold2 text,mainproductsold3 text," &
            " status text,officeaddress text,contact text,contactposition text,contactmail text,visit text,iwhen text,who text,production text,capacityofproduction text," &
            " nofactory text,noworker text,noassemblyline text,totalarea text,factory1 text,factory2 text,factory3 text,product text,sbu1 text,familyname1 text,sbu2 text,familyname2 text," &
            " oemodm text,material1 text,material2 text,originalinfo text,turnoveryears text,customer4 text,synthesis text)) select foo.id,foo.vendorcode,foo.vendorname,foo.shortname,foo.gsm,foo.spm,foo.pm,foo.documentid,foo.doctypename,foo.countbydoc,foo.levelname,foo.expireddate,foo.docdate," &
            " foo.docname,foo.docext,foo.uploaddate,foo.userid,foo.remarks,foo.version,foo.projectname,foo.myyear,foo.turnovery,foo.turnovery1,foo.turnovery2,foo.turnovery3,foo.turnovery4," &
            " foo.ratioy,foo.ratioy1,foo.ratioy2,foo.ratioy3,foo.ratioy4," &
            " sifid.*,foo.category,foo.panelstatus1,foo.paneldescription1,foo.panelstatus2,foo.paneldescription2,foo.fp,foo.cp,foo.sp,foo.mould,foo.groupact,foo.producttype,foo.statusname,foo.rank," &
            " vs.sbuname,case (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) when true then 'Expired' else doc.getstatusdoc(vd.status) end as statusname" &
            " from foo" &
            " left join doc.vendorsbu vs on vs.vendorcode = foo.vendorcode " &
            " left join doc.vendordoc vd on vd.id = foo.id " &
            " left join doc.document d on d.id = foo.documentid " &
            " left join doc.doctypereminder dr on dr.id = d.doctypeid" &
            " left join sifid on sifid.documentid = foo.documentid" &
            " where doctypename in ('SIF','Identity sheet')"

        If Not HelperClass1.UserInfo.IsAdmin Then
            sqlstr = String.Format("with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{0}$'   order by vendorname)," &
                                   " foo as (select * from doc.countdocumentpm union all  select * from doc.vendormissingdocumentpm)," &
            " sifid as (select * from crosstab('with dt as (select *,row_number() over(partition by documentid,labelid order by id,documentid,labelid desc) ," &
            " labelid::text || row_number() over(partition by documentid,labelid order by id,documentid,labelid desc) as newlabel" &
            " from doc.sitx order by id,documentid,labelid) select documentid,newlabel,value from dt','select ivalue from doc.paramdt where paramhdid = 10 order by nvalue')" &
            " as(documentid integer,companyname text,website text,customer1 text,customer2 text,customer3 text,brand1 text,brand2 text,brand3 text," &
            " mainshareholder1share text,mainshareholder2share text,mainshareholder3share text,oem text,odm text,mainproductsold1 text,mainproductsold2 text,mainproductsold3 text," &
            " status text,officeaddress text,contact text,contactposition text,contactmail text,visit text,iwhen text,who text,production text,capacityofproduction text," &
            " nofactory text,noworker text,noassemblyline text,totalarea text,factory1 text,factory2 text,factory3 text,product text,sbu1 text,familyname1 text,sbu2 text,familyname2 text," &
            " oemodm text,material1 text,material2 text,originalinfo text,turnoveryears text,customer4 text,synthesis text)) select foo.id,foo.vendorcode,foo.vendorname,foo.shortname,foo.spm,foo.pm,foo.documentid,foo.doctypename,foo.countbydoc,foo.levelname,foo.expireddate,foo.docdate," &
            " foo.docname,foo.docext,foo.uploaddate,foo.userid,foo.remarks,foo.version,foo.projectname,foo.myyear,foo.turnovery,foo.turnovery1,foo.turnovery2,foo.turnovery3,foo.turnovery4," &
            " foo.ratioy,foo.ratioy1,foo.ratioy2,foo.ratioy3,foo.ratioy4," &
            " sifid.*,foo.category,foo.panelstatus1,foo.paneldescription1,foo.panelstatus2,foo.paneldescription2,foo.fp,foo.cp,foo.sp,foo.mould,foo.groupact,foo.producttype,foo.statusname,foo.rank," &
            " vs.sbuname,case (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) when true then 'Expired' else doc.getstatusdoc(vd.status) end as statusname" &
            " from foo" &
            " inner join va on va.vendorcode = foo.vendorcode " &
            " left join doc.vendorsbu vs on vs.vendorcode = foo.vendorcode " &
            " left join doc.vendordoc vd on vd.id = foo.id " &
            " left join doc.document d on d.id = foo.documentid " &
            " left join doc.doctypereminder dr on dr.id = d.doctypeid" &
            " left join sifid on sifid.documentid = foo.documentid " &
            " where doctypename in ('SIF','Identity sheet');", HelperClass1.UserInfo.userid)

            '"select foo.*,vs.sbuname,case (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) when true then 'Expired' else doc.getstatusdoc(vd.status) end as statusname from (select * from doc.countdocument union all select * from doc.vendormissingdocument) foo" &
            '" inner join va on va.vendorcode = foo.vendorcode " &
            '               " left join doc.vendorsbu vs on vs.vendorcode = va.vendorcode left join doc.vendordoc vd on vd.id = foo.id left join doc.document d on d.id = foo.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid;", HelperClass1.UserInfo.userid)
        End If

        Dim mysaveform As New SaveFileDialog

        mysaveform.FileName = String.Format("SupplierDocumentReportSIFID{0:yyyyMMdd}.xlsx", Date.Today)

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