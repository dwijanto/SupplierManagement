Imports System.Threading
Imports Microsoft.Office.Interop
Imports SupplierManagement.SharedClass
Imports System.Text
Imports SupplierManagement.PublicClass

Public Class FormSupplierMasterData

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)


    Private Sub FormattingReport()
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable()
        'Throw New NotImplementedException
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)


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

        'Dim sqlstr As String = String.Format("with fgcp as((select distinct vf.vendorcode,v.vendorname,vf.familyid,fpm.pmid from doc.vendorfamily vf" &
        '                                     " left join vendor v on v.vendorcode = vf.vendorcode" &
        '                                     " left join  doc.familypm fpm on fpm.familyid = vf.familyid" &
        '                                     " inner join doc.vendorfamilyone({0:yyyyMM},{1:yyyyMM}) vfo on vfo.vendorcode = vf.vendorcode and vfo.familyid = vf.familyid order by v.vendorname)" &
        '                                     " union all" &
        '                                     " (select distinct vp.vendorcode,v.vendorname,null::integer as familyid,vp.pmid from doc.vendorpm vp left join vendor v on v.vendorcode = vp.vendorcode order by v.vendorname))," &
        '                                     " nopm as(select  vendorcode from doc.vendorstatus except select vendorcode from fgcp)," &
        '                                     " allvendor as(select vendorcode,familyid,pmid from fgcp union all (select vendorcode,null::integer,null::integer from nopm order by vendorcode) )" &
        '                                     " select distinct vp.vendorcode,v.vendorname,v.shortname,mu1.username as gsm,mu2.username as spm,mu.username as pm,f.familyid,f.familyname,vvs.status,vpt.producttype from allvendor vp" &
        '                                     " left join vendor v on v.vendorcode = vp.vendorcode" &
        '                                     " left join family f on f.familyid = vp.familyid " &
        '                                     " left join officerseb o on o.ofsebid = vp.pmid" &
        '                                     " left join masteruser mu on mu.id = o.muid" &
        '                                     " left join doc.vendorgsm gsm on gsm.vendorcode = vp.vendorcode" &
        '                                     " left join officerseb o1 on o1.ofsebid = gsm.gsmid" &
        '                                     " left join masteruser mu1 on mu1.id = o1.muid" &
        '                                     " left join officerseb o2 on o2.ofsebid = o.parent" &
        '                                     " left join masteruser mu2 on mu2.id = o2.muid" &
        '                                     " left join doc.vendorstatus vs on vs.vendorcode = vp.vendorcode" &
        '                                     " left join doc.viewvendorstatus vvs on vvs.statusid = vs.status" &
        '                                     " left join doc.viewproducttype vpt on vpt.producttypeid = vs.producttypeid order by status,producttype,shortname", DateTimePicker1.Value, DateTimePicker2.Value)

        Dim sqlstr As String = String.Format("with fg as (select distinct vf.vendorcode,v.vendorname,vf.familyid,fpm.pmid from doc.vendorfamily vf" &
                                             " left join vendor v on v.vendorcode = vf.vendorcode" &
                                             " left join  doc.familypm fpm on fpm.familyid = vf.familyid" &
                                             " inner join doc.vendorfamilyone({0:yyyyMM},{1:yyyyMM}) vfo on vfo.vendorcode = vf.vendorcode and vfo.familyid = vf.familyid order by v.vendorname)," &
                                             " cp as (select distinct vp.vendorcode,v.vendorname,null::integer as familyid,vp.pmid from doc.vendorpm vp left join vendor v on v.vendorcode = vp.vendorcode where vp.vendorcode not in (select vendorcode from fg) order by v.vendorname)," &
                                             " fgcp as( select * from fg" &
                                             " union all" &
                                             " select * from cp)," &
                                             " nopm as(select  vendorcode from doc.vendorstatus except select vendorcode from fgcp)," &
                                             " allvendor as(select vendorcode,familyid,pmid from fgcp union all (select vendorcode,null::integer,null::integer from nopm order by vendorcode) )" &
                                             " select distinct vp.vendorcode,v.vendorname,v.shortname3 as shortname,mu1.username as gsm,mu2.username as spm,mu.username as pm,f.familyid,f.familyname,vvs.status,vpt.producttype from allvendor vp" &
                                             " left join vendor v on v.vendorcode = vp.vendorcode" &
                                             " left join family f on f.familyid = vp.familyid " &
                                             " left join officerseb o on o.ofsebid = vp.pmid" &
                                             " left join masteruser mu on mu.id = o.muid" &
                                             " left join doc.vendorgsm gsm on gsm.vendorcode = vp.vendorcode" &
                                             " left join officerseb o1 on o1.ofsebid = gsm.gsmid" &
                                             " left join masteruser mu1 on mu1.id = o1.muid" &
                                             " left join officerseb o2 on o2.ofsebid = o.parent" &
                                             " left join masteruser mu2 on mu2.id = o2.muid" &
                                             " left join doc.vendorstatus vs on vs.vendorcode = vp.vendorcode" &
                                             " left join doc.viewvendorstatus vvs on vvs.statusid = vs.status" &
                                             " left join doc.viewproducttype vpt on vpt.producttypeid = vs.producttypeid order by status,producttype,shortname", DateTimePicker1.Value, DateTimePicker2.Value)


        Dim mysaveform As New SaveFileDialog

        mysaveform.FileName = String.Format("SupplierMasterData{0:yyyyMMdd}.xlsx", Date.Today)

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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""


        runreport(sender, e)
    End Sub
End Class