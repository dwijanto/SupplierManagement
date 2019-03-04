Public Class FormReportVendorFamilyAssignment   

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        
        runreport(sender, e)
        
    End Sub
    Private Sub runreport(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim mymessage As String = String.Empty

        Dim sqlstr As String = "with status as (select pd.ivalue as statusid,pd.paramname as statusname from doc.paramhd  ph" &
                " left join doc.paramdt pd on pd.paramhdid = ph.paramhdid" &
                " where ph.paramname = 'vendorstatus' order by pd.ivalue)," &
                " type as (select p.paramname as producttype, p.ivalue as producttypeid from doc.paramdt p " &
                " left join doc.paramhd ph on ph.paramhdid = p.paramhdid where ph.paramname = 'producttype')," &
                " fg as (select vf.vendorcode,v.vendorname::text,v.shortname::text,s.statusname,tp.producttype," &
                " vf.familyid as familylv,f.familyname as familylv1desc,ss.sbuname::text,mu.username as pm," &
                " mu2.username as spm,mu3.username as gsm,vf.pmeffectivedate,vf.spmeffectivedate,vgsm.effectivedate as gsmeffectivedate" &
                "  from doc.vendorfamily vf" &
                " left join vendor v on v.vendorcode = vf.vendorcode left join doc.vendorstatus vs on vs.vendorcode = vf.vendorcode" &
                " left join status s on s.statusid = vs.status left join type tp on tp.producttypeid = vs.producttypeid left join doc.familypm fp on fp.familyid = vf.familyid" &
                " left join family f on f.familyid = vf.familyid left join doc.familygroupsbu fg on fg.familyid = vf.familyid left join sbusap ss on ss.sbuid = fg.sbusapid" &
                " left join officerseb o on o.ofsebid = fp.pmid left join masteruser mu on mu.id = o.muid left join officerseb spm on spm.ofsebid = o.parent" &
                " left join masteruser mu2 on mu2.id = spm.muid left join doc.vendorgsm vgsm on vgsm.vendorcode = vf.vendorcode left join officerseb gsm on gsm.ofsebid = vgsm.gsmid" &
                " left join masteruser mu3 on mu3.id = gsm.muid order by vendorname)," &
                " cp as (select vp.vendorcode,v.vendorname::text,v.shortname::text,s.statusname,tp.producttype,null::integer as familylv,null::text as familylv1desc,null::text as sbuname,mu.username as pm," &
                " mu2.username as spm,mu3.username as gsm,vp.pmeffectivedate,vp.spmeffectivedate,vgsm.effectivedate as gsmeffectivedate from doc.vendorpm vp" &
                " left join vendor v on v.vendorcode = vp.vendorcode left join doc.vendorstatus vs on vs.vendorcode = vp.vendorcode left join status s on s.statusid = vs.status" &
                " left join type tp on tp.producttypeid = vs.producttypeid left join officerseb o on o.ofsebid = vp.pmid left join masteruser mu on mu.id = o.muid" &
                " left join officerseb spm on spm.ofsebid = o.parent left join masteruser mu2 on mu2.id = spm.muid  left join doc.vendorgsm vgsm on vgsm.vendorcode = vp.vendorcode" &
                " left join officerseb gsm on gsm.ofsebid = vgsm.gsmid left join masteruser mu3 on mu3.id = gsm.muid order by vendorname)," &
                " fgcp as ((select * from fg) union all (select * from cp))" &
                " select * from fgcp order by vendorcode,familylv;"
        sqlstr = "with status as (select pd.ivalue as statusid,pd.paramname as statusname from doc.paramhd  ph" &
                " left join doc.paramdt pd on pd.paramhdid = ph.paramhdid" &
                " where ph.paramname = 'vendorstatus' order by pd.ivalue)," &
                " type as (select p.paramname as producttype, p.ivalue as producttypeid from doc.paramdt p " &
                " left join doc.paramhd ph on ph.paramhdid = p.paramhdid where ph.paramname = 'producttype')," &
                " fg as (select vf.vendorcode,v.vendorname::text,v.shortname::text,s.statusname,tp.producttype," &
                " vf.familyid as familylv,f.familyname as familylv1desc,ss.sbuname::text,mu.username as pm," &
                " mu2.username as spm,mu3.username as gsm,vf.pmeffectivedate,vf.spmeffectivedate,vgsm.effectivedate as gsmeffectivedate" &
                "  from doc.vendorfamilyex vf" &
                " left join vendor v on v.vendorcode = vf.vendorcode left join doc.vendorstatus vs on vs.vendorcode = vf.vendorcode" &
                " left join status s on s.statusid = vs.status left join type tp on tp.producttypeid = vs.producttypeid left join doc.familypm fp on fp.familyid = vf.familyid" &
                " left join family f on f.familyid = vf.familyid left join doc.familygroupsbu fg on fg.familyid = vf.familyid left join sbusap ss on ss.sbuid = fg.sbusapid" &
                " left join officerseb o on o.ofsebid = vf.pmid left join masteruser mu on mu.id = o.muid left join officerseb spm on spm.ofsebid = o.parent" &
                " left join masteruser mu2 on mu2.id = spm.muid left join doc.vendorgsm vgsm on vgsm.vendorcode = vf.vendorcode left join officerseb gsm on gsm.ofsebid = vgsm.gsmid" &
                " left join masteruser mu3 on mu3.id = gsm.muid order by vendorname)," &
                " cp as (select vp.vendorcode,v.vendorname::text,v.shortname::text,s.statusname,tp.producttype,null::integer as familylv,null::text as familylv1desc,null::text as sbuname,mu.username as pm," &
                " mu2.username as spm,mu3.username as gsm,vp.pmeffectivedate,vp.spmeffectivedate,vgsm.effectivedate as gsmeffectivedate from doc.vendorpm vp" &
                " left join vendor v on v.vendorcode = vp.vendorcode left join doc.vendorstatus vs on vs.vendorcode = vp.vendorcode left join status s on s.statusid = vs.status" &
                " left join type tp on tp.producttypeid = vs.producttypeid left join officerseb o on o.ofsebid = vp.pmid left join masteruser mu on mu.id = o.muid" &
                " left join officerseb spm on spm.ofsebid = o.parent left join masteruser mu2 on mu2.id = spm.muid  left join doc.vendorgsm vgsm on vgsm.vendorcode = vp.vendorcode" &
                " left join officerseb gsm on gsm.ofsebid = vgsm.gsmid left join masteruser mu3 on mu3.id = gsm.muid order by vendorname)," &
                " fgcp as ((select * from fg) union all (select * from cp))" &
                " select * from fgcp order by vendorcode,familylv;"

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("VendorFamilyAssignment{0:yyyyMMdd}.xlsx", Date.Today)

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