Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Imports Microsoft.Office.Interop

Public Class FormDocType
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim DS As DataSet
    Dim VendorBS As BindingSource
    Dim ShortnameBS As BindingSource
    Dim PMBS As BindingSource
    Dim SPMBS As BindingSource
    Dim ProductTypeBS As BindingSource
    Dim StatusNameBS As BindingSource
    Dim SBUBS As BindingSource
    Dim DocTypeBS As BindingSource
    Dim docTypeId As Integer
    Private Sub loaddata()
        'MessageBox.Show(myuser)
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Dim mymessage As String = String.Empty
        ' " left join doc.vendorstatus vs on vs.vendorcode = v1.vendorcode " &
        '" where(vs.status = 1)" &
        Dim sb As New StringBuilder

        'Dim sqlstr As String = "select distinct v1.vendorcode::text || ' - ' || v1.vendorname as name, v1.vendorcode ,v1.vendorname::text from doc.vendordoc vd " &
        '                       " left join vendor v on v.vendorcode = vd.vendorcode" &
        '                       " left join vendor v1 on v1.shortname = v.shortname" &
        '                       " order by v1.vendorname::text;" &
        '                       " select distinct v1.shortname::text from doc.vendordoc vd" &
        '                       " left join vendor v on v.vendorcode = vd.vendorcode " &
        '                       " left join vendor v1 on v1.shortname = v.shortname " &
        '                       " order by v1.shortname::text; " &
        '                       " select distinct o.officersebname::text as spm from doc.vendordoc vd" &
        '                       " left join vendor v on v.vendorcode = vd.vendorcode " &
        '                       " left join officerseb o on o.ofsebid = v.ssmidpl " &
        '                       " order by o.officersebname::text;" &
        '                       " select distinct o.officersebname::text as pm from doc.vendordoc vd" &
        '                       " left join vendor v on v.vendorcode = vd.vendorcode " &
        '                       " left join officerseb o on o.ofsebid = v.pmid " &
        '                       " order by o.officersebname::text ;" &
        '                       " select dt.paramname as producttype from doc.paramdt dt" &
        '                       " where(dt.paramhdid = 3)" &
        '                       " order by ivalue;" &
        '                       " select dt.paramname as statusname from doc.paramdt dt" &
        '                       " where(dt.paramhdid = 2)" &
        '                       " order by ivalue;" &
        '                       " select sbuname from sbusap" &
        '                       " order by sbuname;" &
        '                       " select doctypename,id from doc.doctype" &
        '                       " order by doctypename;"
        'viewvendorpm replace viewvendorfamilypm
        Dim sqlstr As String = "select distinct v1.vendorcode::text || ' - ' || v1.vendorname as name, v1.vendorcode ,v1.vendorname::text from doc.vendordoc vd " &
                               " left join vendor v on v.vendorcode = vd.vendorcode" &
                               " left join vendor v1 on v1.shortname = v.shortname" &
                               " order by v1.vendorname::text;" &
                               " select distinct v1.shortname::text from doc.vendordoc vd" &
                               " left join vendor v on v.vendorcode = vd.vendorcode " &
                               " left join vendor v1 on v1.shortname = v.shortname " &
                               " order by v1.shortname::text; " &
                               " select distinct mu2.username as spm from doc.vendordoc vd" &
                               " left join doc.viewvendorpm v on v.vendorcode = vd.vendorcode " &
                               " left join officerseb o on o.ofsebid = v.pmid " &
                               " left join officerseb o1 on o1.ofsebid = o.parent" &
                               " LEFT JOIN masteruser mu2 ON mu2.id = o1.muid" &
                               " order by mu2.username;" &
                               " select distinct mu2.username as pm from doc.vendordoc vd" &
                               " left join doc.viewvendorpm v on v.vendorcode = vd.vendorcode " &
                               " left join officerseb o on o.ofsebid = v.pmid " &
                               " LEFT JOIN masteruser mu2 ON mu2.id = o.muid" &
                               " order by mu2.username;" &
                               " select dt.paramname as producttype from doc.paramdt dt" &
                               " where(dt.paramhdid = 3)" &
                               " order by ivalue;" &
                               " select dt.paramname as statusname from doc.paramdt dt" &
                               " where(dt.paramhdid = 2)" &
                               " order by ivalue;" &
                               " select sbuname from sbusap" &
                               " order by sbuname;" &
                               " select doctypename,id from doc.doctype" &
                               " order by doctypename;"
        If Not HelperClass1.UserInfo.IsAdmin Then
            'sqlstr = "select distinct v1.vendorcode::text || ' - ' || v1.vendorname as name, v1.vendorcode ,v1.vendorname::text from doc.vendordoc vd " &
            '                   " left join vendor v on v.vendorcode = vd.vendorcode" &
            '                   " left join vendor v1 on v1.shortname = v.shortname" &
            '                   " order by v1.vendorname::text;" &
            '                   " select distinct v1.shortname::text from doc.vendordoc vd" &
            '                   " left join vendor v on v.vendorcode = vd.vendorcode " &
            '                   " left join vendor v1 on v1.shortname = v.shortname " &
            '                   " order by v1.shortname::text; "

            sb.Append(String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text" &
                                  " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                  " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{0}$'   order by vendorname) " &
                                  "    select distinct v1.vendorcode::text || ' - ' || v1.vendorname as name, v1.vendorcode ,v1.vendorname::text from doc.vendordoc vd " &
                                  " inner join va on va.vendorcode = vd.vendorcode" &
                              " left join vendor v on v.vendorcode = va.vendorcode" &
                              " left join vendor v1 on v1.shortname = v.shortname" &
                              " order by v1.vendorname::text;", HelperClass1.UserInfo.userid))

            sb.Append(String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{0}$'   order by vendorname) " &
                                   "   select distinct v1.shortname::text from doc.vendordoc vd" &
                                   " inner join va on va.vendorcode = vd.vendorcode" &
                               " left join vendor v on v.vendorcode = va.vendorcode " &
                               " left join vendor v1 on v1.shortname = v.shortname " &
                               " order by v1.shortname::text;", HelperClass1.UserInfo.userid))


            sqlstr = " select distinct o.officersebname::text as spm from doc.vendordoc vd" &
                               " left join vendor v on v.vendorcode = vd.vendorcode " &
                               " left join officerseb o on o.ofsebid = v.ssmidpl " &
                               " order by o.officersebname::text;" &
                               " select distinct o.officersebname::text as pm from doc.vendordoc vd" &
                               " left join vendor v on v.vendorcode = vd.vendorcode " &
                               " left join officerseb o on o.ofsebid = v.pmid " &
                               " order by o.officersebname::text ;" &
                               " select dt.paramname as producttype from doc.paramdt dt" &
                               " where(dt.paramhdid = 3)" &
                               " order by ivalue;" &
                               " select dt.paramname as statusname from doc.paramdt dt" &
                               " where(dt.paramhdid = 2)" &
                               " order by ivalue;" &
                               " select sbuname from sbusap" &
                               " order by sbuname;" &
                               " select doctypename,id from doc.doctype" &
                               " order by doctypename;"
        End If

        DS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString & sqlstr, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "Vendor"
                DS.Tables(1).TableName = "ShortName"
                DS.Tables(2).TableName = "SPM"
                DS.Tables(3).TableName = "PM"
                DS.Tables(4).TableName = "ProductType"
                DS.Tables(5).TableName = "StatusName"
                DS.Tables(6).TableName = "SBU"
                DS.Tables(7).TableName = "DocType"
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

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        loaddata()
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
                            VendorBS = New BindingSource
                            ShortnameBS = New BindingSource
                            PMBS = New BindingSource
                            SPMBS = New BindingSource
                            ProductTypeBS = New BindingSource
                            StatusNameBS = New BindingSource
                            SBUBS = New BindingSource
                            DocTypeBS = New BindingSource
                            VendorBS.DataSource = DS.Tables("Vendor")
                            ShortnameBS.DataSource = DS.Tables("ShortName")
                            SPMBS.DataSource = DS.Tables("SPM")
                            PMBS.DataSource = DS.Tables("PM")
                            ProductTypeBS.DataSource = DS.Tables("ProductType")
                            StatusNameBS.DataSource = DS.Tables("StatusName")
                            SBUBS.DataSource = DS.Tables("SBU")
                            DocTypeBS.DataSource = DS.Tables("DocType")
                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click
        If myThread.IsAlive Then
            MessageBox.Show("Please wait until the current process is finished.")
            Exit Sub
        End If
        Dim mybutton = DirectCast(sender, Button)
        Select Case mybutton.Name
            'Case "Button1"
            '    Dim myform = New FormHelperMulti(VendorBS)
            '    myform.DataGridView1.Columns(0).DataPropertyName = "name"
            '    myform.Key = 1
            '    myform.ShowDialog()
            '    TextBox1.Text = myform.SelectedResult
            'Case "Button2"
            '    Dim myform = New FormHelperMulti(VendorBS)
            '    myform.DataGridView1.Columns(0).DataPropertyName = "name"
            '    myform.Key = 2
            '    myform.ShowDialog()
            '    TextBox2.Text = myform.SelectedResult
            'Case "Button3"
            '    Dim myform = New FormHelperMulti(ShortnameBS)
            '    myform.DataGridView1.Columns(0).DataPropertyName = "shortname"
            '    myform.Key = 0
            '    myform.ShowDialog()
            '    TextBox3.Text = myform.SelectedResult
            'Case "Button4"
            '    Dim myform = New FormHelperMulti(SPMBS)
            '    myform.DataGridView1.Columns(0).DataPropertyName = "spm"
            '    myform.Key = 0
            '    myform.ShowDialog()
            '    TextBox4.Text = myform.SelectedResult
            'Case "Button5"
            '    Dim myform = New FormHelperMulti(PMBS)
            '    myform.DataGridView1.Columns(0).DataPropertyName = "pm"
            '    myform.Key = 0
            '    myform.ShowDialog()
            '    TextBox5.Text = myform.SelectedResult
            'Case "Button6"
            '    Dim myform = New FormHelperMulti(ProductTypeBS)
            '    myform.DataGridView1.Columns(0).DataPropertyName = "producttype"
            '    myform.Key = 0
            '    myform.ShowDialog()
            '    TextBox6.Text = myform.SelectedResult
            'Case "Button7"
            '    Dim myform = New FormHelperMulti(StatusNameBS)
            '    myform.DataGridView1.Columns(0).DataPropertyName = "statusname"
            '    myform.Key = 0
            '    myform.ShowDialog()
            '    TextBox7.Text = myform.SelectedResult
            'Case "Button8"
            '    Dim myform = New FormHelperMulti(SBUBS)
            '    myform.DataGridView1.Columns(0).DataPropertyName = "sbuname"
            '    'myform.DataGridView1.MultiSelect = False
            '    myform.Key = 0
            '    myform.ShowDialog()
            '    TextBox8.Text = myform.SelectedResult
            Case "Button1"
                Dim myform = New FormHelperMulti(VendorBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "name"
                myform.Key = 1
                myform.ShowDialog()
                TextBox1.Text = TextBox1.Text + IIf(TextBox1.Text = "", "", ",") + myform.SelectedResult
            Case "Button2"
                Dim myform = New FormHelperMulti(VendorBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "name"
                myform.Key = 2
                myform.ShowDialog()
                TextBox2.Text = TextBox2.Text + IIf(TextBox2.Text = "", "", ",") + myform.SelectedResult
            Case "Button3"
                Dim myform = New FormHelperMulti(ShortnameBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "shortname"
                myform.Key = 0
                myform.ShowDialog()
                TextBox3.Text = TextBox3.Text + IIf(TextBox3.Text = "", "", ",") + myform.SelectedResult
            Case "Button4"
                Dim myform = New FormHelperMulti(SPMBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "spm"
                myform.Key = 0
                myform.ShowDialog()
                TextBox4.Text = TextBox4.Text + IIf(TextBox4.Text = "", "", ",") + myform.SelectedResult
            Case "Button5"
                Dim myform = New FormHelperMulti(PMBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "pm"
                myform.Key = 0
                myform.ShowDialog()
                TextBox5.Text = TextBox5.Text + IIf(TextBox5.Text = "", "", ",") + myform.SelectedResult
            Case "Button6"
                Dim myform = New FormHelperMulti(ProductTypeBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "producttype"
                myform.Key = 0
                myform.ShowDialog()
                TextBox6.Text = TextBox6.Text + IIf(TextBox6.Text = "", "", ",") + myform.SelectedResult
            Case "Button7"
                Dim myform = New FormHelperMulti(StatusNameBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "statusname"
                myform.Key = 0
                myform.ShowDialog()
                TextBox7.Text = TextBox7.Text + IIf(TextBox7.Text = "", "", ",") + myform.SelectedResult
            Case "Button8"
                Dim myform = New FormHelperMulti(SBUBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "sbuname"
                'myform.DataGridView1.MultiSelect = False
                myform.Key = 0
                myform.ShowDialog()
                TextBox8.Text = TextBox8.Text + IIf(TextBox8.Text = "", "", ",") + myform.SelectedResult
            Case "Button9"
                Dim myform = New FormHelperMulti(DocTypeBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "doctypename"
                myform.DataGridView1.MultiSelect = False
                myform.Key = 0
                myform.ShowDialog()
                TextBox9.Text = myform.SelectedResult
        End Select


    End Sub


    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        DateTimePicker1.Enabled = RadioButton1.Checked
        DateTimePicker2.Enabled = RadioButton2.Checked
        DateTimePicker3.Enabled = RadioButton2.Checked

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim mymessage As String = String.Empty
        Dim sqlstr As String = String.Empty
        Dim sb As New StringBuilder
        Dim latestdate As String = String.Empty
        Dim startdate As String = String.Empty
        Dim enddate As String = String.Empty
        Dim nonblankdoc As String = String.Empty
        Dim mytemplate As String = String.Format("{0}\SupplierDocumentTypeTemplate.xltm", HelperClass1.template)

        ToolStripStatusLabel1.Text = ""
        ToolStripStatusLabel2.Text = ""

        If TextBox9.Text <> "" Then
            Dim drv As DataRowView = DocTypeBS.Current
            If TextBox9.Text = drv.Row.Item("doctypename") Then
                docTypeId = drv.Row.Item("id")
            Else
                TextBox9.Focus()
                TextBox9.SelectAll()
                MessageBox.Show("Doc Type name is not available.")
                Exit Sub
            End If

        Else
            MessageBox.Show("Please select Document Type.")
            TextBox9.Focus()
            Exit Sub
        End If

        If HelperClass1.UserInfo.IsAdmin Then
            sb.Append("SELECT clt.docid, clt.vendorcode, v.vendorname::text AS vendorname, v.shortname::text AS shortname, mu1.username AS gsm, mu2.username AS spm, mu.username AS pm, d.id AS documentid, dt.doctypename, countbydoc,withoutdocument, dl.levelname, de.expireddate, d.docdate, d.docname, d.docext, d.uploaddate, d.userid, d.remarks, vr.version, (pt.payt::text || ' - '::text) || pt.details::text AS paymentterm, sc.leadtime, sc.sasl, q.nqsu, p.projectname, sa.auditby, sa.audittype, sa.auditgrade, sef.score, sif.myyear, sif.turnovery, sif.turnovery1, sif.turnovery2, sif.turnovery3, sif.turnovery4, sif.ratioy, sif.ratioy1, sif.ratioy2, sif.ratioy3, sif.ratioy4, scy.category, ps1.panelstatus AS panelstatus1, ps1.paneldescription AS paneldescription1, ps2.panelstatus AS panelstatus2, ps2.paneldescription AS paneldescription2, " &
                " CASE WHEN dt.id = 32 THEN 1 ELSE NULL::integer END AS generalcontract, CASE WHEN dt.id = 33 THEN 1 ELSE NULL::integer END AS qualityappendix, CASE WHEN dt.id = 35 THEN 1 ELSE NULL::integer END AS supplychainappendix,null::integer  AS countcontractisfull, " &
                " null::integer AS countcontractfull,  null::integer as countcontractnotfull, vg1.groupsbuname AS fp, vg2.groupsbuname AS cp, vg3.groupsbuname AS sp, vg4.groupsbuname AS mould, doc.groupact(vg1.groupsbuname::character varying, vg2.groupsbuname::character varying, vg3.groupsbuname::character varying, vg4.groupsbuname::character varying, si.fpcp::character varying) AS groupact, pdt.producttype, pr.paramname AS statusname, spp.rank,vsbu.sbuname ")

        Else
            ' o.officersebname AS spm, os.officersebname AS pm
            sb.Append(String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{0}$'   order by vendorname) " &
                "SELECT clt.docid, clt.vendorcode, v.vendorname::text AS vendorname, v.shortname::text AS shortname,  mu1.username AS gsm, mu2.username AS spm, mu.username AS pm, d.id AS documentid, dt.doctypename, countbydoc,withoutdocument, dl.levelname, de.expireddate, d.docdate, d.docname, d.docext, d.uploaddate, d.userid, d.remarks, vr.version, (pt.payt::text || ' - '::text) || pt.details::text AS paymentterm, sc.leadtime, sc.sasl, q.nqsu, p.projectname, sa.auditby, sa.audittype, sa.auditgrade, sef.score, sif.myyear, sif.turnovery, sif.turnovery1, sif.turnovery2, sif.turnovery3, sif.turnovery4, sif.ratioy, sif.ratioy1, sif.ratioy2, sif.ratioy3, sif.ratioy4, scy.category, ps1.panelstatus AS panelstatus1, ps1.paneldescription AS paneldescription1, ps2.panelstatus AS panelstatus2, ps2.paneldescription AS paneldescription2, " &
                " CASE WHEN dt.id = 32 THEN 1 ELSE NULL::integer END AS generalcontract, CASE WHEN dt.id = 33 THEN 1 ELSE NULL::integer END AS qualityappendix, CASE WHEN dt.id = 35 THEN 1 ELSE NULL::integer END AS supplychainappendix,null::integer  AS countcontractisfull, " &
                " null::integer AS countcontractfull,  null::integer as countcontractnotfull, vg1.groupsbuname AS fp, vg2.groupsbuname AS cp, vg3.groupsbuname AS sp, vg4.groupsbuname AS mould, doc.groupact(vg1.groupsbuname::character varying, vg2.groupsbuname::character varying, vg3.groupsbuname::character varying, vg4.groupsbuname::character varying, si.fpcp::character varying) AS groupact, pdt.producttype, pr.paramname AS statusname, spp.rank,vsbu.sbuname ", HelperClass1.UserInfo.userid))


        End If
       
        If RadioButton1.Checked Then           
            latestdate = "'" & Year(DateTimePicker1.Value) & "-12-31'"
            sb.Append(" from doc.doctypelatestallvendor(" & latestdate & "::date, " & docTypeId & ") clt")
        Else
            mytemplate = "\templates\SupplierDocumentTypeDetail.xltm"
            startdate = "'" & Year(DateTimePicker2.Value) & "-1-1'"
            enddate = "'" & Year(DateTimePicker3.Value) & "-12-31'"
            'nonblankdoc = " and not d.id isnull "
            sb.Append(" from doc.doctypedetailallvendor(" & startdate & "::date," & enddate & "::date," & docTypeId & ") clt")
        End If

        sb.Append(" left join doc.document d on d.id = clt.docid ")
        If HelperClass1.UserInfo.IsAdmin Then
            sb.Append(" LEFT JOIN vendor v ON v.vendorcode = clt.vendorcode")
        Else
            sb.Append(" inner JOIN va  ON va.vendorcode = clt.vendorcode")
            sb.Append(" LEFT JOIN vendor v ON v.vendorcode = va.vendorcode")
        End If



        'sb.Append(" LEFT JOIN doc.version vr ON vr.documentid = d.id" &
        '          " LEFT JOIN doc.generalcontract gt ON gt.documentid = d.id" &
        '          " LEFT JOIN paymentterm pt ON pt.paymenttermid = gt.paymentcode" &
        '          " LEFT JOIN doc.supplychain sc ON sc.documentid = d.id" &
        '          " LEFT JOIN doc.qualityappendix q ON q.documentid = d.id" &
        '          " LEFT JOIN doc.project p ON p.documentid = d.id" &
        '          " LEFT JOIN doc.socialaudit sa ON sa.documentid = d.id" &
        '          " LEFT JOIN doc.sef sef ON sef.documentid = d.id" &
        '          " LEFT JOIN doc.sif sif ON sif.documentid = d.id" &
        '          " LEFT JOIN doc.doctype dt ON dt.id = d.doctypeid" &
        '          " LEFT JOIN doc.doclevel dl ON dl.id = d.doclevelid" &
        '          " LEFT JOIN doc.docexpired de ON de.documentid = d.id" &
        '          " LEFT JOIN supplierspanel spp ON spp.vendorcode = clt.vendorcode" &
        '          " LEFT JOIN supplierscategory scy ON scy.supplierscategoryid = spp.supplierscategoryid" &
        '          " LEFT JOIN doc.panelstatus ps1 ON ps1.id = spp.fp" &
        '          " LEFT JOIN doc.panelstatus ps2 ON ps2.id = spp.cp" &
        '          " LEFT JOIN officerseb o ON o.ofsebid = v.ssmidpl" &
        '          " LEFT JOIN officerseb os ON os.ofsebid = v.pmid" &
        '          " LEFT JOIN doc.tvendorgroupsbu vg1 ON vg1.vendorcode = v.vendorcode AND vg1.groupsbuname = 'FP'::text" &
        '          " LEFT JOIN doc.tvendorgroupsbu vg2 ON vg2.vendorcode = v.vendorcode AND vg2.groupsbuname = 'CP'::text" &
        '          " LEFT JOIN doc.tvendorgroupsbu vg3 ON vg3.vendorcode = v.vendorcode AND vg3.groupsbuname = 'SP'::text" &
        '          " LEFT JOIN doc.tvendorgroupsbu vg4 ON vg4.vendorcode = v.vendorcode AND vg4.groupsbuname = 'MOULD'::text" &
        '          " LEFT JOIN doc.vendorstatus vs ON vs.vendorcode = v.vendorcode" &
        '          " LEFT JOIN doc.paramhd ph ON ph.paramname::text = 'vendorstatus'::text" &
        '          " LEFT JOIN doc.paramdt pr ON pr.ivalue = vs.status AND pr.paramhdid = ph.paramhdid" &
        '          " LEFT JOIN doc.shortnameinfo si ON si.shortname = v.shortname::text" &
        '          " LEFT JOIN ( SELECT pd.ivalue, pd.paramname AS producttype" &
        '          " FROM doc.paramdt pd" &
        '          " LEFT JOIN doc.paramhd ph ON ph.paramhdid = pd.paramhdid" &
        '          " WHERE ph.paramname::text = 'producttype'::text) pdt ON pdt.ivalue = vs.producttypeid" &
        '          " left join doc.vendorsbu vsbu on vsbu.vendorcode = clt.vendorcode" &
        '          " WHERE(Not clt.vendorcode Is NULL)")
        'viewvendorpm replace viewvendorfamilypm
        sb.Append(" LEFT JOIN doc.viewvendorpm vfp ON vfp.vendorcode = v.vendorcode" &
                  " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid" &
                  " LEFT JOIN masteruser mu ON mu.id = os.muid" &
                  " LEFT JOIN officerseb o ON o.ofsebid = os.parent" &
                  " LEFT JOIN masteruser mu2 ON mu2.id = o.muid" &
                  " LEFT JOIN doc.vendorgsm gsm ON gsm.vendorcode = v.vendorcode" &
                  " LEFT JOIN officerseb o1 ON o1.ofsebid = gsm.gsmid" &
                  " LEFT JOIN masteruser mu1 ON mu1.id = o1.muid" &
                  " LEFT JOIN doc.version vr ON vr.documentid = d.id" &
                  " LEFT JOIN doc.generalcontract gt ON gt.documentid = d.id" &
                  " LEFT JOIN paymentterm pt ON pt.paymenttermid = gt.paymentcode" &
                  " LEFT JOIN doc.supplychain sc ON sc.documentid = d.id" &
                  " LEFT JOIN doc.qualityappendix q ON q.documentid = d.id" &
                  " LEFT JOIN doc.project p ON p.documentid = d.id" &
                  " LEFT JOIN doc.socialaudit sa ON sa.documentid = d.id" &
                  " LEFT JOIN doc.sef sef ON sef.documentid = d.id" &
                  " LEFT JOIN doc.sif sif ON sif.documentid = d.id" &
                  " LEFT JOIN doc.doctype dt ON dt.id = d.doctypeid" &
                  " LEFT JOIN doc.doclevel dl ON dl.id = d.doclevelid" &
                  " LEFT JOIN doc.docexpired de ON de.documentid = d.id" &
                  " LEFT JOIN supplierspanel spp ON spp.vendorcode = clt.vendorcode" &
                  " LEFT JOIN supplierscategory scy ON scy.supplierscategoryid = spp.supplierscategoryid" &
                  " LEFT JOIN doc.panelstatus ps1 ON ps1.id = spp.fp" &
                  " LEFT JOIN doc.panelstatus ps2 ON ps2.id = spp.cp" &
                  " LEFT JOIN doc.tvendorgroupsbu vg1 ON vg1.vendorcode = v.vendorcode AND vg1.groupsbuname = 'FP'::text" &
                  " LEFT JOIN doc.tvendorgroupsbu vg2 ON vg2.vendorcode = v.vendorcode AND vg2.groupsbuname = 'CP'::text" &
                  " LEFT JOIN doc.tvendorgroupsbu vg3 ON vg3.vendorcode = v.vendorcode AND vg3.groupsbuname = 'SP'::text" &
                  " LEFT JOIN doc.tvendorgroupsbu vg4 ON vg4.vendorcode = v.vendorcode AND vg4.groupsbuname = 'MOULD'::text" &
                  " LEFT JOIN doc.vendorstatus vs ON vs.vendorcode = v.vendorcode" &
                  " LEFT JOIN doc.paramhd ph ON ph.paramname::text = 'vendorstatus'::text" &
                  " LEFT JOIN doc.paramdt pr ON pr.ivalue = vs.status AND pr.paramhdid = ph.paramhdid" &
                  " LEFT JOIN doc.shortnameinfo si ON si.shortname = v.shortname::text" &
                  " LEFT JOIN ( SELECT pd.ivalue, pd.paramname AS producttype" &
                  " FROM doc.paramdt pd" &
                  " LEFT JOIN doc.paramhd ph ON ph.paramhdid = pd.paramhdid" &
                  " WHERE ph.paramname::text = 'producttype'::text) pdt ON pdt.ivalue = vs.producttypeid" &
                  " left join doc.vendorsbu vsbu on vsbu.vendorcode = clt.vendorcode" &
                  " WHERE(Not clt.vendorcode Is NULL)")

        If TextBox1.Text <> "" Then
            sb.Append(" and '" & TextBox1.Text.Replace("'", "''") & "' ~ clt.vendorcode::text")
        End If

        If TextBox2.Text <> "" Then
            sb.Append(" and '" & TextBox2.Text.Replace("'", "''") & "' ~ v.vendorname::text")
        End If

        If TextBox3.Text <> "" Then
            sb.Append(" and '" & TextBox3.Text.Replace("'", "''") & "' ~ v.shortname::text")
        End If

        'If TextBox4.Text <> "" Then
        '    sb.Append(" and '" & TextBox4.Text.Replace("'", "''") & "' ~ o.officersebname::text")
        'End If

        'If TextBox5.Text <> "" Then
        '    sb.Append(" and '" & TextBox5.Text.Replace("'", "''") & "' ~ os.officersebname::text")
        'End If
        If TextBox4.Text <> "" Then
            sb.Append(" and '" & TextBox4.Text.Replace("'", "''") & "' ~ mu2.username")
        End If

        If TextBox5.Text <> "" Then
            sb.Append(" and '" & TextBox5.Text.Replace("'", "''") & "' ~ mu.username")
        End If
        If TextBox6.Text <> "" Then
            'sb.Append(" and '" & TextBox6.Text.Replace("'", "''") & "' ~ pdt.producttype::text")
            Dim ProductTypeSB As New StringBuilder
            Dim myProdType = TextBox6.Text.ToString.Split(",")
            For i = 0 To myProdType.Length - 1
                If ProductTypeSB.Length = 0 Then
                    ProductTypeSB.Append(" and (")
                Else
                    ProductTypeSB.Append(" or ")
                End If

                ProductTypeSB.Append(" pdt.producttype::text ='" & myProdType(i).Replace("'", "''") & "'")
            Next
            ProductTypeSB.Append(")")
            sb.Append(ProductTypeSB.ToString)
        End If

        If TextBox7.Text <> "" Then
            sb.Append(" and '" & TextBox7.Text.Replace("'", "''") & "' ~ pr.paramname::text")
        End If

        If TextBox8.Text <> "" Then
            'sb.Append(" and '" & TextBox8.Text.Replace("'", "''") & "' ~ vsbu.sbuname::text")
            Dim SBUNameSB As New StringBuilder
            Dim mysbu = TextBox8.Text.ToString.Split(",")
            For i = 0 To mysbu.Length - 1
                If SBUNameSB.Length = 0 Then
                    SBUNameSB.Append(" and (")
                Else
                    SBUNameSB.Append(" or ")
                End If

                SBUNameSB.Append(" vsbu.sbuname ~'" & mysbu(i).Replace("'", "''") & "'")
            Next
            SBUNameSB.Append(")")
            sb.Append(SBUNameSB.ToString)
            'sb.Append(" and vsbu.sbuname ~'" & TextBox8.Text.Replace("'", "''") & "'")
        End If
       
        sb.Append(nonblankdoc)
        sb.Append(" ORDER BY v.shortname::text, dt.doctypename;")
        Dim mysaveform As New SaveFileDialog
        'mysaveform.FileName = String.Format("SupplierDocumentFilterReport{0:yyyyMMdd}.xlsm", Date.Today)
        mysaveform.FileName = String.Format("{1}{0:yyyyMMdd}.xlsm", Date.Today, TextBox9.Text).Replace(":", "").Replace(" ", "")
        mysaveform.DefaultExt = ".xlsm"

        sqlstr = sb.ToString
        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 2 'because hidden

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, mytemplate)
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


        owb.Worksheets(2).select()


        'check availability data


        Dim osheet = owb.Worksheets(2)
        Dim orange = osheet.Range("A2")
        If osheet.cells(2, 2).text.ToString = "" Then

            Err.Raise(100, Description:="Data not available.")
        End If
        osheet.name = "RawData"

        owb.Names.Add("rawdata", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C2),COUNTA('RawData'!R1))")

        owb.Worksheets(1).select()
        osheet = owb.Worksheets(1)
        'oXl.Run("ShowFG")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        'oXl.Run("ShowCP")
        osheet.PivotTables("PivotTable2").PivotCache.Refresh()
        Try
            With osheet.PivotTables("PivotTable2").PivotFields("producttype")
                .PivotItems("FP").Visible = False
            End With
        Catch ex As Exception

        End Try
        
        'oXl.Run("OpenTabs")

        'osheet = owb.Worksheets(2)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(3)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(4)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(5)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(6)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(1)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(7)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()


        oXl.Run("BACK")
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

    End Sub
End Class