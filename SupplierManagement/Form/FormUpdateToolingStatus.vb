Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class FormUpdateToolingStatus
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents ToolingListBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim myReportType As ReportType

    'Tooling ID
    'Tooling Status
    'Asset Purchase ID
    'Project Code
    'Project Name
    'Short Name
    'Type of Investment
    'AEB No.
    'Investment Order No.
    'Finance Asset No
    'Tooling PO No.
    Dim myarray() = {"", "lower(tl.toolinglistid)", "lower(doc.getstatusname(tm.status))", "lower(ap.assetpurchaseid)", "lower(projectcode)", "lower(projectname)", "v.vendorcode::text", "lower(v.vendorname::text)", "lower(v.shortname::text)", "lower(doc.gettypeofinvestmentname(ap.typeofinvestment::int))", "lower(aeb)", "lower(investmentorderno)", "lower(financeassetno)", "lower(toolingpono)"}
    Private Sqlstr As String
    Dim SqlstrReport As String
    Dim ToolingStatusList As List(Of ToolingStatus)
    Dim ToolingStatusBS As BindingSource
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        toolingstatuslist = New List(Of ToolingStatus)
        toolingstatuslist.Add(New ToolingStatus(1, "Active"))
        toolingstatuslist.Add(New ToolingStatus(2, "EOL,Scrapped"))
        ToolingStatusList.Add(New ToolingStatus(3, "Transfer to other project"))
        myReportType = ReportType.ToolingList
        ComboBox1.SelectedIndex = 8
        ComboBox2.SelectedIndex = 5
        ToolStrip1.Visible = Not HelperClass1.UserInfo.IsFinance
        If HelperClass1.UserInfo.IsFinance Then
            DataGridView1.ReadOnly = True
        End If
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByVal sqlstr, ByVal Report, ByVal FileName)

        ' This call is required by the designer.
        InitializeComponent()
        Me.Sqlstr = sqlstr
        Me.myReportType = Report
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
        mysaveform.FileName = String.Format("{0}Report{1:yyyyMMdd}.xlsx", myReportType.ToString, Date.Today)

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

        If myReportType = ReportType.AssetPurchase Then
            sender.Columns("R:R").NumberFormat = "0%"
        Else
            sender.Columns("AH:AH").NumberFormat = "0%"
        End If
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
                            ToolingStatusBS = New BindingSource
                            ToolingStatusBS.DataSource = toolingstatuslist
                            DS.Tables(0).TableName = "ToolingList"

                            ToolingListBS.DataSource = DS.Tables(0)

                            DataGridView1.AutoGenerateColumns = False
                            If IsNothing(DirectCast(DataGridView1.Columns("ToolingStatus"), DataGridViewComboBoxColumn).DataSource) Then
                                DirectCast(DataGridView1.Columns("ToolingStatus"), DataGridViewComboBoxColumn).DataSource = ToolingStatusBS
                                DirectCast(DataGridView1.Columns("ToolingStatus"), DataGridViewComboBoxColumn).DisplayMember = "toolingstatusname"
                                DirectCast(DataGridView1.Columns("ToolingStatus"), DataGridViewComboBoxColumn).ValueMember = "toolingstatusid"
                                DirectCast(DataGridView1.Columns("ToolingStatus"), DataGridViewComboBoxColumn).DataPropertyName = "toolingstatus"
                            End If


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
                            sb.Append(String.Format("and {0} like '%{1}%'", myarray(ComboBox1.SelectedIndex), TextBox1.Text))
                        End If
                        If ComboBox2.SelectedIndex > 0 And TextBox2.Text <> "" Then
                            sb.Append(String.Format("and {0} like '%{1}%'", myarray(ComboBox2.SelectedIndex), TextBox2.Text))
                        End If
                        If ComboBox3.SelectedIndex > 0 And TextBox3.Text <> "" Then
                            sb.Append(String.Format("and {0} like '%{1}%'", myarray(ComboBox3.SelectedIndex), TextBox3.Text))
                        End If
                        If CheckBox1.Checked Then
                            sb.Append(String.Format("and ap.applicantdate >= '{0:yyyy-MM-dd}' and applicantdate <= '{1:yyyy-MM-dd}'", DateTimePicker2.Value.Date, DateTimePicker3.Value.Date))
                        End If
                        If CheckBox3.Checked Then
                            sb.Append(String.Format("and ap.sapcapdate >= '{0:yyyy-MM-dd}' and sapcapdate <= '{1:yyyy-MM-dd}'", DateTimePicker4.Value.Date, DateTimePicker5.Value.Date))
                        End If
                        If CheckBox2.Checked Then
                            sb.Append(String.Format("and tm.commontool"))
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
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        SqlstrReport = String.Format("with inv as( select id,count(invoiceid) as countinv from (" &
                       " select   ap.id,ap.assetpurchaseid,tp.invoiceid,sum(tp.invoiceamount * exrate) as invoiceamount" &
                       " from doc.assetpurchase ap" &
                       " left join  doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                       " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                       " group by ap.id,invoiceid" &
                       " order by ap.id) as invoice" &
                       " group by id )  ," &
                       " toolpay as (select tl.toolinglistid,tl.assetpurchaseid," &
                       " sum(tp.invoiceamount * exrate) as invoiceamount" &
                       " from doc.toolinglist tl" &
                       " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                       " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid" &
                       " where not invoiceid isnull" &
                       " group by tl.toolinglistid,tl.assetpurchaseid )" &
                       " select tl.toolinglistid,doc.getstatusname(tm.status) as status,tm.commontool, ap.assetpurchaseid,tp.projectcode,tp.projectname,tp.ppps,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text," &
                       " v.vendorname::text,v.shortname,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount," &
                       " tl.sebmodelno,tl.suppliermodelreference,tl.suppliermoldno,tl.toolsdescription,tl.material,tl.cavities,tl.numberoftools,tl.dailycaps,tl.cost,tl.purchasedate,tl.location,tl.comments,inv.countinv," &
                       " doc.getinvoiceno(ap.id) as invoiceno, tpy.invoiceamount,(1- (tl.cost - tpy.invoiceamount) / case tl.cost when 0 then 1 else tl.cost end )  as paid,(tl.cost - tpy.invoiceamount) as balance,  ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate " &
                       " from  doc.toolingproject tp " &
                       " left join doc.assetpurchase ap on ap.projectid =  tp.id " &
                       " left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                       " left join doc.toolinglistmst tm on tm.toolinglistid = tl.toolinglistid" &
                       " left join vendor v on v.vendorcode = ap.vendorcode" &
                       " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid" &
                       " left join family f on f.familyid = tp.familyid " &
                       " left join sbusap s on s.sbuid = fgs.sbusapid" &
                       " left join toolpay tpy on tpy.toolinglistid = tl.toolinglistid and tpy.assetpurchaseid = ap.id" &
                       " left join inv on inv.id = ap.id" &
                       " where not ap.id isnull and not tl.toolinglistid isnull {0} order by tl.toolinglistid ;", sb.ToString)
        If Not HelperClass1.UserInfo.IsAdmin And Not (HelperClass1.UserInfo.IsFinance) Then
            SqlstrReport = String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{1}$'   order by vendorname) " &
                        " ,inv as( select id,count(invoiceid) as countinv from (" &
                       " select   ap.id,ap.assetpurchaseid,tp.invoiceid,sum(tp.invoiceamount * exrate) as invoiceamount" &
                       " from doc.assetpurchase ap" &
                       " left join  doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                       " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                       " group by ap.id,invoiceid" &
                       " order by ap.id) as invoice" &
                       " group by id )  ," &
                       " toolpay as (select tl.toolinglistid,tl.assetpurchaseid," &
                       " sum(tp.invoiceamount * exrate) as invoiceamount" &
                       " from doc.toolinglist tl" &
                       " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                       " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid" &
                       " where not invoiceid isnull" &
                       " group by tl.toolinglistid,tl.assetpurchaseid )" &
                       " select tl.toolinglistid,doc.getstatusname(tm.status) as status,tm.commontool, ap.assetpurchaseid,tp.projectcode,tp.projectname,tp.ppps,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text," &
                       " v.vendorname::text,v.shortname,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount," &
                       " tl.sebmodelno,tl.suppliermodelreference,tl.suppliermoldno,tl.toolsdescription,tl.material,tl.cavities,tl.numberoftools,tl.dailycaps,tl.cost,tl.purchasedate,tl.location,tl.comments,inv.countinv," &
                       " doc.getinvoiceno(ap.id) as invoiceno, tpy.invoiceamount,(1- (tl.cost - tpy.invoiceamount) / case tl.cost when 0 then 1 else tl.cost end )  as paid,(tl.cost - tpy.invoiceamount) as balance,  ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate " &
                       " from  doc.toolingproject tp " &
                       " left join doc.assetpurchase ap on ap.projectid =  tp.id " &
                       " left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                       " left join doc.toolinglistmst tm on tm.toolinglistid = tl.toolinglistid" &
                       " inner join va on va.vendorcode = ap.vendorcode " &
                       " left join vendor v on v.vendorcode = va.vendorcode" &
                       " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid" &
                       " left join family f on f.familyid = tp.familyid " &
                       " left join sbusap s on s.sbuid = fgs.sbusapid" &
                        " left join toolpay tpy on tpy.toolinglistid = tl.toolinglistid and tpy.assetpurchaseid = ap.id" &
                       " left join inv on inv.id = ap.id" &
                       " where not ap.id isnull and not tl.toolinglistid isnull {0} order by tl.toolinglistid ;", sb.ToString, HelperClass1.UserInfo.userid)
        End If
        myReportType = ReportType.ToolingList
        loadReport()
    End Sub

    Private Sub Button3_Clickori(ByVal sender As System.Object, ByVal e As System.EventArgs)

        SqlstrReport = String.Format("with inv as( select id,count(invoiceid) as countinv from (" &
                       " select   ap.id,ap.assetpurchaseid,tp.invoiceid,sum(tp.invoiceamount * exrate) as invoiceamount" &
                       " from doc.assetpurchase ap" &
                       " left join  doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                       " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                       " group by ap.id,invoiceid" &
                       " order by ap.id) as invoice" &
                       " group by id )  ," &
                       " toolpay as (select tl.toolinglistid," &
                       " sum(tp.invoiceamount * exrate) as invoiceamount" &
                       " from doc.toolinglist tl" &
                       " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                       " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid" &
                       " where not invoiceid isnull" &
                       " group by tl.toolinglistid)" &
                       " select tl.toolinglistid,doc.getstatusname(tm.status) as status,tm.commontool, ap.assetpurchaseid,tp.projectcode,tp.projectname,tp.ppps,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text," &
                       " v.vendorname::text,v.shortname,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount," &
                       " tl.sebmodelno,tl.suppliermodelreference,tl.suppliermoldno,tl.toolsdescription,tl.material,tl.cavities,tl.numberoftools,tl.dailycaps,tl.cost,tl.purchasedate,tl.location,tl.comments,inv.countinv," &
                       " doc.getinvoiceno(ap.id) as invoiceno, tpy.invoiceamount,(1- (tl.cost - tpy.invoiceamount) / case tl.cost when 0 then 1 else tl.cost end )  as paid,(tl.cost - tpy.invoiceamount) as balance,  ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate " &
                       " from  doc.toolingproject tp " &
                       " left join doc.assetpurchase ap on ap.projectid =  tp.id " &
                       " left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                       " left join doc.toolinglistmst tm on tm.toolinglistid = tl.toolinglistid" &
                       " left join vendor v on v.vendorcode = ap.vendorcode" &
                       " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid" &
                       " left join family f on f.familyid = tp.familyid " &
                       " left join sbusap s on s.sbuid = fgs.sbusapid" &
                       " left join toolpay tpy on tpy.toolinglistid = tl.toolinglistid" &
                       " left join inv on inv.id = ap.id" &
                       " where not ap.id isnull and not tl.toolinglistid isnull {0} order by tl.toolinglistid ;", sb.ToString)
        If Not HelperClass1.UserInfo.IsAdmin Then
            SqlstrReport = String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{1}$'   order by vendorname) " &
                        " ,inv as( select id,count(invoiceid) as countinv from (" &
                       " select   ap.id,ap.assetpurchaseid,tp.invoiceid,sum(tp.invoiceamount * exrate) as invoiceamount" &
                       " from doc.assetpurchase ap" &
                       " left join  doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                       " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                       " group by ap.id,invoiceid" &
                       " order by ap.id) as invoice" &
                       " group by id )  ," &
                       " toolpay as (select tl.toolinglistid," &
                       " sum(tp.invoiceamount * exrate) as invoiceamount" &
                       " from doc.toolinglist tl" &
                       " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                       " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid" &
                       " where not invoiceid isnull" &
                       " group by tl.toolinglistid)" &
                       " select tl.toolinglistid,doc.getstatusname(tm.status) as status,tm.commontool, ap.assetpurchaseid,tp.projectcode,tp.projectname,tp.ppps,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text," &
                       " v.vendorname::text,v.shortname,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount," &
                       " tl.sebmodelno,tl.suppliermodelreference,tl.suppliermoldno,tl.toolsdescription,tl.material,tl.cavities,tl.numberoftools,tl.dailycaps,tl.cost,tl.purchasedate,tl.location,tl.comments,inv.countinv," &
                       " doc.getinvoiceno(ap.id) as invoiceno, tpy.invoiceamount,(1- (tl.cost - tpy.invoiceamount) / case tl.cost when 0 then 1 else tl.cost end )  as paid,(tl.cost - tpy.invoiceamount) as balance,  ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate " &
                       " from  doc.toolingproject tp " &
                       " left join doc.assetpurchase ap on ap.projectid =  tp.id " &
                       " left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                       " left join doc.toolinglistmst tm on tm.toolinglistid = tl.toolinglistid" &
                       " inner join va on va.vendorcode = ap.vendorcode " &
                       " left join vendor v on v.vendorcode = va.vendorcode" &
                       " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid" &
                       " left join family f on f.familyid = tp.familyid " &
                       " left join sbusap s on s.sbuid = fgs.sbusapid" &
                       " left join toolpay tpy on tpy.toolinglistid = tl.toolinglistid" &
                       " left join inv on inv.id = ap.id" &
                       " where not ap.id isnull and not tl.toolinglistid isnull {0} order by tl.toolinglistid ;", sb.ToString, HelperClass1.UserInfo.userid)
        End If
        myReportType = ReportType.ToolingList
        loadReport()
    End Sub
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        DateTimePicker2.Enabled = CheckBox1.Checked
        DateTimePicker3.Enabled = CheckBox1.Checked
    End Sub
    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        DateTimePicker4.Enabled = CheckBox3.Checked
        DateTimePicker5.Enabled = CheckBox3.Checked
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

        Dim sqlstr = String.Format("select distinct tl.toolinglistid, tm.commontool, tm.status as toolingstatus,doc.getstatusname(tm.status) as statusname,ap.id, ap.assetpurchaseid,tp.projectcode,tp.projectname,ap.applicantname,ap.vendorcode::text,v.vendorname::text,v.shortname::text,ap.applicantdate::text,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.applicantdate,doc.getinvoiceno(ap.id) as invoiceno " &
                  " from  doc.toolingproject tp" &
                  " left join doc.assetpurchase ap on ap.projectid =  tp.id" &
                  " left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                  " left join doc.toolinglistmst tm on tm.toolinglistid = tl.toolinglistid" &
                  " left join vendor v on v.vendorcode = ap.vendorcode" &
                  " where not tl.toolinglistid isnull {0} order by ap.id desc;", sb.ToString.ToLower)
        If Not HelperClass1.UserInfo.IsAdmin And Not (HelperClass1.UserInfo.IsFinance) Then
            sqlstr = String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{1}$'   order by vendorname) " &
                  "select distinct tl.toolinglistid, tm.commontool, tm.status as toolingstatus,doc.getstatusname(tm.status) as statusname,ap.id, ap.assetpurchaseid,tp.projectcode,tp.projectname,ap.applicantname,ap.vendorcode::text,v.vendorname::text,v.shortname::text,ap.applicantdate::text,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.applicantdate,doc.getinvoiceno(ap.id) as invoiceno " &
                  " from  doc.toolingproject tp" &
                  " left join doc.assetpurchase ap on ap.projectid =  tp.id" &
                  " left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                  " left join doc.toolinglistmst tm on tm.toolinglistid = tl.toolinglistid" &
                  " inner join va on va.vendorcode = ap.vendorcode" &
                  " left join vendor v on v.vendorcode = va.vendorcode" &
                  " where not tl.toolinglistid isnull {0} order by ap.id desc;", sb.ToString.ToLower, HelperClass1.UserInfo.userid)
        End If
        If DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "ToolingListDT"

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

    Private Sub FormUpdateToolingStatus_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not HelperClass1.UserInfo.IsFinance Then
            Me.Validate()
            If Not IsNothing(DS) Then
                Dim abc = DS.GetChanges()
                If Not IsNothing(abc) Then
                    Select Case MessageBox.Show("Save unsaved records?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                        Case Windows.Forms.DialogResult.Yes
                            If Me.Validate Then
                                ToolStripButton2.PerformClick()
                            Else
                                e.Cancel = True
                            End If

                        Case Windows.Forms.DialogResult.Cancel
                            e.Cancel = True
                    End Select
                End If
            End If

        End If
        
       
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Try
            ToolingListBS.EndEdit()
            If Me.Validate Then
                Try
                    'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                    Dim ds2 As DataSet
                    ds2 = DS.GetChanges
                    'update statusname
                    'For i = 0 To ds2.Tables(0).Rows.Count - 1
                    '    ds2.Tables(0).Rows(i).Item("statusname") = getstatus(ds2.Tables(0).Rows(i).Item("toolingstatus"))
                    '    ds2.Tables(0).Rows(i).EndEdit()
                    'Next


                    If Not IsNothing(ds2) Then
                        Dim mymessage As String = String.Empty
                        Dim ra As Integer
                        Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                        If Not DbAdapter1.ToolingStatusTX(Me, mye) Then
                            MessageBox.Show(mye.message)
                            Exit Sub
                        End If
                        DS.Merge(ds2)
                        DS.AcceptChanges()
                        DataGridView1.Invalidate()
                        MessageBox.Show("Saved.")
                    End If
                Catch ex As Exception
                    MessageBox.Show(" Error:: " & ex.Message)
                End Try
            End If
            DataGridView1.Invalidate()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Function getstatus(ByVal data As Object) As String
        Dim myret As String = String.Empty

        Select Case data
            Case 1
                myret = "Active"
            Case 2
                myret = "EOL, Scrapped"
            Case 3
                myret = "Transfer to other project"

        End Select
        Return myret
    End Function

    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        'ds2.Tables(0).Rows(i).Item("statusname") = getstatus(ds2.Tables(0).Rows(i).Item("toolingstatus"))
        Dim drv As DataRowView = ToolingListBS.Current
        DS.Tables(0).Rows(e.RowIndex).Item("statusname") = getstatus(drv.Row.Item("toolingstatus"))
    End Sub


End Class

Public Class ToolingStatus
    Public Property toolingstatusid As Integer
    Public Property toolingstatusname As String

    Public Sub New(ByVal _toolingstatusid As Integer, ByVal _toolingstatusname As String)
        Me.toolingstatusid = _toolingstatusid
        Me.toolingstatusname = _toolingstatusname
    End Sub
End Class