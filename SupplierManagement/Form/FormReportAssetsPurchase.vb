Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Enum ReportType
    AssetPurchase = 0
    ToolingList = 1
    ToolingHistory = 2
    AssetPurchaseAgreement = 3
End Enum
Public Class FormReportAssetsPurchase
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents ToolingListBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim myReportType As ReportType
    Dim ShortNameBS As BindingSource
    Dim AEBBS As BindingSource

    Dim myarray() = {"", "lower(ap.assetpurchaseid)", "lower(projectcode)", "lower(projectname)", "lower(familyname)", "lower(sbuname)", "lower(applicantname)",
                     "v.vendorcode::text", "lower(v.vendorname::text)", "lower(v.shortname3::text)", "lower(doc.gettypeofinvestmentname(ap.typeofinvestment::int))", "lower(aeb)", "lower(doc.getinvoiceno(ap.id::bigint))", "lower(investmentorderno)", "lower(financeassetno)", "lower(toolingpono)", "lower(creator)", "av.agreement::text", "adt.material::text", "lower(ap.paymententity)", "ap.toolingsupplier"}
    Private Sqlstr As String
    Dim SqlstrReport As String
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByVal sqlstr, ByVal Report, ByVal FileName)

        ' This call is required by the designer.
        InitializeComponent()
        Me.Sqlstr = sqlstr
        Me.myReportType = Report
        ' Add any initialization after the InitializeComponent() call.
        ComboBox1.SelectedIndex = 9
        ComboBox2.SelectedIndex = 11 '3
        ComboBox3.SelectedIndex = 1 '6
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False

    End Sub
    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        ProgressReport(7, "InitFilter")

        'Sqlstr = String.Format(SqlstrReport, sb.ToString.ToLower)
        Sqlstr = SqlstrReport
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
            sender.Columns("T:T").NumberFormat = "0%"
        ElseIf myReportType = ReportType.ToolingList Then
            sender.Columns("AF:AF").NumberFormat = "0%"
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
                            DS.Tables(0).TableName = "ToolingList"

                            ToolingListBS.DataSource = DS.Tables(0)

                            DataGridView1.AutoGenerateColumns = False
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
                            sb.Append(String.Format(" and {0} like '%{1}%'", myarray(ComboBox1.SelectedIndex), TextBox1.Text.Replace("\", "\\").ToLower))
                        End If
                        If ComboBox2.SelectedIndex > 0 And TextBox2.Text <> "" Then
                            sb.Append(String.Format(" and {0} like '%{1}%'", myarray(ComboBox2.SelectedIndex), TextBox2.Text.Replace("\", "\\").ToLower))
                        End If
                        If ComboBox3.SelectedIndex > 0 And TextBox3.Text <> "" Then
                            sb.Append(String.Format(" and {0} like '%{1}%'", myarray(ComboBox3.SelectedIndex), TextBox3.Text.Replace("\", "\\").ToLower))
                        End If
                        If ComboBox4.SelectedIndex > 0 Then
                            sb.Append(String.Format(" and paymentmethodid = {0}", ComboBox4.SelectedIndex))
                        End If

                        If CheckBox1.Checked Then
                            sb.Append(String.Format(" and ap.applicantdate >= '{0:yyyy-MM-dd}' and applicantdate <= '{1:yyyy-MM-dd}'", DateTimePicker2.Value.Date, DateTimePicker3.Value.Date))
                        End If
                        If CheckBox2.Checked Then
                            sb.Append(String.Format(" and ap.sapcapdate >= '{0:yyyy-MM-dd}' and sapcapdate <= '{1:yyyy-MM-dd}'", DateTimePicker4.Value.Date, DateTimePicker5.Value.Date))
                        End If
                        If CheckBox3.Checked Then
                            sb.Append(String.Format(" and av.startdate >= '{0:yyyy-MM-dd}' and enddate <= '{1:yyyy-MM-dd}'", DateTimePicker6.Value.Date, DateTimePicker1.Value.Date))
                        End If
                        If CheckBox4.Checked Then
                            sb.Append(String.Format(" and ast.doctypeid = {0}", ComboBox5.SelectedIndex + 1))
                        End If
                    Case 8
                        Button7.Enabled = True
                        Button8.Enabled = True
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


    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        
        SqlstrReport = String.Format("with tl as (" &
                                     " select assetpurchaseid,count(0) as totalofnotoolings,sum(cost) as totaltoolingcost from doc.toolinglist " &
                                     " group by assetpurchaseid" &
                                     ")," &
                                     " inv as(" &
                                     " select assetpurchaseid,sum(tp.invoiceamount * tp.exrate) as totalinvoiceamount" &
                                     " from doc.toolinglist tl" &
                                     " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                                     " group by assetpurchaseid" &
                                     " ) " &
                                     " select ap.assetpurchaseid,tp.projectcode,tp.projectname,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text,v.vendorname::text,v.shortname3::text as shortname,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetcurr as originalbudgetcurrency,ap.budgetamount as originalbudgetamount,ap.exchangerate as budgetexrate,ap.budgetamount * ap.exchangerate as budgetamount,tl.totalofnotoolings,tl.totaltoolingcost as totaltoolingcostusd,doc.getinvoiceno(ap.id),case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice,inv.totalinvoiceamount as totalinvoiceamountusd,inv.totalinvoiceamount/tl.totaltoolingcost as totalpaid,tl.totaltoolingcost- inv.totalinvoiceamount as totalcostpaymentbalanceusd,(ap.budgetamount * ap.exchangerate )-tl.totaltoolingcost as budgetbalancevstotalcostusd, (ap.budgetamount * ap.exchangerate )-inv.totalinvoiceamount as budgetbalancevsinvoiceamountusd, ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.creator" &
                                     " from  doc.toolingproject tp left join doc.assetpurchase ap on ap.projectid =  tp.id left join vendor v on v.vendorcode = ap.vendorcode left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid left join family f on f.familyid = tp.familyid left join sbusap s on s.sbuid = fgs.sbusapid " &
                                     " left join tl on tl.assetpurchaseid = ap.id" &
                                     " left join inv on inv.assetpurchaseid = ap.id" &
            " where not ap.id isnull {0} order by ap.id desc;", sb.ToString)
        If Not HelperClass1.UserInfo.IsAdmin And Not (HelperClass1.UserInfo.IsFinance) Then
            SqlstrReport = String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname3::text as shortname" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{1}$'   order by vendorname) " &
                                    " , tl as (" &
                                     " select assetpurchaseid,count(0) as totalofnotoolings,sum(cost) as totaltoolingcost from doc.toolinglist " &
                                     " group by assetpurchaseid" &
                                     ")," &
                                     " inv as(" &
                                     " select assetpurchaseid,sum(tp.invoiceamount * tp.exrate) as totalinvoiceamount" &
                                     " from doc.toolinglist tl" &
                                     " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                                     " group by assetpurchaseid" &
                                     " ) " &
                                     " select ap.assetpurchaseid,tp.projectcode,tp.projectname,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text,v.vendorname::text,v.shortname3::text as shortname,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetcurr as originalbudgetcurrency,ap.budgetamount as originalbudgetamount,ap.exchangerate as budgetexrate,ap.budgetamount * ap.exchangerate as budgetamount,tl.totalofnotoolings,tl.totaltoolingcost as totaltoolingcostusd,doc.getinvoiceno(ap.id),case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice,inv.totalinvoiceamount as totalinvoiceamountusd,inv.totalinvoiceamount/tl.totaltoolingcost as totalpaid,tl.totaltoolingcost- inv.totalinvoiceamount as totalcostpaymentbalanceusd,(ap.budgetamount * ap.exchangerate )-tl.totaltoolingcost as budgetbalancevstotalcostusd, (ap.budgetamount * ap.exchangerate )-inv.totalinvoiceamount as budgetbalancevsinvoiceamountusd, ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.creator" &
                                     " from  doc.toolingproject tp left join doc.assetpurchase ap on ap.projectid =  tp.id " &
                                     " inner join va on va.vendorcode = ap.vendorcode " &
                                     " left join vendor v on v.vendorcode = va.vendorcode left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid left join family f on f.familyid = tp.familyid left join sbusap s on s.sbuid = fgs.sbusapid " &
                                     " left join tl on tl.assetpurchaseid = ap.id" &
                                     " left join inv on inv.assetpurchaseid = ap.id" &
                                    " where not ap.id isnull {0} order by ap.id desc;", sb.ToString, HelperClass1.UserInfo.userid)
        End If
        'SqlstrReport = String.Format("with tl as (" &
        '                             " select assetpurchaseid,count(0) as totalofnotoolings,sum(cost) as totaltoolingcost from doc.toolinglist " &
        '                             " group by assetpurchaseid" &
        '                             ")," &
        '                             " inv as(" &
        '                             " select assetpurchaseid,sum(tp.invoiceamount * tp.exrate) as totalinvoiceamount" &
        '                             " from doc.toolinglist tl" &
        '                             " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
        '                             " group by assetpurchaseid" &
        '                             " ) " &
        '                             " select distinct ap.assetpurchaseid,tp.projectcode,tp.projectname,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text,v.vendorname::text,v.shortname::text,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount,tl.totalofnotoolings,tl.totaltoolingcost,doc.getinvoiceno(ap.id),case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice,inv.totalinvoiceamount,inv.totalinvoiceamount/tl.totaltoolingcost as totalpaid,tl.totaltoolingcost- inv.totalinvoiceamount as totalcostpaymentbalance,(ap.budgetamount * ap.exchangerate )-tl.totaltoolingcost as budgetbalancevstotalcost, (ap.budgetamount * ap.exchangerate )-inv.totalinvoiceamount as budgetbalancevsinvoiceamount, ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.creator" &
        '                             " from  doc.toolingproject tp left join doc.assetpurchase ap on ap.projectid =  tp.id left join vendor v on v.vendorcode = ap.vendorcode left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid left join family f on f.familyid = tp.familyid left join sbusap s on s.sbuid = fgs.sbusapid " &
        '                             " left join tl on tl.assetpurchaseid = ap.id" &
        '                             " left join inv on inv.assetpurchaseid = ap.id" &
        '                             " left join doc.assetpurchasetracking at on at.assetpurchaseid = ap.id" &
        '                             " left join agvalue av on av.trackingno = ap.trackingno" &
        '    " where not ap.id isnull {0} order by ap.id desc;", sb.ToString)
        'If Not HelperClass1.UserInfo.IsAdmin And Not (HelperClass1.UserInfo.IsFinance) Then
        '    SqlstrReport = String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text" &
        '                           " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
        '                           " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{1}$'   order by vendorname) " &
        '                            " , tl as (" &
        '                             " select assetpurchaseid,count(0) as totalofnotoolings,sum(cost) as totaltoolingcost from doc.toolinglist " &
        '                             " group by assetpurchaseid" &
        '                             ")," &
        '                             " inv as(" &
        '                             " select assetpurchaseid,sum(tp.invoiceamount * tp.exrate) as totalinvoiceamount" &
        '                             " from doc.toolinglist tl" &
        '                             " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
        '                             " group by assetpurchaseid" &
        '                             " ) " &
        '                             " select distinct ap.assetpurchaseid,tp.projectcode,tp.projectname,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text,v.vendorname::text,v.shortname::text,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount,tl.totalofnotoolings,tl.totaltoolingcost,doc.getinvoiceno(ap.id),case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice,inv.totalinvoiceamount,inv.totalinvoiceamount/tl.totaltoolingcost as totalpaid,tl.totaltoolingcost- inv.totalinvoiceamount as totalcostpaymentbalance,(ap.budgetamount * ap.exchangerate )-tl.totaltoolingcost as budgetbalancevstotalcost, (ap.budgetamount * ap.exchangerate )-inv.totalinvoiceamount as budgetbalancevsinvoiceamount, ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.creator" &
        '                             " from  doc.toolingproject tp left join doc.assetpurchase ap on ap.projectid =  tp.id " &
        '                             " inner join va on va.vendorcode = ap.vendorcode " &
        '                             " left join vendor v on v.vendorcode = va.vendorcode left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid left join family f on f.familyid = tp.familyid left join sbusap s on s.sbuid = fgs.sbusapid " &
        '                             " left join tl on tl.assetpurchaseid = ap.id" &
        '                             " left join inv on inv.assetpurchaseid = ap.id" &
        '                             " left join doc.assetpurchasetracking at on at.assetpurchaseid = ap.id" &
        '                             " left join agvalue av on av.trackingno = ap.trackingno" &
        '                            " where not ap.id isnull {0} order by ap.id desc;", sb.ToString, HelperClass1.UserInfo.userid)
        'End If
        myReportType = ReportType.AssetPurchase
        loadReport()
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        'SqlstrReport = String.Format("with inv as( select id,count(invoiceid) as countinv from (" &
        '               " select   ap.id,ap.assetpurchaseid,tp.invoiceid,sum(tp.invoiceamount * exrate) as invoiceamount" &
        '               " from doc.assetpurchase ap" &
        '               " left join  doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
        '               " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
        '               " group by ap.id,invoiceid" &
        '               " order by ap.id) as invoice" &
        '               " group by id )  ," &
        '               " toolpay as (select tl.toolinglistid," &
        '               " sum(tp.invoiceamount * exrate) as invoiceamount" &
        '               " from doc.toolinglist tl" &
        '               " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
        '               " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid" &
        '               " where not invoiceid isnull" &
        '               " group by tl.toolinglistid)" &
        '               " select tl.toolinglistid as toolingid, ap.assetpurchaseid,tp.projectcode,tp.projectname,tp.ppps,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text," &
        '               " v.vendorname::text,v.shortname::text,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount," &
        '               " tl.sebmodelno,tl.suppliermodelreference,tl.suppliermoldno,tl.toolsdescription,tl.material,tl.cavities,tl.numberoftools,tl.dailycaps,tl.cost,tl.purchasedate,tl.location,tl.comments,inv.countinv," &
        '               " doc.getinvoiceno(ap.id) as invoiceno, tpy.invoiceamount,(1- (tl.cost - tpy.invoiceamount) / (case tl.cost when 0 then 1 else tl.cost end ) )  as paid,(tl.cost - tpy.invoiceamount) as balance,  ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate " &
        '               " from  doc.toolingproject tp " &
        '               " left join doc.assetpurchase ap on ap.projectid =  tp.id " &
        '               " left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
        '               " left join vendor v on v.vendorcode = ap.vendorcode" &
        '               " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid" &
        '               " left join family f on f.familyid = tp.familyid " &
        '               " left join sbusap s on s.sbuid = fgs.sbusapid" &
        '               " left join toolpay tpy on tpy.toolinglistid = tl.toolinglistid" &
        '               " left join inv on inv.id = ap.id" &
        '               " where not ap.id isnull and not tl.toolinglistid isnull {0} order by tl.toolinglistid ;", sb.ToString)
        'SqlstrReport = String.Format("with inv as( select id,count(invoiceid) as countinv from (" &
        '               " select   ap.id,ap.assetpurchaseid,tp.invoiceid,sum(tp.invoiceamount * exrate) as invoiceamount" &
        '               " from doc.assetpurchase ap" &
        '               " left join  doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
        '               " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
        '               " group by ap.id,invoiceid" &
        '               " order by ap.id) as invoice" &
        '               " group by id )  ," &
        '               " toolpay as (select tl.toolinglistid,tl.assetpurchaseid," &
        '               " sum(tp.invoiceamount * exrate) as invoiceamount" &
        '               " from doc.toolinglist tl" &
        '               " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
        '               " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid" &
        '               " where not invoiceid isnull" &
        '               " group by tl.toolinglistid,tl.assetpurchaseid)" &
        '               " select tl.toolinglistid as toolingid, case when commontool then true else null end as commontool,case when substring(tl.toolinglistid,1,1) = 'S' then true else null end as setupcost ,ap.assetpurchaseid,tp.projectcode,tp.projectname,tp.ppps,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text," &
        '               " v.vendorname::text,v.shortname::text,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetcurr as originalbudgetcurrency,ap.budgetamount as originalbudgetamount,ap.exchangerate as budgetexrate,ap.budgetamount * ap.exchangerate as budgetamount," &
        '               " tl.sebmodelno,tl.suppliermodelreference,tl.suppliermoldno,tl.toolsdescription,tl.material,tl.cavities,tl.numberoftools,tl.dailycaps,tl.cost,tl.purchasedate,tl.location,tl.comments,inv.countinv," &
        '               " doc.getinvoiceno(ap.id) as invoiceno, tpy.invoiceamount as invoiceamountusd,(1- (tl.cost - tpy.invoiceamount) / (case tl.cost when 0 then 1 else tl.cost end ) )  as paid,(tl.cost - tpy.invoiceamount) as balanceusd,  ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.creator " &
        '               " from  doc.toolingproject tp " &
        '               " left join doc.assetpurchase ap on ap.projectid =  tp.id " &
        '               " left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
        '               " left join vendor v on v.vendorcode = ap.vendorcode" &
        '               " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid" &
        '               " left join family f on f.familyid = tp.familyid " &
        '               " left join sbusap s on s.sbuid = fgs.sbusapid" &
        '               " left join toolpay tpy on tpy.toolinglistid = tl.toolinglistid and tpy.assetpurchaseid = tl.assetpurchaseid" &
        '               " left join inv on inv.id = ap.id" &
        '               " where not ap.id isnull and not tl.toolinglistid isnull {0} order by tl.toolinglistid ;", sb.ToString)
        SqlstrReport = String.Format("with inv as( select id,count(invoiceid) as countinv from (" &
                      " select   ap.id,ap.assetpurchaseid,tp.invoiceid,sum(tp.invoiceamount * exrate) as invoiceamount" &
                      " from doc.assetpurchase ap" &
                      " left join  doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                      " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                      " group by ap.id,invoiceid" &
                      " order by ap.id) as invoice" &
                      " group by id )  " &
                      " select tl.toolinglistid as toolingid, case when commontool then true else null end as commontool,case when substring(tl.toolinglistid,1,1) = 'S' then true else null end as setupcost ,ap.assetpurchaseid,tp.projectcode,tp.projectname,tp.ppps,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text," &
                      " v.vendorname::text,v.shortname3::text as shortname,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetcurr as originalbudgetcurrency,ap.budgetamount as originalbudgetamount,ap.exchangerate as budgetexrate,ap.budgetamount * ap.exchangerate as budgetamount," &
                      " tl.sebmodelno,tl.suppliermodelreference,tl.suppliermoldno,tl.toolsdescription,tl.material,tl.cavities,tl.numberoftools,tl.dailycaps,tl.cost,tl.purchasedate,tl.location,tl.comments,inv.countinv," &
                      " ti.invoiceno,tpymt.currency as originalinvoicecurr,tpymt.exrate as invoiceexrate, tpymt.invoiceamount,tpymt.invoiceamount * tpymt.exrate as invoiceamountusd,(1- (tl.cost - (tpymt.invoiceamount * tpymt.exrate)) / (case tl.cost when 0 then 1 else tl.cost end ) )  as paid,(tl.cost - (tpymt.invoiceamount * tpymt.exrate )) as balanceusd,  ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.creator ," &
                      " ap.paymententity,ts.toolingsupplierid,ts.toolingsuppliername,ti.proformainvoice as proformapo,ti.currency as pocurrency,ti.amount as poamount" &
                      " from  doc.toolingproject tp " &
                      " left join doc.assetpurchase ap on ap.projectid =  tp.id " &
                      " left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                      " left join vendor v on v.vendorcode = ap.vendorcode" &
                      " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid" &
                      " left join family f on f.familyid = tp.familyid " &
                      " left join sbusap s on s.sbuid = fgs.sbusapid" &
                      " left join doc.toolingpayment tpymt on tpymt.toolinglistid = tl.id" &
                      " left join doc.toolinginvoice ti on ti.id = tpymt.invoiceid" &
                      " left join inv on inv.id = ap.id" &
                      " left join doc.toolingsupplier ts on ts.toolingsupplierid = ap.toolingsupplier" &
                      " where not ap.id isnull and not tl.toolinglistid isnull {0} order by tl.toolinglistid ;", sb.ToString)
        If Not HelperClass1.UserInfo.IsAdmin And Not (HelperClass1.UserInfo.IsFinance) Then
            'SqlstrReport = String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text" &
            '                       " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
            '                       " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{1}$'   order by vendorname)  " &
            '          ", inv as( select id,count(invoiceid) as countinv from (" &
            '          " select   ap.id,ap.assetpurchaseid,tp.invoiceid,sum(tp.invoiceamount * exrate) as invoiceamount" &
            '          " from doc.assetpurchase ap" &
            '          " left join  doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
            '          " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
            '          " group by ap.id,invoiceid" &
            '          " order by ap.id) as invoice" &
            '          " group by id )  ," &
            '          " toolpay as (select tl.toolinglistid,tl.assetpurchaseid," &
            '          " sum(tp.invoiceamount * exrate) as invoiceamount" &
            '          " from doc.toolinglist tl" &
            '          " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
            '          " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid" &
            '          " where not invoiceid isnull" &
            '          " group by tl.toolinglistid,tl.assetpurchaseid)" &
            '          " select tl.toolinglistid as toolingid, case when commontool then true else null end as commontool,case when substring(tl.toolinglistid,1,1) = 'S' then true else null end as setupcost , ap.assetpurchaseid,tp.projectcode,tp.projectname,tp.ppps,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text," &
            '          " v.vendorname::text,v.shortname::text,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetcurr as originalbudgetcurrency,ap.budgetamount as originalbudgetamount,ap.exchangerate as budgetexrate,ap.budgetamount * ap.exchangerate as budgetamount," &
            '          " tl.sebmodelno,tl.suppliermodelreference,tl.suppliermoldno,tl.toolsdescription,tl.material,tl.cavities,tl.numberoftools,tl.dailycaps,tl.cost,tl.purchasedate,tl.location,tl.comments,inv.countinv," &
            '          " doc.getinvoiceno(ap.id) as invoiceno, tpy.invoiceamount as invoiceamountusd,(1- (tl.cost - tpy.invoiceamount) / (case tl.cost when 0 then 1 else tl.cost end ) )  as paid,(tl.cost - tpy.invoiceamount) as balanceusd,  ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.creator " &
            '          " from  doc.toolingproject tp " &
            '          " left join doc.assetpurchase ap on ap.projectid =  tp.id " &
            '          " left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
            '          " inner join va on va.vendorcode = ap.vendorcode" &
            '          " left join vendor v on v.vendorcode = va.vendorcode" &
            '          " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid" &
            '          " left join family f on f.familyid = tp.familyid " &
            '          " left join sbusap s on s.sbuid = fgs.sbusapid" &
            '          " left join toolpay tpy on tpy.toolinglistid = tl.toolinglistid and tpy.assetpurchaseid = tl.assetpurchaseid" &
            '          " left join inv on inv.id = ap.id" &
            '          " where not ap.id isnull and not tl.toolinglistid isnull {0} order by tl.toolinglistid ;", sb.ToString, HelperClass1.UserInfo.userid)
            SqlstrReport = String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname3::text as shortname" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{1}$'   order by vendorname)  " &
                      ", inv as( select id,count(invoiceid) as countinv from (" &
                      " select   ap.id,ap.assetpurchaseid,tp.invoiceid,sum(tp.invoiceamount * exrate) as invoiceamount" &
                      " from doc.assetpurchase ap" &
                      " left join  doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                      " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                      " group by ap.id,invoiceid" &
                      " order by ap.id) as invoice" &
                      " group by id )  " &
                      " select tl.toolinglistid as toolingid, case when commontool then true else null end as commontool,case when substring(tl.toolinglistid,1,1) = 'S' then true else null end as setupcost , ap.assetpurchaseid,tp.projectcode,tp.projectname,tp.ppps,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text," &
                      " v.vendorname::text,v.shortname3::text as shortname,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetcurr as originalbudgetcurrency,ap.budgetamount as originalbudgetamount,ap.exchangerate as budgetexrate,ap.budgetamount * ap.exchangerate as budgetamount," &
                      " tl.sebmodelno,tl.suppliermodelreference,tl.suppliermoldno,tl.toolsdescription,tl.material,tl.cavities,tl.numberoftools,tl.dailycaps,tl.cost,tl.purchasedate,tl.location,tl.comments,inv.countinv," &
                      " ti.invoiceno,tpymt.currency as originalinvoicecurr,tpymt.exrate as invoiceexrate, tpymt.invoiceamount, tpymt.invoiceamount as invoiceamountusd,(1- (tl.cost - (tpymt.invoiceamount * tpymt.exrate)) / (case tl.cost when 0 then 1 else tl.cost end ) )  as paid,(tl.cost - (tpymt.invoiceamount * tpymt.exrate)) as balanceusd,  ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.creator " &
                      " ,ap.paymententity,ts.toolingsupplierid,ts.toolingsuppliername,ti.proformainvoice,ti.currency,ti.amount" &
                      " from  doc.toolingproject tp " &
                      " left join doc.assetpurchase ap on ap.projectid =  tp.id " &
                      " left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
                      " inner join va on va.vendorcode = ap.vendorcode" &
                      " left join vendor v on v.vendorcode = va.vendorcode" &
                      " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid" &
                      " left join family f on f.familyid = tp.familyid " &
                      " left join sbusap s on s.sbuid = fgs.sbusapid" &
                      " left join inv on inv.id = ap.id" &
                      " left join doc.toolingpayment tpymt on tpymt.toolinglistid = tl.id" &
                      " left join doc.toolinginvoice ti on ti.id = tpymt.invoiceid" &
                      " left join doc.toolingsupplier ts on ts.toolingsupplierid = ap.toolingsupplier" &
                      " where not ap.id isnull and not tl.toolinglistid isnull {0} order by tl.toolinglistid ;", sb.ToString, HelperClass1.UserInfo.userid)
        End If
        myReportType = ReportType.ToolingList
        loadReport()
    End Sub
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        DateTimePicker2.Enabled = CheckBox1.Checked
        DateTimePicker3.Enabled = CheckBox1.Checked
    End Sub
    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        DateTimePicker4.Enabled = CheckBox2.Checked
        DateTimePicker5.Enabled = CheckBox2.Checked
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

        'Dim sqlstr = String.Format("select distinct ap.id, ap.assetpurchaseid,tp.projectcode,tp.projectname,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.vendorcode::text,v.vendorname::text,v.shortname::text,ap.applicantdate::text,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.investmentorderno,doc.getinvoiceno(ap.id),case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.applicantdate,ap.creator " &
        '          " from  doc.toolingproject tp" &
        '          " left join doc.assetpurchase ap on ap.projectid =  tp.id" &
        '          " left join vendor v on v.vendorcode = ap.vendorcode" &
        '          " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid" &
        '          " left join family f on f.familyid = tp.familyid" &
        '          " left join sbusap s on s.sbuid = fgs.sbusapid" &
        '          " where not ap.id isnull {0} order by ap.id desc;", sb.ToString.ToLower)
        Dim sqlstr = String.Format("select distinct doc.getassetpurchasestatusname(ap.status) as statusname,ap.id, ap.assetpurchaseid,tp.projectcode,tp.projectname,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.vendorcode::text,v.vendorname::text,v.shortname3::text as shortname,ap.applicantdate::text,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.investmentorderno,doc.getinvoiceno(ap.id),case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.applicantdate,ap.creator " &
                  " from  doc.toolingproject tp" &
                  " left join doc.assetpurchase ap on ap.projectid =  tp.id" &
                  " left join vendor v on v.vendorcode = ap.vendorcode" &
                  " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid" &
                  " left join family f on f.familyid = tp.familyid" &
                  " left join sbusap s on s.sbuid = fgs.sbusapid" &
                  " left join doc.assetpurchasetracking at on at.assetpurchaseid = ap.id" &
                  " left join agvalue av on av.trackingno = at.trackingno" &
                  " left join agreementdt adt on adt.agreement = av.agreement" &
                  " left join doc.assetattachment ast on ast.assetpurchaseid = ap.id" &
                  " left join doc.toolingsupplier ts on ts.toolingsupplierid = ap.toolingsupplier" &
                  " where not ap.id isnull {0} order by ap.id desc;", sb.ToString.ToLower)
        If Not HelperClass1.UserInfo.IsAdmin And Not (HelperClass1.UserInfo.IsFinance) Then

            sqlstr = String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname3::text as shortname" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{1}$'   order by vendorname) " &
                       " select distinct doc.getassetpurchasestatusname(ap.status) as statusname,ap.id, ap.assetpurchaseid,tp.projectcode,tp.projectname,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.vendorcode::text,v.vendorname::text,v.shortname3::text as shortname,ap.applicantdate::text,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.investmentorderno,doc.getinvoiceno(ap.id),case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice,ap.toolingpono,ap.financeassetno,ap.sapcapdate,ap.applicantdate,ap.creator " &
                       " from  doc.toolingproject tp" &
                       " left join doc.assetpurchase ap on ap.projectid =  tp.id" &
                       " inner join va on va.vendorcode = ap.vendorcode" &
                       " left join vendor v on v.vendorcode = va.vendorcode" &
                       " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid" &
                       " left join family f on f.familyid = tp.familyid" &
                       " left join sbusap s on s.sbuid = fgs.sbusapid" &
                       " left join doc.assetpurchasetracking at on at.assetpurchaseid = ap.id" &
                       " left join agvalue av on av.trackingno = at.trackingno" &
                       " left join agreementdt adt on adt.agreement = av.agreement" &
                       " left join doc.assetattachment ast on ast.assetpurchaseid = ap.id" &
                       " left join doc.toolingsupplier ts on ts.toolingsupplierid = ap.toolingsupplier" &
                       " where not ap.id isnull {0} order by ap.id desc;", sb.ToString.ToLower, HelperClass1.UserInfo.userid)
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



    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Not IsNothing(ToolingListBS.Current) Then

            Dim drv = ToolingListBS.Current
            Dim myId = drv.row.item("id")
            Dim myformshow As New FormAssetsPurchase(myid)
            myformshow.ShowDialog()
            Me.Button2.PerformClick()




        End If
    End Sub

    Private Sub DataGridView1_CellContentClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        SqlstrReport = "Select toolingid,changedby,changeddate,latesttoolingid from doc.toolingidhistory;"
        myReportType = ReportType.ToolingHistory
        loadReport()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'ComboBox1.SelectedIndex = -1
        'ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        ComboBox5.SelectedIndex = -1
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False

        DateTimePicker1.Value = Today.Date
        DateTimePicker2.Value = Today.Date
        DateTimePicker3.Value = Today.Date
        DateTimePicker4.Value = Today.Date
        DateTimePicker5.Value = Today.Date
        DateTimePicker6.Value = Today.Date
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        DateTimePicker1.Enabled = CheckBox3.Checked
        DateTimePicker6.Enabled = CheckBox3.Checked
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        ComboBox5.Enabled = CheckBox4.Checked
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        SqlstrReport = String.Format("with tl as ( select assetpurchaseid,count(0) as totalofnotoolings,sum(cost) as totaltoolingcost from doc.toolinglist  group by assetpurchaseid), " &
                                     " inv as( select assetpurchaseid,sum(tp.invoiceamount * tp.exrate) as totalinvoiceamount from doc.toolinglist tl " &
                                     " left join doc.toolingpayment tp on tp.toolinglistid = tl.id group by assetpurchaseid ) " &
                                     " select ap.assetpurchaseid,tp.projectcode,tp.projectname,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text,v.vendorname::text," &
                                     " v.shortname3::text as shortname,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount," &
                                     " tl.totalofnotoolings,tl.totaltoolingcost,doc.getinvoiceno(ap.id),case doc.getinvoiceno(ap.id) when '' then null " &
                                     " else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice," &
                                     " inv.totalinvoiceamount, inv.totalinvoiceamount/tl.totaltoolingcost as totalpaid,tl.totaltoolingcost- inv.totalinvoiceamount as totalcostpaymentbalance," &
                                     " (ap.budgetamount * ap.exchangerate )-tl.totaltoolingcost as budgetbalancevstotalcost, " &
                                     " (ap.budgetamount * ap.exchangerate )-inv.totalinvoiceamount as budgetbalancevsinvoiceamount, " &
                                     " amortperiod_1,amortqty_1,totalamortqty_1,totalamortamount_1,amortperunit_1,amortperiod_2,amortqty_2,totalamortqty_2,totalamortamount_2,amortperunit_2,amortcurr,amortexrate,amortremarks," &
                                     " apt.trackingno,ag.agreement,dt.material,ag.shorttext," &
                                     " startdate,enddate,ag.totalqty,doc.getqty(dt.material,ag.startdate,doc.addyear(ag.startdate,1)) " &
                                     " as c1 ,doc.getqty(dt.material,doc.addyear(ag.startdate,1),doc.addyear(ag.startdate,2)) as c2, " &
                                     " doc.getqty(dt.material, doc.addyear(ag.startdate, 2), doc.addyear(ag.startdate, 3)) as c3 ," &
                                     " ap.creator " &
                                     " from  doc.toolingproject tp  left join doc.assetpurchase ap on ap.projectid =  tp.id  left join doc.assetpurchasetracking apt on apt.assetpurchaseid = ap.id" &
                                     " left join agvalue ag on ag.trackingno = apt.trackingno" &
                                     " left join agreementdt dt on dt.agreement = ag.agreement  left join vendor v on v.vendorcode = ap.vendorcode left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid " &
                                     " left join family f on f.familyid = tp.familyid left join sbusap s on s.sbuid = fgs.sbusapid  left join tl on tl.assetpurchaseid = ap.id " &
                                     " left join inv on inv.assetpurchaseid = ap.id where not ap.id isnull {0} and ap.paymentmethodid = {1:d} ", sb.ToString, PaymentMethodIDEnum.Amortization)
        If Not HelperClass1.UserInfo.IsAdmin And Not (HelperClass1.UserInfo.IsFinance) Then
            SqlstrReport = String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname3::text as shortname" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{1}$'   order by vendorname) " &
                                    " , tl as (" &
                                     " select assetpurchaseid,count(0) as totalofnotoolings,sum(cost) as totaltoolingcost from doc.toolinglist " &
                                     " group by assetpurchaseid" &
                                     ")," &
                                     " inv as(" &
                                     " select assetpurchaseid,sum(tp.invoiceamount * tp.exrate) as totalinvoiceamount" &
                                     " from doc.toolinglist tl" &
                                     " left join doc.toolingpayment tp on tp.toolinglistid = tl.id" &
                                     " group by assetpurchaseid" &
                                     " ) " &
                                     " select ap.assetpurchaseid,tp.projectcode,tp.projectname,f.familyname::character varying,s.sbuname2,ap.applicantname,ap.creator,ap.applicantdate::text,ap.vendorcode::text,v.vendorname::text," &
                                     " v.shortname3::text as shortname,doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname,ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount," &
                                     " tl.totalofnotoolings,tl.totaltoolingcost,doc.getinvoiceno(ap.id),case doc.getinvoiceno(ap.id) when '' then null " &
                                     " else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice," &
                                     " inv.totalinvoiceamount, inv.totalinvoiceamount/tl.totaltoolingcost as totalpaid,tl.totaltoolingcost- inv.totalinvoiceamount as totalcostpaymentbalance," &
                                     " (ap.budgetamount * ap.exchangerate )-tl.totaltoolingcost as budgetbalancevstotalcost, " &
                                     " (ap.budgetamount * ap.exchangerate )-inv.totalinvoiceamount as budgetbalancevsinvoiceamount, " &
                                     " amortperiod_1,amortqty_1,totalamortqty_1,totalamortamount_1,amortperunit_1,amortperiod_2,amortqty_2,totalamortqty_2,totalamortamount_2,amortperunit_2,amortcurr,amortexrate,amortremarks," &
                                     " apt.trackingno,ag.agreement,dt.material,ag.shorttext," &
                                     " startdate,enddate,ag.totalqty,doc.getqty(dt.material,ag.startdate,doc.addyear(ag.startdate,1)) " &
                                     " as c1 ,doc.getqty(dt.material,doc.addyear(ag.startdate,1),doc.addyear(ag.startdate,2)) as c2, " &
                                     " doc.getqty(dt.material, doc.addyear(ag.startdate, 2), doc.addyear(ag.startdate, 3)) as c3 ," &
                                     " ap.creator " &
                                     " from  doc.toolingproject tp  left join doc.assetpurchase ap on ap.projectid =  tp.id  left join doc.assetpurchasetracking apt on apt.assetpurchaseid = ap.id" &
                                     " left join agvalue ag on ag.trackingno = apt.trackingno" &
                                     " left join agreementdt dt on dt.agreement = ag.agreement  " &
                                     " inner join va on va.vendorcode = ap.vendorcode " &
                                     " left join vendor v on v.vendorcode = ap.vendorcode left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid " &
                                     " left join family f on f.familyid = tp.familyid left join sbusap s on s.sbuid = fgs.sbusapid  left join tl on tl.assetpurchaseid = ap.id " &
                                     " left join inv on inv.assetpurchaseid = ap.id" &
                                    " where not ap.id isnull {0}  and ap.paymentmethodid = {2:d} order by ap.id desc;", sb.ToString, HelperClass1.UserInfo.userid, PaymentMethodIDEnum.Investment)
        End If
        myReportType = ReportType.AssetPurchaseAgreement
        loadReport()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click, Button8.Click
        Select Case sender.tag
            Case "shortname"
                Dim myform = New FormHelper(ShortNameBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "shortname"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = ShortNameBS.Current
                    TextBox1.Text = drv.Row.Item("shortname")
                End If
                ProgressReport(6, "marquee")
            Case "aeb"
                Dim myform = New FormHelper(AEBBS)
                myform.DataGridView1.Columns(0).DataPropertyName = "aeb"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = AEBBS.Current
                    TextBox2.Text = drv.Row.Item("aeb")
                End If
                ProgressReport(5, "Continuous")
        End Select
       

    End Sub

    Private Sub InitQueryHelper()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf GetQueryHelper)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub GetQueryHelper()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        sb.Append(String.Format("with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description," &
                                   " v.vendorname::text,shortname3::text as shortname from vendor v order by vendorname)  " &
                                   " select distinct va.shortname from doc.assetpurchase ap " &
                                   " inner join va on va.vendorcode = ap.vendorcode order by va.shortname;"))
        sb.Append(String.Format("with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description," &
                                " v.vendorname::text,shortname3::text as shortname from vendor v order by vendorname)  " &
                                " select distinct ap.aeb " &
                                " from doc.assetpurchase ap " &
                                " inner join va on va.vendorcode = ap.vendorcode  order by ap.aeb;"))

        If Not HelperClass1.UserInfo.IsAdmin And Not (HelperClass1.UserInfo.IsFinance) Then
            sb.Clear()
            sb.Append(String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname3::text as shortname" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{0}$'   order by vendorname) " &
                       " select distinct va.shortname " &
                       " from doc.assetpurchase ap inner join va on va.vendorcode = ap.vendorcode order by va.shortname;", HelperClass1.UserInfo.userid))
            sb.Append(String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname3::text as shortname" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{0}$'   order by vendorname) " &
                       " select distinct ap.aeb " &
                       " from doc.assetpurchase ap inner join va on va.vendorcode = ap.vendorcode order by ap.aeb;", HelperClass1.UserInfo.userid))

        End If


        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "HelperTable"
                ShortNameBS = New BindingSource
                AEBBS = New BindingSource
                ShortNameBS.DataSource = DS.Tables(0)
                AEBBS.DataSource = DS.Tables(1)
            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        'Thread.Sleep(5000)
        sb.Clear()
        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
        ProgressReport(8, "Enable button helper.")
    End Sub

    Private Sub FormReportAssetsPurchase_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        InitQueryHelper()
    End Sub
End Class