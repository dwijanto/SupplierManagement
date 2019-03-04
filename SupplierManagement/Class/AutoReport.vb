Imports SupplierManagement.PublicClass
Imports System.Text
Imports System.Net.Mail
Imports System.Net.Mime

Public Class AutoReport
    Inherits Email

    Enum ReportType
        AutoAssetSummary = 0
    End Enum

    Dim MyReportType As ReportType
    Public errorMessage As String = String.Empty
    Dim BS As BindingSource

    Dim minDay As Integer = -1

    Dim DataBS As New BindingSource
    Dim ParamBS As New BindingSource

    Public Sub New(ByRef _reportType As ReportType)
        Me.MyReportType = _reportType
        If IsNothing(DbAdapter1) Then
            DbAdapter1 = New DbAdapter
        End If
    End Sub

    Public Function Execute() As Boolean
        Dim myret = False
        Try
            If MyReportType = ReportType.AutoAssetSummary Then
                'runAutoPriceChangeSummary()
                runAutoAssetSummary()
            End If
            myret = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return myret
    End Function
    Private Sub runAutoAssetSummary()
        Dim myContent As String

        Dim validdate As Date = Today
        'Dim MinDay As Integer = -1
        If validdate.DayOfWeek = DayOfWeek.Monday Then
            minDay = -3
        End If
        Dim DS As New DataSet
        Dim sb As New StringBuilder
        sb.Append(String.Format("(with tl as ( select assetpurchaseid,count(0) as totalofnotoolings,sum(cost) as totaltoolingcost from doc.toolinglist  group by assetpurchaseid), " &
                                " inv as( select assetpurchaseid,sum(tp.invoiceamount * tp.exrate) as totalinvoiceamount from doc.toolinglist tl left join doc.toolingpayment tp on tp.toolinglistid = tl.id group by assetpurchaseid )  " &
                                " Select distinct 'NEW' as taskstatus,v.shortname::text,tp.projectname," &
                                " case paymentmethodid when 1 then 'Amortization' when 2 then 'Invoice Investment' end as paymentmethod," &
                                " doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname," &
                                " ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount,tl.totaltoolingcost," &
                                " case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice," &
                                " inv.totalinvoiceamount,inv.totalinvoiceamount/tl.totaltoolingcost as totalpaid," &
                                " tl.totalofnotoolings,f.familyname:: character varying, s.sbuname2, ap.applicantdate:: text, ap.applicantname, ap.approvalname" &
                                " from doc.assetpurchase ap " &
                                " left join doc.toolingproject tp on tp.id = ap.projectid" &
                                " left join vendor v on v.vendorcode = ap.vendorcode " &
                                " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid " &
                                " left join family f on f.familyid = tp.familyid " &
                                " left join sbusap s on s.sbuid = fgs.sbusapid  " &
                                " left join tl on tl.assetpurchaseid = ap.id " &
                                " left join inv on inv.assetpurchaseid = ap.id " &
                                " where ap.creationdate >= (select pd.ts from paramdt pd left join paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = 'AutoAssetSummary')::date and not shortname isnull)" &
                                " union all " &
                                " (with tl as ( select assetpurchaseid,count(0) as totalofnotoolings,sum(cost) as totaltoolingcost from doc.toolinglist  group by assetpurchaseid), " &
                                " inv as( select assetpurchaseid,sum(tp.invoiceamount * tp.exrate) as totalinvoiceamount from doc.toolinglist tl left join doc.toolingpayment tp on tp.toolinglistid = tl.id group by assetpurchaseid )  " &
                                " Select distinct 'Update' as taskstatus,v.shortname::text,tp.projectname, " &
                                " case paymentmethodid when 1 then 'Amortization' when 2 then 'Invoice Investment' end as paymentmethod," &
                                " doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname," &
                                " ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount,tl.totaltoolingcost," &
                                " case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice," &
                                " inv.totalinvoiceamount,inv.totalinvoiceamount/tl.totaltoolingcost as totalpaid," &
                                " tl.totalofnotoolings,f.familyname::character varying,s.sbuname2,ap.applicantdate::text,ap.applicantname,ap.approvalname" &
                                " from doc.toolingpayment tpay " &
                                " left join doc.toolinginvoice ti on ti.id = tpay.invoiceid" &
                                " left join doc.toolinglist tls on tls.id = tpay.toolinglistid" &
                                " left join doc.assetpurchase ap on ap.id = tls.assetpurchaseid" &
                                " left join doc.toolingproject tp on tp.id = ap.projectid" &
                                " left join vendor v on v.vendorcode = ap.vendorcode " &
                                " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid " &
                                " left join family f on f.familyid = tp.familyid " &
                                " left join sbusap s on s.sbuid = fgs.sbusapid  " &
                                " left join tl on tl.assetpurchaseid = ap.id " &
                                " left join inv on inv.assetpurchaseid = ap.id " &
                                " where ti.creationdate >= (select pd.ts from paramdt pd left join paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = 'AutoAssetSummary')::date and not shortname isnull" &
                                "  and aeb not in (select aeb from doc.assetpurchase where creationdate >= (select pd.ts from paramdt pd left join paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = 'AutoAssetSummary')::date)) ;", Today.Date.AddDays(minDay)))

        sb.Append("select pd.* from paramdt pd" &
                  " left join paramhd ph on ph.paramhdid = pd.paramhdid" &
                  " where ph.paramname = 'AutoAssetSummary' ")

        Dim sqlstr = sb.ToString

        If DbAdapter1.TbgetDataSet(sqlstr, DS, errorMessage) Then
            DataBS.DataSource = DS.Tables(0)
            ParamBS.DataSource = DS.Tables(1)
        End If

        If DS.Tables(0).Rows.Count > 0 Then
            'generate body
            Dim drv As DataRowView = ParamBS.Current
            Dim dbdate As Date = drv.Item("ts")
            If dbdate.Date <> Today.Date Then
                myContent = getbodyMessage(DS) '& "<br></p><p>Supplier Management System Administrator</p></body></html>"

                'Add Image

                Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(String.Format("{0} <br>Or click the Supplier Management icon on your desktop: <br><p> <img src=cid:myLogo> <br></p><p>Supplier Management System Administrator</p></body></html>", myContent), Nothing, MediaTypeNames.Text.Html)

                Dim logo As New LinkedResource(Application.StartupPath & "\SupplierManagement.png")
                logo.ContentId = "myLogo"
                htmlView.LinkedResources.Add(logo)


                'send email
                Dim myrecepient = drv.Row.Item("cvalue").ToString.Split(",")
                'Me.sendto = drv.Row.Item("cvalue") '"dwijanto@yahoo.com"
                Me.sendto = myrecepient(0) '"dwijanto@yahoo.com"
                Me.cc = myrecepient(1)
                Me.isBodyHtml = True
                Me.sender = "no-reply@groupeseb.com"
                Me.subject = String.Format("Supplier Management: Assets Tasks summary. (Date : {0:dd-MMM-yyyy}) ", Today.Date.AddDays(minDay)) '"***Do not reply to this e-mail.***"
                Me.body = myContent
                Me.htmlView = htmlView
                If Not Me.send(errorMessage) Then
                    Logger.log(errorMessage)
                End If

                sqlstr = "update paramdt set ts = now() where" &
                         " paramdtid in (select paramdtid from paramdt pt " &
                         " left join paramhd ph on ph.paramhdid = pt.paramhdid" &
                         " where ph.paramname = 'AutoAssetSummary')"

                If Not DbAdapter1.ExecuteScalar(sqlstr, message:=errorMessage) Then
                    Logger.log(errorMessage)
                End If
            End If



        End If
    End Sub
    'Private Sub runAutoPriceChangeSummary()
    '    'Get Data
    '    Dim myContent As String

    '    Dim validdate As Date = Today
    '    'Dim MinDay As Integer = -1
    '    If validdate.DayOfWeek = DayOfWeek.Monday Then
    '        minDay = -3
    '    End If
    '    Dim ds As New DataSet
    '    Dim sb As New StringBuilder
    '    sb.Append(String.Format("with tl as ( select assetpurchaseid,count(0) as totalofnotoolings,sum(cost) as totaltoolingcost from doc.toolinglist  group by assetpurchaseid), " &
    '                            " inv as( select assetpurchaseid,sum(tp.invoiceamount * tp.exrate) as totalinvoiceamount from doc.toolinglist tl left join doc.toolingpayment tp on tp.toolinglistid = tl.id group by assetpurchaseid )  " &
    '                            " Select distinct 'NEW' as taskstatus,v.shortname::text,tp.projectname," &
    '                            " doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname," &
    '                            " ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount,tl.totaltoolingcost," &
    '                            " case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice," &
    '                            " inv.totalinvoiceamount,inv.totalinvoiceamount/tl.totaltoolingcost as totalpaid," &
    '                            " tl.totalofnotoolings,f.familyname() : character(varying, s.sbuname2, ap.applicantdate) : text, ap.applicantname, ap.approvalname" &
    '                            " from doc.assetpurchase ap " &
    '                            " left join doc.toolingproject tp on tp.id = ap.projectid" &
    '                            " left join vendor v on v.vendorcode = ap.vendorcode " &
    '                            " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid " &
    '                            " left join family f on f.familyid = tp.familyid " &
    '                            " left join sbusap s on s.sbuid = fgs.sbusapid  " &
    '                            " left join tl on tl.assetpurchaseid = ap.id " &
    '                            " left join inv on inv.assetpurchaseid = ap.id " &
    '                            " where ap.creationdate = {0} and not shortname isnull)" &
    '                            " union all " &
    '                            " (with tl as ( select assetpurchaseid,count(0) as totalofnotoolings,sum(cost) as totaltoolingcost from doc.toolinglist  group by assetpurchaseid), " &
    '                            " inv as( select assetpurchaseid,sum(tp.invoiceamount * tp.exrate) as totalinvoiceamount from doc.toolinglist tl left join doc.toolingpayment tp on tp.toolinglistid = tl.id group by assetpurchaseid )  " &
    '                            " Select distinct 'Update' as taskstatus,v.shortname::text,tp.projectname, " &
    '                            " doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname," &
    '                            " ap.aeb,ap.budgetamount * ap.exchangerate as budgetamount,tl.totaltoolingcost," &
    '                            " case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice," &
    '                            " inv.totalinvoiceamount,inv.totalinvoiceamount/tl.totaltoolingcost as totalpaid," &
    '                            " tl.totalofnotoolings,f.familyname::character varying,s.sbuname2,ap.applicantdate::text,ap.applicantname,ap.approvalname" &
    '                            " from doc.toolingpayment tpay " &
    '                            " left join doc.toolinginvoice ti on ti.id = tpay.invoiceid" &
    '                            " left join doc.toolinglist tls on tls.id = tpay.toolinglistid" &
    '                            " left join doc.assetpurchase ap on ap.id = tls.assetpurchaseid" &
    '                            " left join doc.toolingproject tp on tp.id = ap.projectid" &
    '                            " left join vendor v on v.vendorcode = ap.vendorcode " &
    '                            " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid " &
    '                            " left join family f on f.familyid = tp.familyid " &
    '                            " left join sbusap s on s.sbuid = fgs.sbusapid  " &
    '                            " left join tl on tl.assetpurchaseid = ap.id " &
    '                            " left join inv on inv.assetpurchaseid = ap.id " &
    '                            " where ti.creationdate = {0} and not shortname isnull)", Today.Date.AddDays(minDay)))

    '    sb.Append("select pd.* from paramdt pd" &
    '              " left join paramhd ph on ph.paramhdid = pd.paramhdid" &
    '              " where ph.paramname = 'AutoAssetSummary' ")

    '    Dim sqlstr = sb.ToString

    '    If DbAdapter1.TbgetDataSet(sqlstr, ds, errorMessage) Then
    '        ds.Tables(0).TableName = "HD"
    '        BS.DataSource = ds.Tables(0)            
    '    End If

    '    If ds.Tables(0).Rows.Count > 0 Then
    '        'generate body
    '        Dim drv As DataRowView = ParamBS.Current
    '        Dim dbdate As Date = drv.Item("ts")
    '        If dbdate.Date <> Today.Date Then
    '            myContent = getbodyMessage(ds)

    '            'send email
    '            Dim myrecepient = drv.Row.Item("cvalue").ToString.Split(",")
    '            'Me.sendto = drv.Row.Item("cvalue") '"dwijanto@yahoo.com"
    '            Me.sendto = myrecepient(0) '"dwijanto@yahoo.com"
    '            Me.cc = myrecepient(1)
    '            Me.isBodyHtml = True
    '            Me.sender = "no-reply@groupeseb.com"
    '            Me.subject = String.Format("Price CMMF Ex: Tasks summary. (Date : {0:dd-MMM-yyyy}) ", Today.Date.AddDays(minDay)) '"***Do not reply to this e-mail.***"
    '            Me.body = myContent
    '            If Not Me.send(errorMessage) Then
    '                Logger.log(errorMessage)
    '            End If

    '            sqlstr = "update paramdt set ts = now() where" &
    '                     " paramdtid in (select paramdtid from paramdt pt " &
    '                     " left join paramhd ph on ph.paramhdid = pt.paramhdid" &
    '                     " where ph.paramname = 'AutoAssetSummary')"

    '            If Not DbAdapter1.ExecuteScalar(sqlstr, message:=errorMessage) Then
    '                Logger.log(errorMessage)
    '            End If
    '        End If



    '    End If

    'End Sub

    Private Function getbodyMessage(ByVal data As Object) As String


        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[SupplierManagement]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p> <p>Please be informed that we have the following Assets Purchase tasks:<br>Date: {1:dd MMM yyyy}<br><br>", DirectCast(ParamBS.Current, DataRowView).Item("paramname"), Today.Date.AddDays(minDay)))
        'sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr>            <th>Status</th>      <th>Reason</th>            <th>Description</th>      <th>Price Type</th>      <th>Submit Date</th>      <th>Creator</th>      <th>Validator</th>          </tr>")
        sb.Append("List of Tasks:</p>  <table border=1 cellspacing=0 class=""defaultfont"">    <tr>" &
                  " <th>Task Status</th>" &
                  "<th>Supplier Short Name</th>" &
                  " <th>Project Name</th>" &
                  " <th>Payment Method</th>" &
                  " <th>Type of Investment</th>" &
                  " <th>AEB No</th>" &
                  " <th>Budget Amount(USD)</th>" &
                  " <th>Total Tooling Cost (USD)</th>" &
                  " <th>No.of Invoice</th>" &
                  " <th>Total Payment Amount</th>" &
                  " <th>Total Paid</th>" &
                  " <th>No.of Tooling</th>" &
                  " <th>Family</th>" &
                  " <th>SBU</th>" &
                  " <th>Applicant Date</th>" &
                  " <th>Applicant Name</th>" &
                  " <th>Approval From Dept.</th>" &
                  "</tr>")
        For Each n As DataRowView In DataBS.List
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{16}</td><td>{3}</td><td>{4}</td><td>{5:#,##0.00}</td><td>{6:#,##0.00}</td><td>{7}<td>{8:#,##0.00}</td><td>{9:0.00%}</td><td>{10}</td><td>{11}</td><td>{12}</td><td>{13:dd MMM yyyy}</td><td>{14}</td><td>{15}</td></td></tr>", n.Item("taskstatus"), n.Item("shortname"), n.Item("projectname"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("budgetamount"), n.Item("totaltoolingcost"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), CDate(n.Item("applicantdate")), n.Item("applicantname"), n.Item("approvalname"), n.Item("paymentmethod")))

        Next
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>    <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p>")



        Return sb.ToString
    End Function

    Private Function getbodyMessage1(ByVal data As Object) As String
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[SupplierManagement]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p> <p>Please be informed that we have the following Assets Purchase tasks:<br>Date: {1:dd MMM yyyy}<br><br>", DirectCast(ParamBS.Current, DataRowView).Item("paramname"), Today.Date.AddDays(minDay)))
        'sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr>            <th>Status</th>      <th>Reason</th>            <th>Description</th>      <th>Price Type</th>      <th>Submit Date</th>      <th>Creator</th>      <th>Validator</th>          </tr>")
        sb.Append("List of Tasks:</p>  <table border=1 cellspacing=0 class=""defaultfont"">    <tr>" &
                  " <th>Task Status</th>" &
                  "<th>Supplier Short Name</th>" &
                  " <th>Project Name</th>" &
                  " <th>Type of Investment</th>" &
                  " <th>AEB No</th>" &
                  " <th>Budget Amount(USD)</th>" &
                  " <th>Total Tooling Cost (USD)</th>" &
                  " <th>No.of Invoice</th>" &
                  " <th>Total Payment Amount</th>" &
                  " <th>Total Paid</th>" &
                  " <th>No.of Tooling</th>" &
                  " <th>Family</th>" &
                  " <th>SBU</th>" &
                  " <th>Applicant Date</th>" &
                  " <th>Applicant Name</th>" &
                  " <th>Approval From Dept.</th>" &
                  "</tr>")
        For Each n As DataRowView In DataBS.List
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5:#,##0.00}</td><td>{6:#,##0.00}</td><td>{7}<td>{8:#,##0.00}</td><td>{9:0.00%}</td><td>{10}</td><td>{11}</td><td>{12}</td><td>{13:dd MMM yyyy}</td><td>{14}</td><td>{15}</td></td></tr>", n.Item("taskstatus"), n.Item("shortname"), n.Item("projectname"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("budgetamount"), n.Item("totaltoolingcost"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), CDate(n.Item("applicantdate")), n.Item("applicantname"), n.Item("approvalname")))

        Next
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system by below link:<br>    <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br><p>Supplier Management System Administrator</p></body></html>")


        Return sb.ToString
    End Function

End Class
