Imports SupplierManagement.PublicClass
Imports System.Text
Imports System.Net.Mail
Imports System.Net.Mime
Public Class RoleAssetPurchaseTask
    Inherits Email

    Enum RoleAssetPurchaseTaskEnum
        Creator = 0
        PurchasingValidator1 = 1
        PurchasingValidator2 = 2
        InvestmentApproval = 3
        ControllingDeptAmortization = 4
        ControllingDeptInvestment = 5
        AccountingDept = 6
        ITDept = 7
        Applicant = 8
    End Enum


    Private ds As DataSet
    Public errormessage As String = String.Empty
    Private APPBS As BindingSource
    Private GroupBS As BindingSource
    Private DTBS As BindingSource
    Private myRole As RoleAssetPurchaseTaskEnum
    Public Sub New(ByVal _role As RoleAssetPurchaseTaskEnum)
        myRole = _role
        If IsNothing(DbAdapter1) Then
            DbAdapter1 = New DbAdapter
        End If

    End Sub

    Public Function Execute() As Boolean
        'Initialize Dataset
        Dim myret = False
        Dim mygroup As Object = vbNull
        Dim roleTask As iRoleAssetPurchaseTask = Nothing
        If InitDataset() Then
            If myRole = RoleAssetPurchaseTaskEnum.PurchasingValidator1 Then
                roleTask = New PurchasingValidator1BM(APPBS, GroupBS)
            ElseIf myRole = RoleAssetPurchaseTaskEnum.PurchasingValidator2 Then
                roleTask = New PurchasingValidator2BM(APPBS, GroupBS)
            ElseIf myRole = RoleAssetPurchaseTaskEnum.Creator Then
                roleTask = New RejectedTaskBM(APPBS, GroupBS)
            ElseIf myRole = RoleAssetPurchaseTaskEnum.InvestmentApproval Then
                roleTask = New CreatorInvestmentBM(APPBS, GroupBS)
            ElseIf myRole = RoleAssetPurchaseTaskEnum.ControllingDeptAmortization Then
                roleTask = New ControllingDeptAmortizationBM(APPBS, GroupBS)
            ElseIf myRole = RoleAssetPurchaseTaskEnum.AccountingDept Then
                roleTask = New AccountingDeptBM(APPBS, GroupBS)
            ElseIf myRole = RoleAssetPurchaseTaskEnum.ITDept Then
                roleTask = New ITDeptBM(APPBS, GroupBS)
            ElseIf myRole = RoleAssetPurchaseTaskEnum.Applicant Then
                roleTask = New ApplicantBM(APPBS, GroupBS)
            End If
            myret = True

            For Each n In roleTask.GetQuery
                'Using select is correct. Because Sendto is inside the roleTask.GetQuery
                Me.sendto = Nothing
                Select Case myRole
                    Case RoleAssetPurchaseTaskEnum.PurchasingValidator1
                        If Not IsDBNull(n.data(0).item("approvalname")) Then
                            Me.sendto = n.data(0).item("approvalnameemail")
                            Me.cc = roleTask.GetEmailCC
                            'sendtoname = n.data(0).item("validator1name")
                        End If
                    Case RoleAssetPurchaseTaskEnum.PurchasingValidator2
                        If Not IsDBNull(n.data(0).item("approvalname2")) Then
                            Me.sendto = n.data(0).item("approvalnameemail2")
                            Me.cc = roleTask.GetEmailCC
                        End If
                    Case RoleAssetPurchaseTaskEnum.Creator
                        If Not IsDBNull(n.data(0).item("creatorname")) Then
                            Me.sendto = n.data(0).item("creatoremail")
                            Me.cc = roleTask.GetEmailCC
                        End If
                    Case RoleAssetPurchaseTaskEnum.InvestmentApproval
                        If Not IsDBNull(n.data(0).item("creatorname")) Then
                            Me.sendto = n.data(0).item("creatoremail")
                            Me.cc = roleTask.GetEmailCC
                        End If
                    Case RoleAssetPurchaseTaskEnum.ControllingDeptAmortization
                        Me.sendto = roleTask.GetEmail
                        Me.cc = roleTask.GetEmailCC
                    Case RoleAssetPurchaseTaskEnum.ControllingDeptInvestment
                        Me.sendto = roleTask.GetEmail
                        Me.cc = roleTask.GetEmailCC
                    Case RoleAssetPurchaseTaskEnum.AccountingDept
                        Me.sendto = roleTask.GetEmail
                        Me.cc = roleTask.GetEmailCC
                    Case RoleAssetPurchaseTaskEnum.ITDept
                        Me.sendto = String.Format("{0}", n.data(0).item("creatoremail"))
                        Me.cc = roleTask.GetEmailCC
                    Case RoleAssetPurchaseTaskEnum.Applicant
                        Me.sendto = String.Format("{0}", n.data(0).item("applicantemail"))
                        Me.cc = roleTask.GetEmailCC
                End Select
                If Not IsNothing(Me.sendto) Then
                    'Me.sendto = Me.sendto '"dwijanto@yahoo.com"

                    Me.sender = "no-reply@groupeseb.com"
                    Me.subject = String.Format("Supplier Management - Asset Purchase: Tasks status (Date : {0:dd-MMM-yyyy})", Date.Today) '"***Do not reply to this e-mail.***"
                    Dim mycontent = roleTask.getbodymessage(n.data)
                    Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(String.Format("{0} <br>Or click the Supplier Management icon on your desktop: <br><p> <img src=cid:myLogo> <br></p><p>Supplier Management System Administrator</p></body></html>", mycontent), Nothing, MediaTypeNames.Text.Html)

                    Dim logo As New LinkedResource(Application.StartupPath & "\SupplierManagement.png")
                    logo.ContentId = "myLogo"
                    htmlView.LinkedResources.Add(logo)
                    Me.htmlView = htmlView
                    Me.isBodyHtml = True

                    Me.body = mycontent 'roleTask.getbodymessage(n.data)
                    If Not Me.send(errormessage) Then
                        Logger.log(errormessage)
                    End If
                End If
                'HDBS.EndEdit()
            Next
            Dim ds2 = ds.GetChanges
            'Update Changes
            If Not IsNothing(ds2) Then
                Dim mymessage As String = String.Empty
                Dim ra As Integer
                Dim mye As New ContentBaseEventArgs(ds2, True, errormessage, ra, True)
                If Not DbAdapter1.AssetPurchaseSendEmail(Me, mye) Then
                    Logger.log(mye.message)
                Else
                    ds.AcceptChanges()
                End If
            End If
        End If
        Return myret
    End Function

    Private Function InitDataset() As Boolean
        Dim myret = False
        Dim sb As New StringBuilder
        'sb.Append("select doc.getasetpurchasestatusname(ap.status) as statusname,case ap.paymentmethodid when 1 then 'Amortization' when 2 then 'Invoice Investment' end as paymentmethod,v.shortname,ap.*,u.email as approvalnameemail,u1.email as approvalnameemail2,u2.email as creatoremail from doc.assetpurchase  ap" &
        '          " left join vendor v on v.vendorcode = ap.vendorcode" &
        '          " left join doc.user u on u.username = ap.approvalname left join doc.user u1 on u1.username = ap.approvalname2" &
        '          " left join doc.user u2 on lower(u2.userid) = lower(ap.creator) where status > 0 and sendcomplete isnull;")
        'sb.Append("(with tl as " &
        '          " ( select assetpurchaseid,count(0) as totalofnotoolings,sum(cost) as totaltoolingcost from doc.toolinglist  group by assetpurchaseid), " &
        '          " inv as( select assetpurchaseid,sum(tp.invoiceamount * tp.exrate) as totalinvoiceamount " &
        '          " from doc.toolinglist tl left join doc.toolingpayment tp on tp.toolinglistid = tl.id group by assetpurchaseid )  " &
        '          " Select doc.getassetpurchasestatusname(ap.status) as statusname,ap.status," &
        '          " case ap.paymentmethodid when 1 then 'Amortization' when 2 then 'Invoice Investment' end as paymentmethod," &
        '          " v.shortname::text,ap.vendorcode,tp.projectname,ap.assetdescription," &
        '          " doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname," &
        '          " ap.aeb,ap.financeassetno,ap.budgetcurr,ap.budgetamount * ap.exchangerate as budgetamount,tl.totaltoolingcost," &
        '          " ap.totalamortqty_1,ap.amortperunit_1,ap.amortperiod_1, case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice," &
        '          " inv.totalinvoiceamount,inv.totalinvoiceamount/tl.totaltoolingcost as totalpaid," &
        '          " tl.totalofnotoolings,f.familyname:: character varying, s.sbuname2, ap.applicantdate:: text, ap.applicantname, ap.approvalname,u.email as approvalnameemail,ap.approvalname2,u1.email as approvalnameemail2,u2.username as creatorname,u2.email as creatoremail" &
        '          " from doc.assetpurchase ap left join doc.toolingproject tp on tp.id = ap.projectid left join vendor v on v.vendorcode = ap.vendorcode " &
        '          " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid left join family f on f.familyid = tp.familyid left join sbusap s on s.sbuid = fgs.sbusapid  " &
        '          " left join doc.user u on u.username = ap.approvalname left join doc.user u1 on u1.username = ap.approvalname2" &
        '          " left join doc.user u2 on lower(u2.userid) = lower(ap.creator)" &
        '          " left join tl on tl.assetpurchaseid = ap.id left join inv on inv.assetpurchaseid = ap.id where not ap.status isnull and sendcomplete isnull);")
        sb.Append("(with tl as " &
                  " ( select assetpurchaseid,count(0) as totalofnotoolings,sum(cost) as totaltoolingcost from doc.toolinglist  group by assetpurchaseid), " &
                  " inv as( select assetpurchaseid,sum(tp.invoiceamount * tp.exrate) as totalinvoiceamount " &
                  " from doc.toolinglist tl left join doc.toolingpayment tp on tp.toolinglistid = tl.id group by assetpurchaseid )  " &
                  " Select doc.getassetpurchasestatusname(ap.status) as statusname,ap.*," &
                  " case ap.paymentmethodid when 1 then 'Amortization' when 2 then 'Invoice Investment' end as paymentmethod," &
                  " v.shortname::text,tp.projectname," &
                  " doc.gettypeofinvestmentname(ap.typeofinvestment::int) as typeofinvestmentname," &
                  " ap.budgetamount * ap.exchangerate as budgetamount,tl.totaltoolingcost," &
                  " case doc.getinvoiceno(ap.id) when '' then null else array_length(regexp_split_to_array(doc.getinvoiceno(ap.id),','),1) end as noofinvoice," &
                  " inv.totalinvoiceamount,inv.totalinvoiceamount/tl.totaltoolingcost as totalpaid," &
                  " tl.totalofnotoolings,f.familyname:: character varying, s.sbuname2, u.email as approvalnameemail,u1.email as approvalnameemail2,u2.username as creatorname,u2.email as creatoremail" &
                  "  , u3.email as applicantemail" &
                  " from doc.assetpurchase ap left join doc.toolingproject tp on tp.id = ap.projectid left join vendor v on v.vendorcode = ap.vendorcode " &
                  " left join doc.familygroupsbu fgs on fgs.familyid = tp.familyid left join family f on f.familyid = tp.familyid left join sbusap s on s.sbuid = fgs.sbusapid  " &
                  " left join doc.user u on u.username = ap.approvalname left join doc.user u1 on u1.username = ap.approvalname2" &
                  " left join doc.user u2 on lower(u2.userid) = lower(ap.creator)" &
                  " left join doc.user u3 on lower(u3.username) = lower(ap.applicantname)" &
                  " left join tl on tl.assetpurchaseid = ap.id left join inv on inv.assetpurchaseid = ap.id where not ap.status isnull and sendcomplete isnull);")
        sb.Append("select * from doc.groupuser gu left join doc.user u on u.id = gu.userid left join doc.groupauth ga on ga.groupid  = gu.groupid" &
                  " where ga.groupname in ('Controlling Dept Tooling Investment Approval','Controlling Dept Tooling Amortization Approval', 'Accounting Dept', 'Logistics Dept', 'IT', 'Controlling Dept Tooling Investment CC'," &
                  " 'Controlling Dept Tooling Amortization CC', 'Accounting Dept CC', 'IT CC','Purchasing Dept CC')")
        ds = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString, ds, errormessage) Then
            ds.Tables(0).TableName = "AssetPurchase"

            Dim pk(0) As DataColumn
            pk(0) = ds.Tables(0).Columns("id")
            ds.Tables(0).PrimaryKey = pk

            APPBS = New BindingSource
            GroupBS = New BindingSource

            APPBS.DataSource = ds.Tables(0)
            GroupBS.DataSource = ds.Tables(1)
            myret = True
        End If
        Return myret
    End Function

    
End Class

Public Class PurchasingValidator1BM
    Implements iRoleAssetPurchaseTask

    Private APPBS As BindingSource
    Private GroupBS As BindingSource

    Public Sub New(ByVal APPBS As BindingSource, ByVal GroupBS As BindingSource)
        Me.APPBS = APPBS
        APPBS.Filter = String.Format("status in({0:d},{1:d})", AssetPurchaseStatusEnum.StatusNew, AssetPurchaseStatusEnum.StatusReSubmit)
        Me.GroupBS = GroupBS
    End Sub


    Public Function GetBodyMessage(ByVal Data As Object) As String Implements iRoleAssetPurchaseTask.GetBodyMessage
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {1:dd-MMM-yyyy}<br><br>", Data(0).item("approvalname"), Date.Today))
        'sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th> <th>Quantity</th> <th>Tooling Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th><th>Total Amortization Amount</th><th>Total Amortization Quantity</th> <th>Amortization Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        For Each n As DataRowView In Data

            'sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td><td>{11}</td><td>{12}</td><td>{13}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17}</td><td>{18:0%}</td><td>{19}</td><td>{20}</td><td>{21}</td><td>{22:dd-MMM-yyyy}</td><td>{23}</td><td>{24}</td></tr>",
            '                        n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"),
            '                        n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10:#,##0}</td><td>{11:#,##0}</td><td>{12:#,##0}</td><td>{13:#,##0}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:#,##0}</td><td>{18:0%}</td><td>{19}</td><td>{20}</td><td>{21}</td><td>{22:dd-MMM-yyyy}</td><td>{23}</td><td>{24}</td></tr>",
                                    n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totaltoolingcost"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"),
                                    n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))

            'If n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted Then
            '    n.Item("sendcomplete") = True
            'End If
        Next
        APPBS.EndEdit()
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        Return sb.ToString
    End Function

    Public Function GetQuery() As Object Implements iRoleAssetPurchaseTask.GetQuery
        Return From n In APPBS.List
                       Group n By key = n.item("approvalname") Into Group
                       Select key, data = Group
    End Function

    Public Function GetEmailCC() As String Implements iRoleAssetPurchaseTask.GetEmailCC
        Dim myemail As New StringBuilder

        GroupBS.Filter = String.Format("groupname = 'Purchasing Dept CC'")
        For Each drv In GroupBS.List
            If myemail.Length > 0 Then
                myemail.Append(";")
            End If
            myemail.Append(drv.item("email"))
        Next
        Return myemail.ToString
    End Function

    Public Function GetEmail() As String Implements iRoleAssetPurchaseTask.GetEmail
        Throw New System.Exception("Not implemented.")
    End Function

    Public Function GetEmailName() As String Implements iRoleAssetPurchaseTask.GetEmailName
        Throw New System.Exception("Not implemented.")
    End Function
End Class

Public Class PurchasingValidator2BM
    Implements iRoleAssetPurchaseTask

    Private APPBS As BindingSource
    Private GroupBS As BindingSource

    Public Sub New(ByVal APPBS As BindingSource, ByVal GroupBS As BindingSource)
        Me.APPBS = APPBS
        APPBS.Filter = String.Format("status in({0:d})", AssetPurchaseStatusEnum.StatusFirstValidatedByPurchasing)
        Me.GroupBS = GroupBS
    End Sub


    Public Function GetBodyMessage(ByVal Data As Object) As String Implements iRoleAssetPurchaseTask.GetBodyMessage
        'Dim sb As New StringBuilder
        'sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        'sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {1:dd-MMM-yyyy}<br><br>", Data(0).item("approvalname2"), Date.Today))
        'sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th> <th>Quantity</th> <th>Tooling Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        'For Each n As DataRowView In Data
        '    sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td><td>{11}</td><td>{12}</td><td>{13}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:0%}</td><td>{18}</td><td>{19}</td><td>{20}</td><td>{21:dd-MMM-yyyy}</td><td>{22}</td><td>{23}</td></tr>", n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"), n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
        '    'If n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted Then
        '    '    n.Item("sendcomplete") = True
        '    'End If
        'Next
        'APPBS.EndEdit()
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        'Return sb.ToString
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {1:dd-MMM-yyyy}<br><br>", Data(0).item("approvalname"), Date.Today))
        sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th><th>Total Amortization Amount</th><th>Total Amortization Quantity</th> <th>Amortization Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        For Each n As DataRowView In Data
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10:#,##0}</td><td>{11:#,##0}</td><td>{12:#,##0}</td><td>{13:#,##0}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:#,##0}</td><td>{18:0%}</td><td>{19}</td><td>{20}</td><td>{21}</td><td>{22:dd-MMM-yyyy}</td><td>{23}</td><td>{24}</td></tr>",
                                    n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totaltoolingcost"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"),
                                    n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
            'If n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted Then
            '    n.Item("sendcomplete") = True
            'End If
        Next
        APPBS.EndEdit()
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        Return sb.ToString
    End Function

    Public Function GetQuery() As Object Implements iRoleAssetPurchaseTask.GetQuery
        Return From n In APPBS.List
                       Group n By key = n.item("approvalname2") Into Group
                       Select key, data = Group
    End Function

    Public Function GetEmailCC() As String Implements iRoleAssetPurchaseTask.GetEmailCC
        Dim myemail As New StringBuilder

        GroupBS.Filter = String.Format("groupname = 'Purchasing Dept CC'")
        For Each drv In GroupBS.List
            If myemail.Length > 0 Then
                myemail.Append(";")
            End If
            myemail.Append(drv.item("email"))
        Next
        Return myemail.ToString
    End Function

    Public Function GetEmail() As String Implements iRoleAssetPurchaseTask.GetEmail
        Throw New System.Exception("Not implemented.")
    End Function

    Public Function GetEmailName() As String Implements iRoleAssetPurchaseTask.GetEmailName
        Throw New System.Exception("Not implemented.")
    End Function
End Class

Public Class RejectedTaskBM
    Implements iRoleAssetPurchaseTask

    Private APPBS As BindingSource
    Private GroupBS As BindingSource

    Public Sub New(ByVal APPBS As BindingSource, ByVal GroupBS As BindingSource)
        Me.APPBS = APPBS
        APPBS.Filter = String.Format("status in({0:d},{1:d})", AssetPurchaseStatusEnum.StatusRejectedByControlling, AssetPurchaseStatusEnum.StatusRejectedByPurchasing)
        Me.GroupBS = GroupBS
    End Sub


    Public Function GetBodyMessage(ByVal Data As Object) As String Implements iRoleAssetPurchaseTask.GetBodyMessage
        'Dim sb As New StringBuilder
        'sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        'sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {1:dd-MMM-yyyy}<br><br>", Data(0).item("creatorname"), Date.Today))
        'sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th> <th>Quantity</th> <th>Tooling Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        'For Each n As DataRowView In Data
        '    sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td><td>{11}</td><td>{12}</td><td>{13}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:0%}</td><td>{18}</td><td>{19}</td><td>{20}</td><td>{21:dd-MMM-yyyy}</td><td>{22}</td><td>{23}</td></tr>", n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"), n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
        '    'If n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted Then
        '    '    n.Item("sendcomplete") = True
        '    'End If
        'Next
        'APPBS.EndEdit()
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        'Return sb.ToString
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {1:dd-MMM-yyyy}<br><br>", Data(0).item("approvalname"), Date.Today))
        sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th><th>Total Amortization Amount</th><th>Total Amortization Quantity</th> <th>Amortization Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        For Each n As DataRowView In Data
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10:#,##0}</td><td>{11:#,##0}</td><td>{12:#,##0}</td><td>{13:#,##0}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:#,##0}</td><td>{18:0%}</td><td>{19}</td><td>{20}</td><td>{21}</td><td>{22:dd-MMM-yyyy}</td><td>{23}</td><td>{24}</td></tr>",
                                    n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totaltoolingcost"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"),
                                    n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
            'If n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted Then
            '    n.Item("sendcomplete") = True
            'End If
        Next
        APPBS.EndEdit()
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        Return sb.ToString
    End Function

    Public Function GetQuery() As Object Implements iRoleAssetPurchaseTask.GetQuery
        Return From n In APPBS.List
                       Group n By key = n.item("creator") Into Group
                       Select key, data = Group
    End Function

    Public Function GetEmailCC() As String Implements iRoleAssetPurchaseTask.GetEmailCC
        Dim myemail As New StringBuilder

        GroupBS.Filter = String.Format("groupname = 'Purchasing Dept CC'")
        For Each drv In GroupBS.List
            If myemail.Length > 0 Then
                myemail.Append(";")
            End If
            myemail.Append(drv.item("email"))
        Next
        Return myemail.ToString
    End Function

    Public Function GetEmail() As String Implements iRoleAssetPurchaseTask.GetEmail
        Throw New System.Exception("Not implemented.")
    End Function

    Public Function GetEmailName() As String Implements iRoleAssetPurchaseTask.GetEmailName
        Throw New System.Exception("Not implemented.")
    End Function
End Class

Public Class ControllingDeptAmortizationBM
    Implements iRoleAssetPurchaseTask

    Private APPBS As BindingSource
    Private GroupBS As BindingSource

    Public Sub New(ByVal APPBS As BindingSource, ByVal GroupBS As BindingSource)
        Me.APPBS = APPBS
        APPBS.Filter = String.Format("status in({0:d}) and paymentmethodid = {1:d}", AssetPurchaseStatusEnum.StatusValidatedByPurchasing, PaymentMethodIDEnum.Amortization)
        Me.GroupBS = GroupBS
    End Sub


    Public Function GetBodyMessage(ByVal Data As Object) As String Implements iRoleAssetPurchaseTask.GetBodyMessage
        'Dim sb As New StringBuilder
        'sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        'sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {1:dd-MMM-yyyy}<br><br>", GetEmailName, Date.Today))
        'sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th> <th>Quantity</th> <th>Tooling Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        'For Each n As DataRowView In Data
        '    sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td><td>{11}</td><td>{12}</td><td>{13}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:0%}</td><td>{18}</td><td>{19}</td><td>{20}</td><td>{21:dd-MMM-yyyy}</td><td>{22}</td><td>{23}</td></tr>", n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"), n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
        '    'If n.Item("status") = AssetPurchaseStatusEnum.StatusSendApprovalCreator Then
        '    '    n.Item("status") = AssetPurchaseStatusEnum.StatusWaitingApprovalControlling
        '    'End If
        'Next
        'APPBS.EndEdit()
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        'Return sb.ToString
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        'sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {1:dd-MMM-yyyy}<br><br>", Data(0).item("approvalname"), Date.Today))
        sb.Append(String.Format("<p>Dear User/Approver,</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {0:dd-MMM-yyyy}<br><br>", Date.Today))
        sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th><th>Total Amortization Amount</th><th>Total Amortization Quantity</th> <th>Amortization Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        For Each n As DataRowView In Data
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10:#,##0}</td><td>{11:#,##0}</td><td>{12:#,##0}</td><td>{13:#,##0}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:#,##0}</td><td>{18:0%}</td><td>{19}</td><td>{20}</td><td>{21}</td><td>{22:dd-MMM-yyyy}</td><td>{23}</td><td>{24}</td></tr>",
                                    n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totaltoolingcost"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"),
                                    n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
            'If n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted Then
            '    n.Item("sendcomplete") = True
            'End If
        Next
        APPBS.EndEdit()
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        Return sb.ToString
    End Function

    Public Function GetQuery() As Object Implements iRoleAssetPurchaseTask.GetQuery
        Return From n In APPBS.List
                       Group n By key = n.item("paymentmethodid") Into Group
                       Select key, data = Group
    End Function

    Public Function GetEmailCC() As String Implements iRoleAssetPurchaseTask.GetEmailCC
        Dim myemail As New StringBuilder

        GroupBS.Filter = String.Format("groupname = 'Controlling Dept Tooling Amortization CC'")
        For Each drv In GroupBS.List
            If myemail.Length > 0 Then
                myemail.Append(";")
            End If
            myemail.Append(drv.item("email"))
        Next
        Return myemail.ToString
    End Function

    Public Function GetEmail() As String Implements iRoleAssetPurchaseTask.GetEmail
        Dim myemail As New StringBuilder

        GroupBS.Filter = String.Format("groupname = 'Controlling Dept Tooling Amortization Approval'")
        For Each drv In GroupBS.List
            If myemail.Length > 0 Then
                myemail.Append(";")
            End If
            myemail.Append(drv.item("email"))
        Next
        Return myemail.ToString
    End Function

    Public Function GetEmailName() As String Implements iRoleAssetPurchaseTask.GetEmailName
        Dim myname As New StringBuilder

        GroupBS.Filter = String.Format("groupname = 'Controlling Dept Tooling Amortization Approval'")
        For Each drv In GroupBS.List
            If myname.Length > 0 Then
                myname.Append(",")
            End If
            myname.Append(drv.item("username"))
        Next
        Return myname.ToString
    End Function
End Class

Public Class AccountingDeptBM
    Implements iRoleAssetPurchaseTask

    Private APPBS As BindingSource
    Private GroupBS As BindingSource

    Public Sub New(ByVal APPBS As BindingSource, ByVal GroupBS As BindingSource)
        Me.APPBS = APPBS
        APPBS.Filter = String.Format("status in({0:d})", AssetPurchaseStatusEnum.StatusValidatedByControlling)
        Me.GroupBS = GroupBS
    End Sub


    Public Function GetBodyMessage(ByVal Data As Object) As String Implements iRoleAssetPurchaseTask.GetBodyMessage
        'Dim sb As New StringBuilder
        'sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        'sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {1:dd-MMM-yyyy}<br><br>", GetEmailName, Date.Today))
        'sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th> <th>Quantity</th> <th>Tooling Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        'For Each n As DataRowView In Data
        '    sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td><td>{11}</td><td>{12}</td><td>{13}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:0%}</td><td>{18}</td><td>{19}</td><td>{20}</td><td>{21:dd-MMM-yyyy}</td><td>{22}</td><td>{23}</td></tr>", n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"), n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
        '    If n.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByControlling Then
        '        'n.Item("status") = AssetPurchaseStatusEnum.StatusAccountingStartEdit
        '    End If
        'Next
        'APPBS.EndEdit()
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        'Return sb.ToString
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {1:dd-MMM-yyyy}<br><br>", GetEmailName, Date.Today))
        sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th><th>Total Amortization Amount</th><th>Total Amortization Quantity</th> <th>Amortization Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        For Each n As DataRowView In Data
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10:#,##0}</td><td>{11:#,##0}</td><td>{12:#,##0}</td><td>{13:#,##0}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:#,##0}</td><td>{18:0%}</td><td>{19}</td><td>{20}</td><td>{21}</td><td>{22:dd-MMM-yyyy}</td><td>{23}</td><td>{24}</td></tr>",
                                    n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totaltoolingcost"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"),
                                    n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
            'If n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted Then
            '    n.Item("sendcomplete") = True
            'End If
        Next
        APPBS.EndEdit()
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        Return sb.ToString
    End Function

    Public Function GetQuery() As Object Implements iRoleAssetPurchaseTask.GetQuery
        Return From n In APPBS.List
                       Group n By key = n.item("status") Into Group
                       Select key, data = Group
    End Function

    Public Function GetEmailCC() As String Implements iRoleAssetPurchaseTask.GetEmailCC
        Dim myemail As New StringBuilder

        GroupBS.Filter = String.Format("groupname = 'Accounting Dept CC'")
        For Each drv In GroupBS.List
            If myemail.Length > 0 Then
                myemail.Append(";")
            End If
            myemail.Append(drv.item("email"))
        Next
        Return myemail.ToString
    End Function

    Public Function GetEmail() As String Implements iRoleAssetPurchaseTask.GetEmail
        Dim myemail As New StringBuilder

        GroupBS.Filter = String.Format("groupname = 'Accounting Dept'")
        For Each drv In GroupBS.List
            If myemail.Length > 0 Then
                myemail.Append(";")
            End If
            myemail.Append(drv.item("email"))
        Next
        Return myemail.ToString
    End Function

    Public Function GetEmailName() As String Implements iRoleAssetPurchaseTask.GetEmailName
        Dim myname As New StringBuilder

        GroupBS.Filter = String.Format("groupname = 'Accounting Dept'")
        For Each drv In GroupBS.List
            If myname.Length > 0 Then
                myname.Append(",")
            End If
            myname.Append(drv.item("username"))
        Next
        Return myname.ToString
    End Function
End Class


Public Class ITDeptBM
    Implements iRoleAssetPurchaseTask

    Private APPBS As BindingSource
    Private GroupBS As BindingSource

    Public Sub New(ByVal APPBS As BindingSource, ByVal GroupBS As BindingSource)
        Me.APPBS = APPBS
        APPBS.Filter = String.Format("status in({0:d}) and sendcomplete is null ", AssetPurchaseStatusEnum.StatusCompleted)
        Me.GroupBS = GroupBS
    End Sub


    Public Function GetBodyMessage(ByVal Data As Object) As String Implements iRoleAssetPurchaseTask.GetBodyMessage
        'Dim sb As New StringBuilder        

        'sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        'sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks has been approved and contract# is to be created:<br>Date : {1:dd-MMM-yyyy}<br><br>", Data(0).item("creatorname"), Date.Today))
        'sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th> <th>Quantity</th> <th>Tooling Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        'For Each n As DataRowView In Data
        '    sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td><td>{11}</td><td>{12}</td><td>{13}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:0%}</td><td>{18}</td><td>{19}</td><td>{20}</td><td>{21:dd-MMM-yyyy}</td><td>{22}</td><td>{23}</td></tr>", n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"), n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
        '    If n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted Then
        '        'n.Item("status") = AssetPurchaseStatusEnum.StatusAccountingStartEdit
        '        'n.Item("sendcomplete") = True 'this move to applicant process
        '    End If
        'Next
        'APPBS.EndEdit()
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        'Return sb.ToString
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {1:dd-MMM-yyyy}<br><br>", Data(0).item("creatorname"), Date.Today)) 'approvalname -> creatorname
        sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th><th>Total Amortization Amount</th><th>Total Amortization Quantity</th> <th>Amortization Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        For Each n As DataRowView In Data
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10:#,##0}</td><td>{11:#,##0}</td><td>{12:#,##0}</td><td>{13:#,##0}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:#,##0}</td><td>{18:0%}</td><td>{19}</td><td>{20}</td><td>{21}</td><td>{22:dd-MMM-yyyy}</td><td>{23}</td><td>{24}</td></tr>",
                                    n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totaltoolingcost"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"),
                                    n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
            'If n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted Then
            '    n.Item("sendcomplete") = True
            'End If
        Next
        APPBS.EndEdit()
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        Return sb.ToString
    End Function

    Public Function GetQuery() As Object Implements iRoleAssetPurchaseTask.GetQuery
        Return From n In APPBS.List
                       Group n By key = n.item("creator") Into Group
                       Select key, data = Group
    End Function

    Public Function GetEmailCC() As String Implements iRoleAssetPurchaseTask.GetEmailCC
        Dim myemail As New StringBuilder

        'GroupBS.Filter = String.Format("groupname = 'IT CC'")
        GroupBS.Filter = String.Format("groupname = 'Purchasing Dept CC'")
        For Each drv In GroupBS.List
            If myemail.Length > 0 Then
                myemail.Append(";")
            End If
            myemail.Append(drv.item("email"))
        Next
        Return myemail.ToString
    End Function

    Public Function GetEmail() As String Implements iRoleAssetPurchaseTask.GetEmail
        Dim myemail As New StringBuilder

        'GroupBS.Filter = String.Format("groupname = 'IT'")
        GroupBS.Filter = String.Format("groupname = 'Purchasing Dept CC'")
        For Each drv In GroupBS.List
            If myemail.Length > 0 Then
                myemail.Append(";")
            End If
            myemail.Append(drv.item("email"))
        Next
        Return myemail.ToString
    End Function

    Public Function GetEmailName() As String Implements iRoleAssetPurchaseTask.GetEmailName
        Dim myname As New StringBuilder

        'GroupBS.Filter = String.Format("groupname = 'IT'")
        GroupBS.Filter = String.Format("groupname = 'Purchasing Dept CC'")
        For Each drv In GroupBS.List
            If myname.Length > 0 Then
                myname.Append(",")
            End If
            myname.Append(drv.item("username"))
        Next
        Return myname.ToString
    End Function
End Class

Public Class CreatorInvestmentBM
    Implements iRoleAssetPurchaseTask

    Private APPBS As BindingSource
    Private GroupBS As BindingSource

    Public Sub New(ByVal APPBS As BindingSource, ByVal GroupBS As BindingSource)
        Me.APPBS = APPBS
        APPBS.Filter = String.Format("status in({0:d}) and paymentmethodid = {1:d}", AssetPurchaseStatusEnum.StatusValidatedByPurchasing, PaymentMethodIDEnum.Investment)
        Me.GroupBS = GroupBS
    End Sub


    Public Function GetBodyMessage(ByVal Data As Object) As String Implements iRoleAssetPurchaseTask.GetBodyMessage
        'Dim sb As New StringBuilder
        'sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        'sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks has been approved by Purchasing Validator:<br>Date : {1:dd-MMM-yyyy}<br><br>", Data(0).item("creatorname"), Date.Today))
        'sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th> <th>Quantity</th> <th>Tooling Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        'For Each n As DataRowView In Data
        '    sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td><td>{11}</td><td>{12}</td><td>{13}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:0%}</td><td>{18}</td><td>{19}</td><td>{20}</td><td>{21:dd-MMM-yyyy}</td><td>{22}</td><td>{23}</td></tr>", "Completed", n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"), n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
        '    If n.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByPurchasing Then
        '        n.Item("sendcomplete") = True
        '        n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted
        '    End If
        'Next
        'APPBS.EndEdit()
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        'Return sb.ToString
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {1:dd-MMM-yyyy}<br><br>", Data(0).item("creatorname"), Date.Today)) 'approvalname -> creatorname
        sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th><th>Total Amortization Amount</th><th>Total Amortization Quantity</th> <th>Amortization Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        For Each n As DataRowView In Data
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10:#,##0}</td><td>{11:#,##0}</td><td>{12:#,##0}</td><td>{13:#,##0}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:#,##0}</td><td>{18:0%}</td><td>{19}</td><td>{20}</td><td>{21}</td><td>{22:dd-MMM-yyyy}</td><td>{23}</td><td>{24}</td></tr>",
                                    n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totaltoolingcost"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"),
                                    n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
            If n.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByPurchasing Then
                n.Item("sendcomplete") = True
                'n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted
                n.Item("status") = AssetPurchaseStatusEnum.StatusPendingPayment
            End If
        Next
        APPBS.EndEdit()
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        Return sb.ToString
    End Function

    Public Function GetQuery() As Object Implements iRoleAssetPurchaseTask.GetQuery
        Return From n In APPBS.List
                       Group n By key = n.item("creator") Into Group
                       Select key, data = Group
    End Function

    Public Function GetEmailCC() As String Implements iRoleAssetPurchaseTask.GetEmailCC
        Dim myemail As New StringBuilder

        GroupBS.Filter = String.Format("groupname = 'Purchasing Dept CC'")
        For Each drv In GroupBS.List
            If myemail.Length > 0 Then
                myemail.Append(";")
            End If
            myemail.Append(drv.item("email"))
        Next
        Return myemail.ToString

    End Function

    Public Function GetEmail() As String Implements iRoleAssetPurchaseTask.GetEmail
        'Dim myemail As New StringBuilder

        'GroupBS.Filter = String.Format("groupname = 'IT'")
        'For Each drv In GroupBS.List
        '    If myemail.Length > 0 Then
        '        myemail.Append(";")
        '    End If
        '    myemail.Append(drv.item("email"))
        'Next
        'Return myemail.ToString
        Throw New System.Exception("Not implemented.")
    End Function

    Public Function GetEmailName() As String Implements iRoleAssetPurchaseTask.GetEmailName
        'Dim myname As New StringBuilder

        'GroupBS.Filter = String.Format("groupname = 'IT'")
        'For Each drv In GroupBS.List
        '    If myname.Length > 0 Then
        '        myname.Append(",")
        '    End If
        '    myname.Append(drv.item("username"))
        'Next
        'Return myname.ToString
        Throw New System.Exception("Not implemented.")
    End Function
End Class

Public Class ApplicantBM
    Implements iRoleAssetPurchaseTask

    Private APPBS As BindingSource
    Private GroupBS As BindingSource

    Public Sub New(ByVal APPBS As BindingSource, ByVal GroupBS As BindingSource)
        Me.APPBS = APPBS
        APPBS.Filter = String.Format("status in({0:d}) and sendcomplete is null ", AssetPurchaseStatusEnum.StatusCompleted)
        Me.GroupBS = GroupBS
    End Sub


    Public Function GetBodyMessage(ByVal Data As Object) As String Implements iRoleAssetPurchaseTask.GetBodyMessage
        'Dim sb As New StringBuilder

        'sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        'sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks has been approved and contract# is to be created:<br>Date : {1:dd-MMM-yyyy}<br><br>", Data(0).item("applicantname"), Date.Today))
        'sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th> <th>Quantity</th> <th>Tooling Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        'For Each n As DataRowView In Data
        '    sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td><td>{11}</td><td>{12}</td><td>{13}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:0%}</td><td>{18}</td><td>{19}</td><td>{20}</td><td>{21:dd-MMM-yyyy}</td><td>{22}</td><td>{23}</td></tr>", n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"), n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
        '    If n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted Then
        '        'n.Item("status") = AssetPurchaseStatusEnum.StatusAccountingStartEdit
        '        n.Item("sendcomplete") = True
        '    End If
        'Next
        'APPBS.EndEdit()
        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        'Return sb.ToString
        Dim sb As New StringBuilder
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[Asset Purchase]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt;	font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont"">")
        sb.Append(String.Format("<p>Dear {0},</p><br><p>Please be informed that we have the following Assets Purchase tasks need to follow up:<br>Date : {1:dd-MMM-yyyy}<br><br>", Data(0).item("applicantname"), Date.Today))
        sb.Append("    List of Tasks:</p>  <table border=1 cellspacing=0>    <tr><th>Task Status</th><th>Payment Method</th><th>Supplier Code</th><th>Supplier Short Name</th><th>Project Name</th><th>Asset Description</th><th>Type of Investment</th><th>AEB No</th><th>Fixed Asset#</th><th>Currency</th><th>Budget Amount</th><th>Total Tooling Cost</th><th>Total Amortization Amount</th><th>Total Amortization Quantity</th> <th>Amortization Cost/Unit</th><th>Amortization period</th> <th>No. of Invoice</th><th>Total Payment Amount</th><th>Total Paid</th> <th>No.of Tooling</th> <th>Family</th> <th>SBU</th> <th>Application Date</th> <th>Creator Name</th> <th>Approval From Dept.</th>          </tr>")
        For Each n As DataRowView In Data
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10:#,##0}</td><td>{11:#,##0}</td><td>{12:#,##0}</td><td>{13:#,##0}</td><td>{14}</td><td>{15}</td><td>{16}</td><td>{17:#,##0}</td><td>{18:0%}</td><td>{19}</td><td>{20}</td><td>{21}</td><td>{22:dd-MMM-yyyy}</td><td>{23}</td><td>{24}</td></tr>",
                                    n.Item("statusname"), n.Item("paymentmethod"), n.Item("vendorcode"), n.Item("shortname"), n.Item("projectname"), n.Item("assetdescription"), n.Item("typeofinvestmentname"), n.Item("aeb"), n.Item("financeassetno"), n.Item("budgetcurr"), n.Item("budgetamount"), n.Item("totaltoolingcost"), n.Item("totalamortamount_2"), n.Item("totalamortqty_2"),
                                    n.Item("amortperunit_2"), n.Item("amortperiod_2"), n.Item("noofinvoice"), n.Item("totalinvoiceamount"), n.Item("totalpaid"), n.Item("totalofnotoolings"), n.Item("familyname"), n.Item("sbuname2"), n.Item("applicantdate"), n.Item("creatorname"), n.Item("approvalname")))
            If n.Item("status") = AssetPurchaseStatusEnum.StatusCompleted Then
                'n.Item("status") = AssetPurchaseStatusEnum.StatusAccountingStartEdit
                n.Item("sendcomplete") = True
            End If
        Next
        APPBS.EndEdit()
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        Return sb.ToString
    End Function

    Public Function GetQuery() As Object Implements iRoleAssetPurchaseTask.GetQuery
        Return From n In APPBS.List
                       Group n By key = n.item("applicantname") Into Group
                       Select key, data = Group
    End Function

    Public Function GetEmailCC() As String Implements iRoleAssetPurchaseTask.GetEmailCC
        Dim myemail As New StringBuilder

        'GroupBS.Filter = String.Format("groupname = 'Purchasing Dept CC'")
        GroupBS.Filter = String.Format("groupname = 'IT CC'")
        For Each drv In GroupBS.List
            If myemail.Length > 0 Then
                myemail.Append(";")
            End If
            myemail.Append(drv.item("email"))
        Next
        Return myemail.ToString
    End Function

    Public Function GetEmail() As String Implements iRoleAssetPurchaseTask.GetEmail
        Dim myemail As New StringBuilder

        'GroupBS.Filter = String.Format("groupname = 'Purchasing Dept CC'")
        GroupBS.Filter = String.Format("groupname = 'IT'")
        For Each drv In GroupBS.List
            If myemail.Length > 0 Then
                myemail.Append(";")
            End If
            myemail.Append(drv.item("email"))
        Next
        'GroupBS.Filter = String.Format("groupname = 'IT'")
        'For Each drv In GroupBS.List
        '    If myemail.Length > 0 Then
        '        myemail.Append(";")
        '    End If
        '    myemail.Append(drv.item("email"))
        'Next
        Return myemail.ToString
    End Function

    Public Function GetEmailName() As String Implements iRoleAssetPurchaseTask.GetEmailName
        Dim myname As New StringBuilder

        'GroupBS.Filter = String.Format("groupname = 'Purchasing Dept CC'")
        'For Each drv In GroupBS.List
        '    If myname.Length > 0 Then
        '        myname.Append(",")
        '    End If
        '    myname.Append(drv.item("username"))
        'Next
        GroupBS.Filter = String.Format("groupname = 'IT'")
        For Each drv In GroupBS.List
            If myname.Length > 0 Then
                myname.Append(",")
            End If
            myname.Append(drv.item("username"))
        Next
        Return myname.ToString
    End Function
End Class
