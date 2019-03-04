Imports SupplierManagement.PublicClass
Imports System.Text
Imports System.Net.Mail
Imports System.Net.Mime
Public Class ActionPlanEmail
    Inherits Email
    Enum RoleDocumentTaskEnum
        CC
        Validator
    End Enum

    Dim DS As DataSet
    'Dim NewBS As BindingSource
    'Dim ExpiredBS As BindingSource
    'Dim EmailMappingBS As BindingSource
    'Private myrole As RoleDocumentTaskEnum
    Public errormessage As String = String.Empty
    Dim BS As BindingSource

    Public Sub New()

        If IsNothing(DbAdapter1) Then
            DbAdapter1 = New DbAdapter
        End If
        initData()
    End Sub

    Public Function Execute() As Boolean
        Dim myret = False
        'Send Email
        For Each n In getQuery()
            'Using select is correct. Because Sendto is inside the roleTask.GetQuery
            Me.sendto = Nothing
           
            If Not IsDBNull(n.data(0).item("email")) Then
                Me.sendto = n.data(0).item("email")

                Me.subject = String.Format("Supplier Management: Action Plan Status ({0:dd-MMM-yyyy}).", Today.Date) '"***Do not reply to this e-mail.***"
                Me.cc = "ttom@groupeseb.com;afok@groupeseb.com"
            End If
           

            If Not IsNothing(Me.sendto) Then
                'Check Mapping User for sendto
                'Me.sendto = "ttom@groupeseb.com;afok@groupeseb.com;vhui@groupeseb.com;dlie@groupeseb.com" 
                'Me.sendto = "dwijanto@yahoo.com"

                Dim mycontent = getBodyMessage(n.data)

                Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(String.Format("{0} <br>Or click the Supplier Management icon on your desktop: <br><p> <img src=cid:myLogo> <br></p><p>Supplier Management System Administrator</p></body></html>", mycontent), Nothing, MediaTypeNames.Text.Html)

                Dim logo As New LinkedResource(Application.StartupPath & "\SupplierManagement.png")
                logo.ContentId = "myLogo"
                htmlView.LinkedResources.Add(logo)

                Me.htmlView = htmlView
                Me.isBodyHtml = True
                Me.sender = "no-reply@groupeseb.com"
                Me.body = mycontent 'roleTask.getbodymessage(n.data)
                If Not Me.send(errormessage) Then
                    Logger.log(errormessage)
                End If
            End If
            myret = True
        Next

        Return myret

    End Function

    Public Function initData() As Boolean
        Dim myret = False
        Dim sb As New StringBuilder
        sb.Append("with v as (select distinct first_value(id) over(partition by actionid order by id desc) as id," &
                  " first_value(validator) over(partition by actionid order by id desc) as validator," &
                  " first_value(status) over(partition by actionid order by id desc) as status from doc.actionplan)" &
                  " select ap.*,u.username,u.email from doc.actionplan ap" &
                  " left join v on v.id = ap.id" &
                  " left join doc.user u on u.userid = v.validator" &
                  " where not v.validator isnull and v.status = 'Closed'")

        Dim sqlstr = sb.ToString
        DS = New DataSet
        If DbAdapter1.TbgetDataSet(sqlstr, DS, errormessage) Then
            DS.Tables(0).TableName = "Action Closed"
            BS = New BindingSource            
            BS.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret
    End Function

    Public Function getBodyMessage(ByVal data As Object) As String
        Dim sb As New StringBuilder
        Dim validatorname As String = DirectCast(data(0), DataRowView).Item("username")
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[SupplierManagement]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt; font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont""><p>Dear " & validatorname & ",</p>" &
                  " <p>Please be informed that we have the following action plan need to follow up:<br><br>List of Tasks:</p>  <table border=1 cellspacing=0 class=""defaultfont"">" &
                  "<tr><th>Short Name</th><th>Action ID</th> <th>Validator</th> <th>Priority (H/M/L)</th> <th>Situation</th><th>Target(s)</th><th>Proposal/Action Plan</th><th>Responsible Person</th><th>Start Date (MM/YY)</th><th>Expected End Date (MMM/YY)</th><th>Actual Result/ Evidence /status</th><th>Actual Finish date (MM/YY)</th><th>Status (Not started / On-going/ Delay/ Closed/ Validated)</th><th>Status Changed Date</th></tr>")
        For Each n As DataRowView In data
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8:MMM-yy}</td><td>{9:MMM-yy}</td><td>{10}</td><td>{11:MMM-yy}</td><td>{12}</td><td>{13:dd-MMM-yyyy}</tr></tr>", n.Item("shortname"), n.Item("actionid"), n.Item("username"), n.Item("priority"), n.Item("situation"), n.Item("target"), n.Item("proposal"), n.Item("responsibleperson"), n.Item("startdate"), n.Item("enddate"), n.Item("result"), n.Item("finishdate"), n.Item("status"), n.Item("statuschangedate")))
        Next

        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        Return sb.ToString
    End Function
    Public Function getQuery() As Object
        'Return group and data
        Return From n In bs.List
                         Group n By key = n.item("validator") Into Group
                         Select key, data = Group

    End Function
End Class
