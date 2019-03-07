Imports System.Net.Mail
Imports System.Text
Imports System.Net.Mime

Public Class VIMEmail
    Inherits Email

    Private dbAdapter1 As DbAdapter = DbAdapter.getInstance
    Private sendtoname As String
    Private drv As DataRowView
    Public errorMessage As String = String.Empty
    Private statusname As String = String.Empty


    Public Function Execute(ByVal sendto As String, ByVal sendtoname As String, ByVal statusname As String, ByVal drv As DataRowView, Optional ByVal cc As String = "") As Boolean
        Dim myret As Boolean = False
        'Prepare Email
        Me.statusname = statusname
        Me.sendtoname = sendtoname
        Me.drv = drv
        Me.sendto = Trim(sendto)
        Me.subject = String.Format("Vendor Information Modification: Tasks status. ({0:dd-MMM-yyyy}).", Today.Date)

        If Not IsNothing(Me.sendto) Then

            Dim mycontent = getBodyMessage()

            Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(String.Format("{0} <br>Or click the Supplier Management icon on your desktop: <br><p> <img src=cid:myLogo> <br></p><p>Supplier Management System Administrator</p></body></html>", mycontent), Nothing, MediaTypeNames.Text.Html)

            Dim logo As New LinkedResource(Application.StartupPath & "\SupplierManagement.png")
            logo.ContentId = "myLogo"
            htmlView.LinkedResources.Add(logo)

            Me.htmlView = htmlView
            Me.isBodyHtml = True
            Me.sender = "no-reply@groupeseb.com"
            Me.body = mycontent 'roleTask.getbodymessage(n.data)
            Me.cc = String.Format("{0}afok@groupeseb.com", cc)
            'If Not Me.send(errorMessage) Then
            '    Logger.log(errorMessage)
            'End If
            myret = Me.send(errorMessage)
        End If
        'myret = True

        Return myret
    End Function

    Private Function getBodyMessage() As String
        Dim sb As New StringBuilder

        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[SupplierManagement]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt; font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont""><p>Dear " & sendtoname & ",</p>" &
                  " <p>Please be informed that we have the following Vendor Information Modification tasks need to follow up:<br><br>List of Tasks:</p>  <table border=1 cellspacing=0 class=""defaultfont"">" &
                  "<tr><th>Status</th><th>Vendor Code</th> <th>Vendor Name</th> <th>Short Name</th> <th>Application Date</th><th>Applicant Name</th><th>Creator</th><th>Supplier Modification ID</th><th>Modification Field</th></tr>")

        sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4:dd-MMM-yyyy}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td></tr>", statusname, drv.Item("vendorcode"), drv.Item("vendorname"), drv.Item("vendorname"), drv.Item("applicantdate"), drv.Item("applicantname"), drv.Item("creator"), drv.Item("suppliermodificationid"), drv.Item("modifiedfield")))
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in RD Web Access by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p><br>")
        Return sb.ToString
    End Function

End Class
