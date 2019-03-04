Imports System.Text
Public Class TaskDocumentExpired
    Implements IDocumentTask

    Private bs As BindingSource

    Public Sub New(ByVal bs As BindingSource)
        Me.bs = bs
    End Sub

    Public Function getBodyMessage(ByVal data As Object) As String Implements IDocumentTask.getBodyMessage
        Dim sb As New StringBuilder
        Dim assignedUser As String = DirectCast(data(0), DataRowView).Item("sendtoname")
        sb.Append("<!DOCTYPE html><html><head><meta name=""description"" content=""[SupplierManagement]"" /><meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii""></head><style>  td,th {padding-left:5px;         padding-right:10px;         text-align:left;  }  th {background-color:red;    color:white}  .defaultfont{    font-size:11.0pt; font-family:""Calibri"",""sans-serif"";    }</style><body class=""defaultfont""><p>Dear " & assignedUser & ",</p>" &
                 " <p>Please be informed that we have the following tasks need to follow up:<br><br>List of Tasks:</p>  <table border=1 cellspacing=0 class=""defaultfont"">" &
                 "<tr><th>Creation Date</th><th>User Name</th> <th>Short Name</th> <th>Supplier(s)</th> <th>Supplier Code</th><th>Supplier Status</th><th>Document Type(s)</th><th>Project/Product Name</th><th>Version</th><th>Document Date</th><th>Expired Date</th><th>Status</th><th>Validator</th><th>CC</th></tr>")
        For Each n As DataRowView In data
            sb.Append(String.Format("<tr><td>{0:dd-MMM-yyyy}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{16}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8:dd-MMM-yyyy}</td><td>{9:dd-MMM-yyyy}</td><td>{10}</td><td>{11}</td><td>{12} {13} {14} {15}</td></tr>", n.Item("creationdate"), n.Item("userid"), n.Item("shortname"), n.Item("suppliername"), n.Item("vendorcode"), n.Item("doctypename"), n.Item("projectname"), n.Item("version"), n.Item("docdate"), n.Item("expireddate"), "Expired", n.Item("validatorname"), n.Item("cc1name"), n.Item("cc2name"), n.Item("cc3name"), n.Item("cc4name"), n.Item("vstatusname")))

        Next

        'sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""http://hon08nt"">Supplier Management</a></p><br><br><p>Supplier Management System Administrator</p></body></html>")
        sb.Append("</table>  <br>  <p>Thank you.<br><br>You can access the system in Citrix by below link:<br>   <a href=""https://sw07e601/RDWeb"">Supplier Management</a></p>")
        Return sb.ToString
    End Function

    Public Function getQuery() As Object Implements IDocumentTask.getQuery
        'Return group and data
        Return From n In bs.List
                         Group n By key = n.item("sendto") Into Group
                         Select key, data = Group
    End Function
End Class
