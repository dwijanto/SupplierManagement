Imports SupplierManagement.PublicClass
Imports System.Text
Imports System.Net.Mail
Imports System.Net.Mime

Public Class RoleDocumentTask
    Inherits Email

    Enum RoleDocumentTaskEnum
        CC
        Validator
    End Enum


    Dim DS As DataSet
    Dim NewBS As BindingSource
    Dim ExpiredBS As BindingSource
    Dim EmailMappingBS As BindingSource
    Private myrole As RoleDocumentTaskEnum
    Public errormessage As String = String.Empty

    Public Sub New(ByVal _role As RoleDocumentTaskEnum)
        myrole = _role
        If IsNothing(DbAdapter1) Then
            DbAdapter1 = New DbAdapter
        End If

    End Sub

    Public Function Execute() As Boolean
        Dim myret = False
        Dim roleTask As Object = vbNull
        If initData() Then
            If myrole = RoleDocumentTaskEnum.Validator Then
                roleTask = New TaskDocumentNew(NewBS)
            ElseIf myrole = RoleDocumentTaskEnum.CC Then
                roleTask = New TaskDocumentExpired(ExpiredBS)
            End If
            myret = True
        Else
            Return False
        End If

        'Send Email
        For Each n In roleTask.GetQuery
            'Using select is correct. Because Sendto is inside the roleTask.GetQuery
            Me.sendto = Nothing
            Select Case myrole
                Case RoleDocumentTaskEnum.Validator
                    If Not IsDBNull(n.data(0).item("validatoremail")) Then
                        Me.sendto = n.data(0).item("validatoremail")

                        'sendtoname = n.data(0).item("validator1name")
                        Me.subject = String.Format("Upload document: Validate the task ({0:dd-MMM-yyyy}).", Today.Date) '"***Do not reply to this e-mail.***"
                    End If
                Case RoleDocumentTaskEnum.CC
                    If Not IsDBNull(n.data(0).item("sendtoemail")) Then
                        Me.sendto = n.data(0).item("sendtoemail")
                        Dim pk(0) As Object
                        pk(0) = Me.sendto
                        Dim result As DataRow = DS.Tables(2).Rows.Find(pk)
                        If Not IsNothing(result) Then
                            Me.sendto = result.Item("newemail")
                        End If
                        Me.cc = "ttom@groupeseb.com;afok@groupeseb.com"
                        'sendtoname = n.data(0).item("validator2name")
                        Me.subject = String.Format("Upload document: Task Summary ({0:dd-MMM-yyyy}).", Today.Date) '"***Do not reply to this e-mail.***"
                    End If
            End Select
            If Not IsNothing(Me.sendto) Then
                'Check Mapping User for sendto
                'Me.sendto = "ttom@groupeseb.com;afok@groupeseb.com;vhui@groupeseb.com;dlie@groupeseb.com" 'Me.sendto '"dwijanto@yahoo.com"

                Dim mycontent = roleTask.getbodymessage(n.data)

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
        Next
        Return myret

    End Function

    Public Function initData() As Boolean
        Dim myret = False
        Dim sb As New StringBuilder

        sb.Append("with vd as ( select distinct headerid,status from doc.vendordoc where not status isnull ) Select h.id, 'New'::text as statusname,h.creationdate,o.userid::character varying as username,doc.sp_getshortname(h.id) as shortname,doc.sp_getsuppliers(h.id) as suppliername ,doc.sp_getvendorcode(h.id)::text as vendorcode,  null::text as vstatusname,doc.sp_getdoctypename(h.id)::character varying as doctypename,  null::character varying as projectname,null::character varying as version, null::date as docdate," &
                  " null::date as expireddate,null::integer as statusexp,o2.officersebname::character varying as validatorname, o3.officersebname::character varying as cc1name ,o4.officersebname::character varying as cc2name,  o5.officersebname::character varying as cc3name,o6.officersebname::character varying as cc4name , h.userid,h.validator,h.cc1,h.cc2,h.cc3,h.cc4,h.otheremail,h.latestupdate,Null::integer as doctypeid,  null::bigint as vdid ,o2.email as validatoremail,o.username as creatorname from doc.header h   left join doc.user o on lower(o.userid) = h.userid " &
                  " left join officerseb o2 on lower(o2.userid) = h.validator  left join officerseb o3 on lower(o3.userid) = h.cc1 left join officerseb o4 on lower(o4.userid) = h.cc2 left join officerseb o5 on lower(o5.userid) = h.cc3 left join officerseb o6 on lower(o6.userid) = h.cc4 left join vd on vd.headerid = h.id  where  vd.status = 1 ;")
        sb.Append("with dt as (select * from (select vd.id,hd.userid as sendto,o.email as sendtoemail,vd.status from doc.document d left join doc.vendordoc vd on vd.documentid = d.id left join doc.header hd on hd.id = vd.headerid left join doc.user o on lower(o.userid) = hd.userid inner join doc.docexpired de on d.id = de.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid         where(vd.status < 3 And (expireddate - current_date >= 0 And expireddate - current_date <= dr.reminder))" &
                  " union all select vd.id,hd.cc1,o.email as sendtoemail,vd.status from doc.document d left join doc.vendordoc vd on vd.documentid = d.id left join doc.header hd on hd.id = vd.headerid left join doc.user o on lower(o.userid) = hd.userid inner join doc.docexpired de on d.id = de.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid where vd.status < 3 and (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) and not hd.cc1 isnull " &
                  " union all select vd.id,hd.cc2,o.email as sendtoemail,vd.status from doc.document d left join doc.vendordoc vd on vd.documentid = d.id left join doc.header hd on hd.id = vd.headerid left join doc.user o on lower(o.userid) = hd.userid inner join doc.docexpired de on d.id = de.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid where vd.status < 3 and (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) and not hd.cc2 isnull " &
                  " Union all select vd.id,hd.cc3,o.email as sendtoemail,vd.status from doc.document d left join doc.vendordoc vd on vd.documentid = d.id left join doc.header hd on hd.id = vd.headerid left join doc.user o on lower(o.userid) = hd.userid inner join doc.docexpired de on d.id = de.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid where vd.status < 3 and (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) and not hd.cc3 isnull " &
                  " union all select vd.id,hd.cc4,o.email as sendtoemail,vd.status from doc.document d left join doc.vendordoc vd on vd.documentid = d.id left join doc.header hd on hd.id = vd.headerid left join doc.user o on lower(o.userid) = hd.userid inner join doc.docexpired de on d.id = de.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid where vd.status < 3 and (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) and not hd.cc4 isnull) foo), " &
                  " myfull as (select * from (with vst as(select paramname as statusname ,paramdtid as id from doc.paramdt where paramhdid = 2) Select hd.id,'Expired'::text as statusname ,hd.creationdate,u.userid as username,v.shortname::text,v.vendorname::text as suppliername,v.vendorcode::text,vst.statusname::text as vstatusname,dt.doctypename,p.projectname,vr.version,d.docdate,de.expireddate,de.status as statusexp,o2.officersebname::character varying as validatorname,o3.officersebname::character varying as cc1name ,o4.officersebname::character varying as cc2name," &
                  " o5.officersebname::character varying as cc3name,o6.officersebname::character varying as cc4name ,hd.userid,hd.validator,hd.cc1,hd.cc2,hd.cc3,hd.cc4,hd.otheremail,hd.latestupdate,dt.id as doctypeid ,vd.id as vdid from doc.document d left join doc.vendordoc vd on vd.documentid = d.id 	left join doc.header hd on hd.id = vd.headerid 	left join doc.user o on lower(o.userid) = hd.userid left join officerseb o2 on lower(o2.userid) = hd.validator left join officerseb o3 on lower(o3.userid) = hd.cc1 left join officerseb o4 on lower(o4.userid) = hd.cc2 " &
                  " left join officerseb o5 on lower(o5.userid) = hd.cc3 left join officerseb o6 on lower(o6.userid) = hd.cc4 inner join doc.docexpired de on d.id = de.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid left join doc.user u on u.userid = hd.userid  left join vendor v on v.vendorcode = vd.vendorcode left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode left join vst on vst.id = vs.status left join doc.doctype dt on dt.id = d.doctypeid left join doc.project p on p.documentid = d.id left join doc.version vr on vr.documentid = d.id" &
                  " where vd.status < 3 and (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) and (de.status isnull or de.status <> 3)  and not hd.id isnull) as foo  ) select dt.*,m.*,u.username as sendtoname from dt left join myfull m on m.vdid = dt.id left join doc.user u on u.userid = dt.sendto	;")
        sb.Append("select * from doc.emailmapping;")
        Dim sqlstr = sb.ToString
        DS = New DataSet
        If DbAdapter1.TbgetDataSet(sqlstr, DS, errormessage) Then
            DS.Tables(0).TableName = "New Document"
            DS.Tables(1).TableName = "Expired Document"
            DS.Tables(2).TableName = "Email Mapping"
            Dim pk2(0) As DataColumn
            pk2(0) = DS.Tables(2).Columns("oldemail")
            DS.Tables(2).PrimaryKey = pk2

            'Dim rel As DataRelation
            'Dim hcol As DataColumn
            'Dim dcol As DataColumn
            ''create relation ds.table(0) and ds.table(4)
            'hcol = ds.Tables(0).Columns("pricechangehdid")
            'dcol = ds.Tables(1).Columns("pricechangehdid")
            'rel = New DataRelation("hdrel", hcol, dcol)
            'ds.Relations.Add(rel)

            'HDBS = New BindingSource
            'DTBS = New BindingSource
            'HDBS.DataSource = ds.Tables(0)
            'DTBS.DataSource = HDBS
            'DTBS.DataMember = "hdrel"

            NewBS = New BindingSource
            ExpiredBS = New BindingSource
            EmailMappingBS = New BindingSource
            NewBS.DataSource = DS.Tables(0)
            ExpiredBS.DataSource = DS.Tables(1)
            EmailMappingBS.DataSource = DS.Tables(2)

            myret = True
        End If
        Return myret
    End Function

End Class
