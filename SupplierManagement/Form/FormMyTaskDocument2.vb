Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class FormMyTaskDocument2
    Dim limit As String = " limit 1"
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)

    Dim myFilter As New System.Threading.Thread(AddressOf runfilterWork)


    Dim bsheader As New BindingSource
    Dim bshistory As New BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim creator As String
    Dim myuser As String = String.Empty
    Dim headerid As Long = 0
    Dim myListArray As String() = {"lower(username)", "lower(suppliername)", "lower(shortname)", "lower(doctypename)"}
    Dim limitHistory As String
    WithEvents EPCheckBox1 As New CheckBox
    Dim dtpicker1 As New DateTimePicker
    Dim lbl1 As New Label
    Dim dtpicker2 As New DateTimePicker
    WithEvents button1 As New Button
    Dim insertAnd As String = ""
    Dim Criteria As String = ""
    Dim CriteriaNew As String = ""
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Text = Me.Text & "-" & HelperClass1.UserId
        limitHistory = "limit 100"
        EPCheckBox1.Text = "Expired Date Filter"
        EPCheckBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        ToolStrip1.Items.Insert(8, New ToolStripControlHost(EPCheckBox1))
        lbl1.Text = " To "
        dtpicker1.CustomFormat = "dd-MMM-yyyy"
        dtpicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        dtpicker1.Size = New Point(100, 20)
        dtpicker2.CustomFormat = "dd-MMM-yyyy"
        dtpicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        dtpicker2.Size = New Point(100, 20)
        button1.Text = "Do Filter"
        button1.Size = New Point(80, 20)
        ToolStrip1.Items.Insert(9, New ToolStripControlHost(dtpicker1))
        ToolStrip1.Items.Insert(10, New ToolStripControlHost(lbl1))
        ToolStrip1.Items.Insert(11, New ToolStripControlHost(dtpicker2))
        ToolStrip1.Items.Insert(12, New ToolStripControlHost(button1))

        ComboBox1.SelectedIndex = 1
        ToolStripComboBox1.SelectedIndex = 0
    End Sub




    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If Not myThread.IsAlive Then

            'Dim myform = New FormDocumentVendor()
            'Dim myform = New FormDocumentVendor2()
            'Dim myform = New FormDocumentVendor3()
            'Dim myform = New FormDocumentVendor4()
            'Dim myform = New FormDocumentVendor5()
            'Dim myform = New FormDocumentVendor6()
            Dim myform = New FormDocumentVendor7()
            myform.loaddata(0)
            If myform.ShowDialog = DialogResult.OK Then
                loaddata()
            End If

        Else
            MessageBox.Show("Still loading... Please wait.")
        End If


    End Sub



    Private Sub loaddata()
        EPCheckBox1_CheckedChanged(EPCheckBox1, New System.EventArgs)
        myuser = HelperClass1.UserId.ToLower

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
        'Thread.Sleep(10000)
        ProgressReport(1, "Loading Data.")
        'ProgressReport(4, "InitData")
        '2 Dataset 
        '1 contains All tx except Completed
        'the other only contains Completed

        DS = New DataSet
        Dim mymessage As String = String.Empty
        sb.Clear()
        'Admin checking first
        'sb.Append("select * from pricechangehd ph where (ph.creator = '" & HelperClass1.UserId & "')")

        myuser = HelperClass1.UserId.ToLower
        'myuser = "as\dlie"
        'myuser = "as\elai"
        'myuser = "as\rlo"
        'myuser = "AS\afok".ToLower
        'myuser = "as\weho"
        'myuser = "AS\shxu".ToLower
        'myuser = "as\jdai"
        'myuser = "AS\SCHAN".ToLower
        'myuser = "as\vhui"
        'myuser = "as\cchiu"
        'HelperClass1.UserInfo.IsAdmin = False
        'HelperClass1.UserInfo.AllowUpdateDocument = False

        'sb.Append("select distinct h.*,doc.sp_getsuppliers(h.id) as suppliername,doc.sp_getshortname(h.id) as shortname,doc.sp_getdoctypename(h.id) as doctypename ,o.username::text as username,o2.officersebname::text as validatorname,o3.officersebname::text as cc1name ,o4.officersebname::text as cc2name,o5.officersebname::text as cc3name,o6.officersebname::text as cc4name from doc.header h " &
        '      " left join doc.user o on lower(o.userid) = h.userid" &
        '      " left join officerseb o2 on lower(o2.userid) = h.validator" &
        '      " left join officerseb o3 on lower(o3.userid) = h.cc1" &
        '      " left join officerseb o4 on lower(o4.userid) = h.cc2" &
        '      " left join officerseb o5 on lower(o5.userid) = h.cc3" &
        '      " left join officerseb o6 on lower(o6.userid) = h.cc4" &                 
        '      " where h.userid = '" & myuser & "' or h.validator = '" & myuser & "'  or " & HelperClass1.UserInfo.IsAdmin & " or " & HelperClass1.UserInfo.AllowUpdateDocument & " order by latestupdate desc;")
        'sb.Append("with vd as (" &
        '          " select distinct headerid,status from doc.vendordoc where not status isnull )" &
        '          " select distinct 'New' as statusname,h.*,doc.sp_getsuppliers(h.id) as suppliername,doc.sp_getshortname(h.id) as shortname,doc.sp_getdoctypename(h.id) as doctypename ,o.username::text as username,o2.officersebname::text as validatorname,o3.officersebname::text as cc1name ,o4.officersebname::text as cc2name,o5.officersebname::text as cc3name,o6.officersebname::text as cc4name from doc.header h " &
        '          " left join doc.user o on lower(o.userid) = h.userid" &
        '          " left join officerseb o2 on lower(o2.userid) = h.validator" &
        '          " left join officerseb o3 on lower(o3.userid) = h.cc1" &
        '          " left join officerseb o4 on lower(o4.userid) = h.cc2" &
        '          " left join officerseb o5 on lower(o5.userid) = h.cc3" &
        '          " left join officerseb o6 on lower(o6.userid) = h.cc4" &
        '          " left join vd on vd.headerid = h.id" &
        '          " where (h.userid = '" & myuser & "' or h.validator = '" & myuser & "' or h.cc1 = '" & myuser & "' or h.cc2 = '" & myuser & "' or h.cc3 = '" & myuser & "' or h.cc4 = '" & myuser & "'  or " & HelperClass1.UserInfo.IsAdmin & " or " & HelperClass1.UserInfo.AllowUpdateDocument & ") and vd.status = 1 order by latestupdate desc;")

        sb.Append(String.Format("select * from (select * from (with vd as (" &
                  " select distinct headerid,status from doc.vendordoc where not status isnull )" &
                  " Select h.id, 'New'::text as statusname,h.creationdate,o.userid::character varying as username,doc.sp_getshortname(h.id) as shortname,doc.sp_getsuppliers(h.id) as suppliername ,doc.sp_getvendorcode(h.id)::text as vendorcode, null::text as vstatusname,doc.sp_getdoctypename(h.id)::character varying as doctypename,null::character varying as projectname,null::character varying as version, null::date as docdate,null::date as expireddate,null::integer as statusexp,o2.officersebname::character varying as validatorname,o3.officersebname::character varying as cc1name ,o4.officersebname::character varying as cc2name,o5.officersebname::character varying as cc3name,o6.officersebname::character varying as cc4name ,h.userid,h.validator,h.cc1,h.cc2,h.cc3,h.cc4,h.otheremail,h.latestupdate,0::integer as doctypeid,0::bigint as vdid,0::bigint as documentid" &
                  " from doc.header h " &
                  " left join doc.user o on lower(o.userid) = h.userid" &
                  " left join officerseb o2 on lower(o2.userid) = h.validator" &
                  " left join officerseb o3 on lower(o3.userid) = h.cc1" &
                  " left join officerseb o4 on lower(o4.userid) = h.cc2" &
                  " left join officerseb o5 on lower(o5.userid) = h.cc3" &
                  " left join officerseb o6 on lower(o6.userid) = h.cc4" &
                  " left join vd on vd.headerid = h.id" &
                  " where (h.userid = '" & myuser & "' or h.validator = '" & myuser & "' or h.cc1 = '" & myuser & "' or h.cc2 = '" & myuser & "' or h.cc3 = '" & myuser & "' or h.cc4 = '" & myuser & "'  or " & (HelperClass1.UserInfo.IsAdmin) & " or " & (HelperClass1.UserInfo.AllowUpdateDocument) & ") and vd.status = 1 ) as foo {0}) as foo1 union all ", CriteriaNew))
        'sb.Append(String.Format(" select * from (with vst as(select paramname as statusname ,ivalue as id from doc.paramdt where paramhdid = 2)" &
        '          " select hd.id,'Expired'::text as statusname ,hd.creationdate,u.userid as username,v.shortname::text,v.vendorname::text as suppliername,v.vendorcode::text,vst.statusname::text as vstatusname,dt.doctypename,p.projectname,vr.version,d.docdate,de.expireddate,de.status as statusexp,o2.officersebname::character varying as validatorname,o3.officersebname::character varying as cc1name ,o4.officersebname::character varying as cc2name,o5.officersebname::character varying as cc3name,o6.officersebname::character varying as cc4name ,hd.userid,hd.validator,hd.cc1,hd.cc2,hd.cc3,hd.cc4,hd.otheremail,hd.latestupdate,dt.id as doctypeid ,vd.id as vdid" &
        '          " from doc.document d" &
        '          " left join doc.vendordoc vd on vd.documentid = d.id left join doc.header hd on hd.id = vd.headerid left join doc.user o on lower(o.userid) = hd.userid  left join officerseb o2 on lower(o2.userid) = hd.validator left join officerseb o3 on lower(o3.userid) = hd.cc1 left join officerseb o4 on lower(o4.userid) = hd.cc2 left join officerseb o5 on lower(o5.userid) = hd.cc3 left join officerseb o6 on lower(o6.userid) = hd.cc4 inner join doc.docexpired de on d.id = de.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid left join doc.user u on u.userid = hd.userid " &
        '          " left join vendor v on v.vendorcode = vd.vendorcode left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode left join vst on vst.id = vs.status left join doc.doctype dt on dt.id = d.doctypeid left join doc.project p on p.documentid = d.id left join doc.version vr on vr.documentid = d.id where vd.status < 3 and (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) and (de.status isnull or de.status <> 3) and " &
        '          " (hd.userid = '" & myuser & "' or hd.validator = '" & myuser & "' or hd.cc1 = '" & myuser & "' or hd.cc2 = '" & myuser & "' or hd.cc3 = '" & myuser & "' or hd.cc4 = '" & myuser & "'  or " & (HelperClass1.UserInfo.IsAdmin) & " or " & (HelperClass1.UserInfo.AllowUpdateDocument) & ")  and not hd.id isnull) as foo {0} ;", CriteriaNe w))


        'sb.Append(String.Format(" select * from (with vst as(select paramname as statusname ,ivalue as id from doc.paramdt where paramhdid = 2)" &
        '         " select hd.id,'Expired'::text as statusname ,hd.creationdate,u.userid as username,v.shortname::text,v.vendorname::text as suppliername,v.vendorcode::text,vst.statusname::text as vstatusname,dt.doctypename,p.projectname,vr.version,d.docdate,de.expireddate,de.status as statusexp,o2.officersebname::character varying as validatorname,o3.officersebname::character varying as cc1name ,o4.officersebname::character varying as cc2name,o5.officersebname::character varying as cc3name,o6.officersebname::character varying as cc4name ,hd.userid,hd.validator,hd.cc1,hd.cc2,hd.cc3,hd.cc4,hd.otheremail,hd.latestupdate,dt.id as doctypeid ,vd.id as vdid,vd.documentid as documentid" &
        '         " from doc.document d" &
        '         " left join doc.vendordoc vd on vd.documentid = d.id left join doc.header hd on hd.id = vd.headerid left join doc.user o on lower(o.userid) = hd.userid  left join officerseb o2 on lower(o2.userid) = hd.validator left join officerseb o3 on lower(o3.userid) = hd.cc1 left join officerseb o4 on lower(o4.userid) = hd.cc2 left join officerseb o5 on lower(o5.userid) = hd.cc3 left join officerseb o6 on lower(o6.userid) = hd.cc4 inner join doc.docexpired de on d.id = de.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid left join doc.user u on u.userid = hd.userid " &
        '         " left join vendor v on v.vendorcode = vd.vendorcode left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode left join vst on vst.id = vs.status left join doc.doctype dt on dt.id = d.doctypeid left join doc.project p on p.documentid = d.id left join doc.version vr on vr.documentid = d.id where vd.status < 3 and (expireddate - current_date  >= 0 and  expireddate - current_date  <= dr.reminder) and (de.status isnull or de.status <> 3) and " &
        '         " (hd.userid in (select '" & myuser & "' union all" &
        '        " select o.userid from doc.emailmapping e" &
        '        " left join doc.user o on o.email =  e.oldemail" &
        '        " left join doc.user n on n.email = e.newemail" &
        '        " where lower(n.userid) = '" & myuser & "') or hd.validator = '" & myuser & "' or hd.cc1 = '" & myuser & "' or hd.cc2 = '" & myuser & "' or hd.cc3 = '" & myuser & "' or hd.cc4 = '" & myuser & "'  or " & (HelperClass1.UserInfo.IsAdmin) & " or " & (HelperClass1.UserInfo.AllowUpdateDocument) & ")  and not hd.id isnull) as foo {0} union all ", CriteriaNew))
        'Remove reminder until current date, allow to show all (expireddate - current_date  >= 0 and 
        sb.Append(String.Format(" select * from (with vst as(select paramname as statusname ,ivalue as id from doc.paramdt where paramhdid = 2)" &
                 " select hd.id,'Expired'::text as statusname ,hd.creationdate,u.userid as username,v.shortname::text,v.vendorname::text as suppliername,v.vendorcode::text,vst.statusname::text as vstatusname,dt.doctypename,p.projectname,vr.version,d.docdate,de.expireddate,de.status as statusexp,o2.officersebname::character varying as validatorname,o3.officersebname::character varying as cc1name ,o4.officersebname::character varying as cc2name,o5.officersebname::character varying as cc3name,o6.officersebname::character varying as cc4name ,hd.userid,hd.validator,hd.cc1,hd.cc2,hd.cc3,hd.cc4,hd.otheremail,hd.latestupdate,dt.id as doctypeid ,vd.id as vdid,vd.documentid as documentid" &
                 " from doc.document d" &
                 " left join doc.vendordoc vd on vd.documentid = d.id left join doc.header hd on hd.id = vd.headerid left join doc.user o on lower(o.userid) = hd.userid  left join officerseb o2 on lower(o2.userid) = hd.validator left join officerseb o3 on lower(o3.userid) = hd.cc1 left join officerseb o4 on lower(o4.userid) = hd.cc2 left join officerseb o5 on lower(o5.userid) = hd.cc3 left join officerseb o6 on lower(o6.userid) = hd.cc4 inner join doc.docexpired de on d.id = de.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid left join doc.user u on u.userid = hd.userid " &
                 " left join vendor v on v.vendorcode = vd.vendorcode left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode left join vst on vst.id = vs.status left join doc.doctype dt on dt.id = d.doctypeid left join doc.project p on p.documentid = d.id left join doc.version vr on vr.documentid = d.id where vd.status < 3 and (expireddate - current_date  <= dr.reminder) and (de.status isnull or de.status <> 3) and " &
                 " (hd.userid in (select '" & myuser & "' union all" &
                " select o.userid from doc.emailmapping e" &
                " left join doc.user o on o.email =  e.oldemail" &
                " left join doc.user n on n.email = e.newemail" &
                " where lower(n.userid) = '" & myuser & "') or hd.validator = '" & myuser & "' or hd.cc1 = '" & myuser & "' or hd.cc2 = '" & myuser & "' or hd.cc3 = '" & myuser & "' or hd.cc4 = '" & myuser & "'  or " & (HelperClass1.UserInfo.IsAdmin) & " or " & (HelperClass1.UserInfo.AllowUpdateDocument) & ")  and not hd.id isnull) as foo {0} union all ", CriteriaNew))
        'Follow up
        'sb.Append(String.Format(" select * from (" &
        '                        "with ct as (with rm as (select * from doc.doctypereminder)," &
        '                        " so as (select max(so.id) as id, rm.id as doctypeid,documentid from doc.supplychainother so left join rm  on rm.id = 35 where otherdate - now()::date <= rm.reminder and so.status = true group by documentid,doctypeid)," &
        '                        " go as (select max(go.id) as id,  rm.id as doctypeid,documentid from doc.generalcontractother go left join rm  on rm.id = 32 where otherdate - now()::date <= rm.reminder and go.status = true group by documentid,doctypeid)," &
        '                        " qo as (select max(qo.id) as id,  rm.id as doctypeid,documentid from doc.qualityappendixother qo left join rm  on rm.id = 33 where otherdate - now()::date <= rm.reminder and qo.status = true group by documentid,doctypeid) " &
        '                        " select * from go union all select * from so union all select * from qo ), " &
        '                        " vst as(select paramname as statusname ,ivalue as id from doc.paramdt where paramhdid = 2)" &
        '         " select hd.id,'Follow up'::text as statusname ,hd.creationdate,u.userid as username,v.shortname::text,v.vendorname::text as suppliername,v.vendorcode::text,vst.statusname::text as vstatusname,dt.doctypename,p.projectname,vr.version,d.docdate,null::date,ct.id as statusexp,o2.officersebname::character varying as validatorname,o3.officersebname::character varying as cc1name ,o4.officersebname::character varying as cc2name,o5.officersebname::character varying as cc3name,o6.officersebname::character varying as cc4name ,hd.userid,hd.validator,hd.cc1,hd.cc2,hd.cc3,hd.cc4,hd.otheremail,hd.latestupdate,dt.id as doctypeid ,vd.id as vdid,vd.documentid as documentid" &
        '        " from ct left join doc.document d on d.id = ct.documentid" &
        '         " left join doc.vendordoc vd on vd.documentid = d.id left join doc.header hd on hd.id = vd.headerid left join doc.user o on lower(o.userid) = hd.userid  left join officerseb o2 on lower(o2.userid) = hd.validator left join officerseb o3 on lower(o3.userid) = hd.cc1 left join officerseb o4 on lower(o4.userid) = hd.cc2 left join officerseb o5 on lower(o5.userid) = hd.cc3 left join officerseb o6 on lower(o6.userid) = hd.cc4 " &
        '         " left join doc.doctypereminder dr on dr.id = d.doctypeid left join doc.user u on u.userid = hd.userid " &
        '         " left join vendor v on v.vendorcode = vd.vendorcode left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode left join vst on vst.id = vs.status left join doc.doctype dt on dt.id = d.doctypeid left join doc.project p on p.documentid = d.id left join doc.version vr on vr.documentid = d.id where  " &
        '         " (hd.userid in (select '" & myuser & "' union all" &
        '        " select o.userid from doc.emailmapping e" &
        '        " left join doc.user o on o.email =  e.oldemail" &
        '        " left join doc.user n on n.email = e.newemail" &
        '        " where lower(n.userid) = '" & myuser & "') or hd.validator = '" & myuser & "' or hd.cc1 = '" & myuser & "' or hd.cc2 = '" & myuser & "' or hd.cc3 = '" & myuser & "' or hd.cc4 = '" & myuser & "'  or " & (HelperClass1.UserInfo.IsAdmin) & " or " & (HelperClass1.UserInfo.AllowUpdateDocument) & ")  and not hd.id isnull) as foo {0} ;", CriteriaNew))


        sb.Append(String.Format(" select * from (" &
                                "with ct as (with rm as (select * from doc.doctypereminder)," &
                                " so as (select max(so.id) as id, rm.id as doctypeid,documentid from doc.supplychainother so left join rm  on rm.id = 35 where otherdate - now()::date <= rm.reminder and so.status = true group by documentid,doctypeid)," &
                                " go as (select max(go.id) as id,  rm.id as doctypeid,documentid from doc.generalcontractother go left join rm  on rm.id = 32 where otherdate - now()::date <= rm.reminder and go.status = true group by documentid,doctypeid)," &
                                " qo as (select max(qo.id) as id,  rm.id as doctypeid,documentid from doc.qualityappendixother qo left join rm  on rm.id = 33 where otherdate - now()::date <= rm.reminder and qo.status = true group by documentid,doctypeid) " &
                                " select * from go union all select * from so union all select * from qo ), " &
                                " vst as(select paramname as statusname ,ivalue as id from doc.paramdt where paramhdid = 2)" &
                 " select hd.id,'Follow up'::text as statusname ,hd.creationdate,u.userid as username,v.shortname::text,v.vendorname::text as suppliername,v.vendorcode::text,vst.statusname::text as vstatusname,dt.doctypename,p.projectname,vr.version,d.docdate,null::date,ct.id as statusexp,o2.officersebname::character varying as validatorname,o3.officersebname::character varying as cc1name ,o4.officersebname::character varying as cc2name,o5.officersebname::character varying as cc3name,o6.officersebname::character varying as cc4name ,hd.userid,hd.validator,hd.cc1,hd.cc2,hd.cc3,hd.cc4,hd.otheremail,hd.latestupdate,dt.id as doctypeid ,vd.id as vdid,vd.documentid as documentid" &
                " from ct left join doc.document d on d.id = ct.documentid" &
                 " left join doc.vendordoc vd on vd.documentid = d.id left join doc.header hd on hd.id = vd.headerid left join doc.user o on lower(o.userid) = hd.userid  left join officerseb o2 on lower(o2.userid) = hd.validator left join officerseb o3 on lower(o3.userid) = hd.cc1 left join officerseb o4 on lower(o4.userid) = hd.cc2 left join officerseb o5 on lower(o5.userid) = hd.cc3 left join officerseb o6 on lower(o6.userid) = hd.cc4 " &
                 " left join doc.doctypereminder dr on dr.id = d.doctypeid left join doc.user u on u.userid = hd.userid " &
                 " left join vendor v on v.vendorcode = vd.vendorcode left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode left join vst on vst.id = vs.status left join doc.doctype dt on dt.id =ct.doctypeid left join doc.project p on p.documentid = d.id left join doc.version vr on vr.documentid = d.id where  " &
                 " (hd.userid in (select '" & myuser & "' union all" &
                " select o.userid from doc.emailmapping e" &
                " left join doc.user o on o.email =  e.oldemail" &
                " left join doc.user n on n.email = e.newemail" &
                " where lower(n.userid) = '" & myuser & "') or hd.validator = '" & myuser & "' or hd.cc1 = '" & myuser & "' or hd.cc2 = '" & myuser & "' or hd.cc3 = '" & myuser & "' or hd.cc4 = '" & myuser & "'  or " & (HelperClass1.UserInfo.IsAdmin) & " or " & (HelperClass1.UserInfo.AllowUpdateDocument) & ")  and not hd.id isnull) as foo {0} ;", CriteriaNew))

        sb.Append("select * from officerseb o  where lower(o.userid) = '" & myuser & "' limit 1;")
        'sb.Append("with vd as ( select distinct headerid,status from doc.vendordoc where not status isnull ) Select h.id, 'New'::text as statusname,h.creationdate,o.username::character varying as username,doc.sp_getshortname(h.id) as shortname,doc.sp_getsuppliers(h.id) as suppliername ,doc.sp_getvendorcode(h.id)::text as vendorcode, null::text as vstatusname,doc.sp_getdoctypename(h.id)::character varying as doctypename,null::character varying as projectname,null::character varying as version, null::date as docdate,null::date as expireddate,null::integer as statusexp,o2.officersebname::character varying as validatorname,o3.officersebname::character varying as cc1name ,o4.officersebname::character varying as cc2name,o5.officersebname::character varying as cc3name,o6.officersebname::character varying as cc4name ,h.userid,h.validator,h.cc1,h.cc2,h.cc3,h.cc4,h.otheremail,h.latestupdate" &
        '          " from doc.header h  left join doc.user o on lower(o.userid) = h.userid left join officerseb o2 on lower(o2.userid) = h.validator left join officerseb o3 on lower(o3.userid) = h.cc1 left join officerseb o4 on lower(o4.userid) = h.cc2 left join officerseb o5 on lower(o5.userid) = h.cc3 left join officerseb o6 on lower(o6.userid) = h.cc4 left join vd on vd.headerid = h.id" &
        '          " where (h.userid = '" & myuser & "' or h.validator = '" & myuser & "' or h.cc1 = '" & myuser & "' or h.cc2 = '" & myuser & "' or h.cc3 = '" & myuser & "' or h.cc4 = '" & myuser & "'  or " & HelperClass1.UserInfo.IsAdmin & " or " & HelperClass1.UserInfo.AllowUpdateDocument & ") and not h.id isnull")


        sb.Append(String.Format("select * from (with vst as(select paramname as statusname ,ivalue as id from doc.paramdt where paramhdid = 2) " &
                                " select hd.id,doc.getstatusdoc(vd.status)::text as statusname ,hd.creationdate,u.userid as username,v.shortname::text,v.vendorname::text as suppliername,v.vendorcode,vst.statusname::text as vstatusname,dt.doctypename,p.projectname,vr.version,d.docdate,de.expireddate,de.status as statusexp,o2.officersebname::character varying as validatorname,o3.officersebname::character varying as cc1name ,o4.officersebname::character varying as cc2name,o5.officersebname::character varying as cc3name,o6.officersebname::character varying as cc4name ,hd.userid,hd.validator,hd.cc1,hd.cc2,hd.cc3,hd.cc4,hd.otheremail,hd.latestupdate,vd.id as vdid " &
                  "from doc.document d left join doc.vendordoc vd on vd.documentid = d.id left join doc.header hd on hd.id = vd.headerid left join doc.user o on lower(o.userid) = hd.userid  left join officerseb o2 on lower(o2.userid) = hd.validator left join officerseb o3 on lower(o3.userid) = hd.cc1 left join officerseb o4 on lower(o4.userid) = hd.cc2 left join officerseb o5 on lower(o5.userid) = hd.cc3 left join officerseb o6 on lower(o6.userid) = hd.cc4 left join doc.docexpired de on d.id = de.documentid left join doc.doctypereminder dr on dr.id = d.doctypeid left join doc.user u on u.userid = hd.userid  left join vendor v on v.vendorcode = vd.vendorcode left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
                  " left join vst on vst.id = vs.status left join doc.doctype dt on dt.id = d.doctypeid left join doc.project p on p.documentid = d.id left join doc.version vr on vr.documentid = d.id " &
                  " where (hd.userid = '" & myuser & "' or hd.validator = '" & myuser & "' or hd.cc1 = '" & myuser & "' or hd.cc2 = '" & myuser & "' or hd.cc3 = '" & myuser & "' or hd.cc4 = '" & myuser & "'  or " & (HelperClass1.UserInfo.IsAdmin) & " or " & (HelperClass1.UserInfo.AllowUpdateDocument) & ")  and not hd.id isnull " &
                  " order by creationdate desc) as foo {0} {1};", CriteriaNew, limitHistory))
        'sb.Append(String.Format("select * from doc.user where allowupdatedocument and isactive and id <> 1 order by username;"))

        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                Dim pk(2) As DataColumn
                pk(0) = DS.Tables(0).Columns("id")
                pk(1) = DS.Tables(0).Columns("vdid")
                pk(2) = DS.Tables(0).Columns("doctypeid")
                DS.Tables(0).TableName = "Header"
                DS.Tables(0).PrimaryKey = pk
                DS.Tables(0).Columns("id").AutoIncrement = True
                DS.Tables(0).Columns("id").AutoIncrementSeed = -1
                DS.Tables(0).Columns("id").AutoIncrementStep = -1
                ProgressReport(4, "InitData")
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
        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
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
                            bsheader = New BindingSource
                            bshistory = New BindingSource

                            'Dim pk(0) As DataColumn
                            'pk(0) = DS.Tables(0).Columns("id")
                            'DS.Tables(0).PrimaryKey = pk
                            'DS.Tables(0).Columns("id").AutoIncrement = True
                            'DS.Tables(0).Columns("id").AutoIncrementSeed = 0
                            'DS.Tables(0).Columns("id").AutoIncrementStep = -1
                            DS.Tables(0).TableName = "Header"
                            'ToolStripComboBox1.SelectedIndex = 0

                            bsheader.DataSource = DS.Tables(0)
                            bshistory.DataSource = DS.Tables(2)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = bsheader

                            DataGridView2.AutoGenerateColumns = False
                            DataGridView2.DataSource = bshistory


                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                    Case 11
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub FormMyTask_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
        Button4.Visible = HelperClass1.UserInfo.IsAdmin
        DataGridView1.Columns("column2").ReadOnly = Not (HelperClass1.UserInfo.IsAdmin)
        'ComboBox1.SelectedIndex = 1
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        'MessageBox.Show(DataGridView1.Columns(e.ColumnIndex).HeaderText)
    End Sub

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        Dim tb As DataGridViewTextBoxEditingControl = DirectCast(e.Control, DataGridViewTextBoxEditingControl)
        RemoveHandler (tb.KeyDown), AddressOf datagridviewTextBox_Keypdown
        AddHandler (tb.KeyDown), AddressOf datagridviewTextBox_Keypdown
    End Sub

    Private Sub datagridviewTextBox_Keypdown(ByVal sender As Object, ByVal e As KeyEventArgs)
        If e.KeyValue = 112 Then 'F1 
            MessageBox.Show("Help")
        End If
    End Sub



    'Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
    '    If Not myThread.IsAlive Then
    '        Dim myrow As DataRowView = bsheader.Current
    '        Dim myform = New FormPriceChange(bsheader, DS, False)
    '        Select Case myrow.Row.Item("status")
    '            Case 2, 4
    '                myform.ToolStripButton2.Visible = False
    '                myform.ToolStripButton7.Visible = False
    '            Case 3
    '                myform.ToolStripButton4.Visible = False
    '                myform.ToolStripButton5.Visible = False
    '            Case 5
    '                myform.ToolStripButton2.Visible = False
    '                myform.ToolStripButton4.Visible = False
    '                myform.ToolStripButton5.Visible = False
    '        End Select

    '        If Not myform.ShowDialog = DialogResult.OK Then

    '            bsheader.CancelEdit()
    '        Else
    '            bsheader.EndEdit()
    '        End If
    '        loaddata()
    '    End If

    'End Sub



    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not myThread.IsAlive Then
            loaddata()
        End If
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If Not IsNothing(bsheader.Current) Then
            If Not myThread.IsAlive Then
                Dim drv As DataRowView = bsheader.Current
                Dim myrole As Role
                myrole = HelperClass1.getRole(drv, HelperClass1)
                'Dim myform = New FormDocumentVendor(drv.Row.Item("id"), myrole)
                'Dim myform = New FormDocumentVendor2(drv.Row.Item("id"), myrole)
                'Dim myform = New FormDocumentVendor2(drv, myrole)
                'Dim myform = New FormDocumentVendor3(drv, myrole)
                'Dim myform = New FormDocumentVendor4(drv, myrole)
                'Dim myform = New FormDocumentVendor5(drv, myrole)
                'Dim myform = New FormDocumentVendor6(drv, myrole)
                Dim myform = New FormDocumentVendor7(drv, myrole)
                If myform.ShowDialog = DialogResult.OK Then
                    'If Follow up then Flag True for contract(GO,SO,QO) based on documentid 
                    If drv.Item("statusname") = "Follow up" Then


                        Dim mytable As String = String.Empty
                        Select Case drv.Item("doctypeid")
                            Case 32
                                mytable = "doc.generalcontractother"
                            Case 33
                                mytable = "doc.qualityappendixother"
                            Case 35
                                mytable = "doc.supplychainother"
                        End Select
                        Dim sqlstr = String.Format("update {0} set status = false where documentid = {1} and id = {2}", mytable, drv.Item("documentid"), drv.Item("statusexp"))
                        DbAdapter1.ExecuteNonQuery(sqlstr)
                    End If
                    loaddata()
                Else
                End If

            Else
                MessageBox.Show("Still loading... Please wait.")
            End If
        End If
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim mydrv = bsheader.Current
        If mydrv.item("statusname") <> "New" Then
            MessageBox.Show("Sorry. This record cannot be deleted.")
            Exit Sub
        End If

        If MessageBox.Show("Delete selected record?", "Delete Record(s)", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                If drv.Cells("Status").ToString = "New" Then
                    bsheader.RemoveAt(drv.Index)
                End If


            Next
            Dim ds2 As DataSet
            ds2 = DS.GetChanges
            If Not IsNothing(ds2) Then
                Dim mymessage As String = String.Empty
                Dim ra As Integer
                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                If DbAdapter1.DocumentVendorTx(Me, mye) Then
                    'delete original Dataset (DS) for those table having added record -> Merged with modified Dataset (DS2)
                    'For update record, no need to delete the original dataset (DS) because the id is the same. 
                    'Why need to delete the added one, because when we create new record, the id started with 0,-1,-2 and so on.
                    'when we update to database, we put the real id from database.
                    'so we have different value id for DS and DS2. if we do merged without deleting the original one, we will have 2 records.
                    For i = 0 To 1
                        Dim modifiedRows = From row In DS.Tables(i)
                            Where row.RowState = DataRowState.Modified Or row.RowState = DataRowState.Added
                        For Each row In modifiedRows.ToArray
                            row.Delete()
                        Next
                    Next
                Else
                    MessageBox.Show(mye.message)
                    Exit Sub
                End If
                DS.Merge(ds2)
                DS.AcceptChanges()
                MessageBox.Show("Saved.")
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ToolStripButton4.PerformClick()
    End Sub

    Private Sub AddRecordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddRecordToolStripMenuItem.Click
        ToolStripButton1.PerformClick()
    End Sub

    Private Sub DeleteRecordToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteRecordToolStripMenuItem1.Click
        ToolStripButton2.PerformClick()
    End Sub


    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        'runfilter()

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        'runfilter()
    End Sub

    Private Sub runfilterWork()

        'ProgressReport(6, "Marquee")
        'ProgressReport(1, "Executing Filter...")
        ProgressReport(11, "Execute Filter")
        'ProgressReport(1, "Execute Filter Done.")
        'ProgressReport(5, "Continuous")
    End Sub

    Private Sub runfilter()

        'If Not myFilter.IsAlive Then
        '    myFilter = New Thread(AddressOf runfilterWork)
        '    myFilter.Start()
        'Else
        '    'MessageBox.Show("Still loading... Please wait.")
        'End If

        'bsheader.Filter = ""
        'bshistory.Filter = ""
        insertAnd = ""
        Criteria = ""
        CriteriaNew = ""
        If ToolStripTextBox1.Text <> "" Then
            'bsheader.Filter = myListArray(ToolStripComboBox1.SelectedIndex) & " like '*" & ToolStripTextBox1.Text.Replace("'", "''") & "*'"
            'bshistory.Filter = myListArray(ToolStripComboBox1.SelectedIndex) & " like '*" & ToolStripTextBox1.Text.Replace("'", "''") & "*'"
            Criteria = myListArray(ToolStripComboBox1.SelectedIndex) & " like '%" & ToolStripTextBox1.Text.ToLower.Replace("'", "''").Replace("\", "\\") & "%'"
        End If
        If Criteria <> "" Then
            CriteriaNew = " where "
        End If
        CriteriaNew = CriteriaNew & Criteria
        Criteria = " and " & Criteria
        If EPCheckBox1.Checked Then
            'If bsheader.Filter <> "" Then
            'insertAnd = " and "
            'End If
            'bsheader.Filter = bsheader.Filter & String.Format(" {0} (expireddate >= '{1:yyyy-MM-dd}' and expireddate <= '{2:yyyy-MM-dd}')", insertAnd, dtpicker1.Value, dtpicker2.Value)
            'bshistory.Filter = bshistory.Filter & String.Format(" {0} (expireddate >= '{1:yyyy-MM-dd}' and expireddate <= '{2:yyyy-MM-dd}')", insertAnd, dtpicker1.Value, dtpicker2.Value)
            'If Criteria <> "" Then
            If CriteriaNew = "" Then
                insertAnd = " where "
            Else
                insertAnd = " and "
            End If

            'End If
            CriteriaNew = CriteriaNew & String.Format(" {0} (expireddate >= '{1:yyyy-MM-dd}' and expireddate <= '{2:yyyy-MM-dd}')", insertAnd, dtpicker1.Value, dtpicker2.Value)
        End If


    End Sub



    Private Sub DataGridView2_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        If Not IsNothing(bshistory.Current) Then
            If Not myThread.IsAlive Then
                Dim drv As DataRowView = bshistory.Current
                Dim myrole As Role
                myrole = HelperClass1.getRole(drv, HelperClass1)
                'Dim myform = New FormDocumentVendor(drv.Row.Item("id"), myrole)
                'Dim myform = New FormDocumentVendor2(drv.Row.Item("id"), myrole)
                'Dim myform = New FormDocumentVendor2(drv, myrole)
                'Dim myform = New FormDocumentVendor3(drv, myrole)
                'Dim myform = New FormDocumentVendor4(drv, myrole)
                'Dim myform = New FormDocumentVendor5(drv, myrole)
                'Dim myform = New FormDocumentVendor6(drv, myrole)
                Dim myform = New FormDocumentVendor7(drv, myrole)
                If myform.ShowDialog = DialogResult.OK Then
                    loaddata()
                Else
                End If

            Else
                MessageBox.Show("Still loading... Please wait.")
            End If
        End If
    End Sub



    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        limitHistory = ""
        If ComboBox1.Text = "100" Then
            limitHistory = "limit 100"
        End If
        loaddata()
    End Sub

    Private Sub EPCheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles EPCheckBox1.CheckedChanged
        'button1.enable = EPCheckBox1.Checked
        dtpicker1.Enabled = EPCheckBox1.Checked
        lbl1.Enabled = EPCheckBox1.Checked
        dtpicker2.Enabled = EPCheckBox1.Checked

        ToolStrip1.Invalidate()
        If Not myThread.IsAlive Then
            ' If EPCheckBox1.Checked = False Then runfilter()
        End If

    End Sub

    Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
        runfilter()
        loaddata()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click, Button2.Click
        Dim drv = bsheader.Current
        If drv.item("statusname") = "Expired" Then
            ' drv.item("status") = 3 'renew completed
            'Dim myform = New FormDocumentVendor()
            'Dim myform = New FormDocumentVendor2()
            'Dim myform = New FormDocumentVendor3()
            'Dim myform = New FormDocumentVendor4()
            'Dim myform = New FormDocumentVendor5()
            'Dim myform = New FormDocumentVendor6()
            Dim myform = New FormDocumentVendor7()
            myform.RenewDRV = drv
            myform.loaddata(0)
            If myform.ShowDialog = DialogResult.OK Then
                'update existing to Renew Completed
                '
                'Dim sqlstr = String.Format("update doc.vendordoc set status = 3 where id = {0};", drv.item("vdid"))

                Dim sqlstr = String.Format("update doc.vendordoc set status = 3 where headerid = {0}  and documentid = {1};", drv.item("id"), drv.Item("documentid"))
                Dim errmessage As String = String.Empty
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=errmessage) Then
                    MessageBox.Show(errmessage)
                End If
                loaddata()
            Else
            End If


        Else
            MessageBox.Show("This record cannot be renewed.")
        End If
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click, Button3.Click
        'Open FormDocumentVendor based on Header Id
        'find vendordoc based on MyTaskDocument vendordocid, set status = 4.
        Dim drv = bsheader.Current
        If drv.item("statusname") = "Expired" Then
            ' drv.item("status") = 3 'renew completed
            'Dim myform = New FormDocumentVendor()
            'Dim myform = New FormDocumentVendor2()
            'Dim myform = New FormDocumentVendor3()
            'Dim myform = New FormDocumentVendor4()
            'Dim myform = New FormDocumentVendor6()
            Dim myform = New FormDocumentVendor7()
            myform.headerid = drv.item("id")
            myform.RenewDRV = drv

            myform.loaddata(drv.item("id"))
            If myform.ShowDialog = DialogResult.OK Then
                'update existing to Renew Completed
                '
                'Dim sqlstr = String.Format("update doc.vendordoc set status = 3 where id = {0};", drv.item("vdid"))
                'Dim errmessage As String = String.Empty
                'If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=errmessage) Then
                '    MessageBox.Show(errmessage)
                'End If
                loaddata()
            Else
            End If


        Else
            ' MessageBox.Show("This record cannot change status into nonrenewable.")
            MessageBox.Show("No Renew not allowed for this record.")
        End If
    End Sub





    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Validate()
        bsheader.EndEdit()

        Dim ds2 As DataSet
        ds2 = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If Not DbAdapter1.TaskDocumentForwardUser(Me, mye) Then

                MessageBox.Show(mye.message)
                Exit Sub
            End If
            DS.Merge(ds2)
            DS.AcceptChanges()
            MessageBox.Show("Saved.")
        End If

    End Sub



    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class