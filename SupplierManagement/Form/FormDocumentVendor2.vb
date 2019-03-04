Imports SupplierManagement.SharedClass
Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Public Class FormDocumentVendor2
    Private Const SIF As Integer = 21
    Private Const IDENTITY_SHEET As Integer = 54

    Dim WithEvents bsheader As BindingSource
    Dim WithEvents bsDetail As BindingSource

    Dim bsShortname As BindingSource
    Dim bsShortnameHelper As BindingSource
    Dim bsVendorname As BindingSource
    Dim bsVendornameHelper As BindingSource
    Dim bsShortnameVendor As BindingSource
    Dim WithEvents bsDocType As BindingSource
    Dim bsDocTypeHelper As BindingSource
    Dim bsDocLevel As BindingSource
    Dim bsPaymentTerm As BindingSource
    Dim bsfolder As BindingSource

    Dim bsSiTx As BindingSource

    Dim auditbylist As String() = {"", "SGS", "SEB Asia", "Intertek", "Waived Audit"}
    Dim auditTypelist As String() = {"", "Initial", "1st follow up", "2nd follow up", "3rd follow up", "4th follow up"}
    Dim auditGradelist As String() = {"", "Minor", "Major", "Critical", "Zero Tolerance NC", "Red Level", "Orange Level", "Yellow Level", "Green Level"}
    Dim overallAuditList As String() = {"", "Zero Tolerance", "Red Level", "Orange Level", "Yellow Level", "Green Level"}
    Dim zeroList As String() = {"", "Labor", "Wages & Hours", "Health & Safety", "Management Systems", "Environment"}

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Public Property DS As DataSet

    Dim sb As New StringBuilder
    Dim myuser As String
    Public headerid As Long
    Dim validatorid As String

    Dim cc1id As String
    Dim cc2id As String
    Dim cc3id As String
    Dim cc4id As String
    Dim isSave As Boolean = False
    Dim isRole As Role
    Protected Friend RenewDRV As DataRowView
    Dim recPosition As Integer = 0
    Dim VDID As Integer

    Public Sub New(ByVal headerid As Long, ByVal isRole As Role)
        InitializeComponent()
        'Update
        Me.headerid = headerid
        Me.isRole = isRole
        loaddata(headerid)
    End Sub
    Public Sub New(ByVal headerid As Long, ByVal VDID As Long, ByVal isRole As Role)
        InitializeComponent()
        'Update
        Me.headerid = headerid
        Me.isRole = isRole
        loaddata(headerid)
    End Sub
    Public Sub New(ByVal drv As DataRowView, ByVal isRole As Role)
        InitializeComponent()
        'Update
        Me.headerid = drv.Item("id")
        Me.isRole = isRole
        If Not IsDBNull(drv.Item("vdid")) Then
            VDID = drv.Item("vdid")
        End If
        loaddata(headerid)
    End Sub
    Public Sub New(ByVal headerid As Long)
        InitializeComponent()
        'Update
        Me.headerid = headerid
        loaddata(headerid)
    End Sub


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'Create New
        headerid = 0
        'loaddata(headerid)


    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        If MessageBox.Show("Cancel current task?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.Yes Then
            bsheader.CancelEdit()
            Me.DialogResult = DialogResult.Cancel
        End If

    End Sub
    Private Sub FormCutoff_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not Me.DialogResult = DialogResult.Cancel Then
            If Not isRole = Role.Other Then
                If Not isSave Then
                    Dim abc = getChanges()
                    If Not IsNothing(abc) Then
                        If abc.Tables(1).Rows.Count <> 0 Then
                            Select Case MessageBox.Show("Save unsaved records?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                                Case Windows.Forms.DialogResult.Yes
                                    If Me.validate Then
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
            End If
        End If


    End Sub

    Private Function getChanges() As DataSet
        If IsNothing(bsDetail) Then
            Return Nothing
        End If
        bsheader.EndEdit()
        bsDetail.EndEdit()
        Return DS.GetChanges()
    End Function

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Logger.log("**Start Submit***")
        If Me.validate() Then
            'remove RowError caused by Document date Today's Date 
            For Each drv As DataRowView In bsDetail.List
                drv.Row.RowError = ""
            Next
            DataGridView1.Invalidate()

            bsheader.EndEdit()
            bsDetail.EndEdit()
            Try
                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                Dim ds2 As DataSet
                ds2 = DS.GetChanges
                Logger.log(String.Format("**DS2 is nothing {0}**", IsNothing(ds2)))
                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)

                    Logger.log("** Delete Original File **")
                    For Each drv As DataRowView In bsDetail.List
                        If Not IsDBNull(drv.Row.Item("filename")) Then
                            Try
                                If Not IsDBNull(drv.Row.Item("docname", DataRowVersion.Original)) Then
                                    Dim orifile = DS.Tables(9).Rows(0).Item("cvalue") & "\" & drv.Row.Item("documentid") & drv.Row.Item("docext", DataRowVersion.Original)
                                    If FileIO.FileSystem.FileExists(orifile) Then
                                        FileIO.FileSystem.DeleteFile(orifile)
                                    End If

                                End If
                            Catch ex As Exception
                                Logger.log(String.Format("** Delete Original File error {0}**", ex.Message))
                            End Try
                        End If
                    Next


                    Logger.log("**DbAdapter**")
                    If DbAdapter1.DocumentVendorTx(Me, mye) Then
                        isSave = True
                        Logger.log("**DbAdapter True**")
                        'delete original Dataset (DS) for those table having added record -> Merged with modified Dataset (DS2)
                        'For update record, no need to delete the original dataset (DS) because the id is the same. 
                        'Why need to delete the added one, because when we create new record, the id started with 0,-1,-2 and so on.
                        'when we update to database, we put the real id from database.
                        'so we have different value id for DS and DS2. if we do merged without deleting the original one, we will have 2 records.
                        For i = 0 To 1

                            Dim modifiedRows = From row In DS.Tables(i)
                                Where row.RowState = DataRowState.Added 'Or row.RowState = DataRowState.Modified
                            For Each row In modifiedRows.ToArray
                                Logger.log("**For Row Delete**")
                                Try
                                    row.Delete()
                                Catch ex As Exception
                                    Logger.log(String.Format("** Row Delete error {0}**", ex.Message))
                                End Try
                            Next
                        Next
                        Logger.log("**Row Deleted**")
                    Else
                        MessageBox.Show(mye.message)
                        Exit Sub
                    End If
                    Logger.log("**DS Merged**")

                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    DataGridView1.Invalidate()
                    MessageBox.Show("Saved.")
                End If

                'copy file
                Logger.log("**Copy File**")

                For Each drv As DataRowView In bsDetail.List
                    If Not IsDBNull(drv.Row.Item("filename")) Then
                        Dim mytarget = DS.Tables(9).Rows(0).Item("cvalue") & "\" & drv.Row.Item("documentid") & drv.Row.Item("docext")
                        'Delete original File


                        'Add New One
                        Try
                            FileIO.FileSystem.CopyFile(drv.Row.Item("filename"), mytarget, True)
                        Catch ex As Exception
                            Logger.log(String.Format("** CopyFile error {0}**", ex.Message))
                            MessageBox.Show(String.Format("** CopyFile error {0}**", ex.Message))
                        End Try
                    End If
                Next
                'create record
                Me.DialogResult = DialogResult.OK
            Catch ex As Exception
                Logger.log(String.Format("**Error {0}**", ex.Message))
                MessageBox.Show(" Error:: " & ex.Message)
            End Try



        Else
            Logger.log("**Validate False**")
            'bsheader.CancelEdit()
            'Me.DialogResult = DialogResult.Cancel
        End If
        Logger.log("**End Submit***")
    End Sub
    Public Overloads Function validate() As Boolean
        Logger.log("**Start Validate***")
        Dim myret As Boolean = True
        MyBase.Validate()
        Dim myheader As DataRowView = bsheader.Current
        myheader.BeginEdit()
        myheader.Row.Item("latestupdate") = Date.Now
        Dim foundTodayDocdate As Boolean = False
        For Each drv As DataRowView In bsDetail.List

            'drv.Row.RowError = "Has Error loh"
            'Check Status based on Validator blank or not
            'if blank status = 0
            Logger.log("**For each***")
            If TextBox2.Text <> "" Then
                'validator is not blank status Completed
                If IsDBNull(drv.Row.Item("status")) Then
                    drv.Row.Item("status") = 1
                Else
                    If drv.Row.Item("status") < 2 Then
                        drv.Row.Item("status") = 1
                        drv.Row.Item("statusname") = "NEW"
                    End If
                End If



            Else
                'validator is blank status = new
                If IsDBNull(drv.Row.Item("status")) Then
                    drv.Row.Item("status") = 2
                Else
                    If drv.Row.Item("status") = 4 Or drv.Row.Item("status") = 3 Then
                        'skip
                    ElseIf drv.Row.Item("status") < 2 Then
                        drv.Row.Item("status") = 2
                        drv.Row.Item("statusname") = "COMPLETED"
                    End If
                End If

            End If

            If Not validaterecord(drv, foundTodayDocdate) Then
                myret = False
            End If
        Next

        DataGridView1.Invalidate()
        If foundTodayDocdate Then
            If MessageBox.Show("Found Document date is equal Today's date. Continue?", "Please confirm.", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Cancel Then
                myret = False
            End If
        End If
        Logger.log(String.Format("**End Validate Myret {0}**", myret))
        Return myret
    End Function

    Private Function validaterecord(ByVal drv As DataRowView, ByRef foundTodayDocdate As Boolean) As Boolean
        Dim myerror As New StringBuilder
        Dim myret As Boolean = True
        Try


            If IsDBNull(drv.Row.Item("doctypename")) Then
                myerror.Append("Document Type cannot be blank.")
                myret = False
            End If

            'Document type
            If IsDBNull(drv.Row.Item("vendorcode")) And IsDBNull(drv.Row.Item("shortname")) Then
                myerror.Append("Vendor Name and Shortname cannot be blank.")
                myret = False
            End If

            If Not IsDBNull(drv.Row.Item("doctypename")) Then
                Select Case drv.Row.Item("doctypename")
                    Case "Contract: General Contract"
                        'If ComboBox5.SelectedIndex < 1 Then
                        If IsDBNull(drv.Row.Item("paymentcode")) Then
                            myerror.Append("Please select payment term!")
                            myret = False
                        End If
                    Case "Contract: Supply Chain Appendix"
                        'If TextBox11.Text = "" Then
                        If IsDBNull(drv.Row.Item("leadtime")) Then
                            myerror.Append("Lead time cannot be blank!")
                            myret = False
                        End If
                        'If TextBox12.Text = "" Then
                        If IsDBNull(drv.Row.Item("sasl")) Then
                            myerror.Append("SASL cannot be blank!")
                            myret = False
                        End If
                    Case "Contract: Quality Appendix"
                        'If TextBox13.Text = "" Then
                        If IsDBNull(drv.Row.Item("nqsu")) Then
                            myerror.Append("NQSU cannot be blank!")
                            myret = False
                        End If
                    Case "SEF"
                        'Score 0-100
                        If IsDBNull(drv.Row.Item("score")) Then
                            myerror.Append("Score cannot be blank!")
                            myret = False
                        ElseIf drv.Row.Item("score") < 0 Or drv.Row.Item("score") > 100 Then
                            myerror.Append("Score Range from 0 - 100!")
                            myret = False
                        End If
                    Case "Social Audit"
                        If IsDBNull(drv.Row.Item("auditby")) Or IsDBNull(drv.Row.Item("audittype")) Or IsDBNull(drv.Row.Item("auditgrade")) Then
                            myerror.Append("Auditby or Audittype or Auditgrade cannot be blank!")
                            myret = False
                        End If
                End Select
            End If


            'Document Level
            If Not IsDBNull(drv.Row.Item("levelname")) Then
                Select Case drv.Row.Item("levelname")
                    Case ""
                        myerror.Append("Document Level cannot be blank!")
                        myret = False
                    Case "Product / Project"
                        If IsDBNull(drv.Row.Item("projectname")) Then
                            myerror.Append("Project Name cannot be blank!")
                            myret = False
                        End If
                End Select
            End If

            'filename cannot be blank
            If IsDBNull(drv.Row.Item("docname")) Then
                myerror.Append("File Name cannot be blank!")
                myret = False
            End If

            'If DateTimePicker3.Checked Then
            '    If DateTimePicker3.Value.Subtract(DateTimePicker2.Value).TotalDays <= 0 Then
            '        myerror.Append("Expired date must be after the Document date!")
            '        myret = False
            '    End If
            'End If

            If Not IsDBNull(drv.Row.Item("expireddate")) Then
                If CType(drv.Row.Item("expireddate"), Date).Subtract(drv.Row.Item("docdate")).totaldays <= 0 Then
                    myerror.Append("Expired date must be after the Document date!")
                    myret = False
                End If
            End If

            If (drv.Row.Item("docdate") = Today.Date) Then
                myerror.Append("Document date is equal Today's date! Please confirm.")
                foundTodayDocdate = True
            End If

            drv.Row.RowError = myerror.ToString
        Catch ex As Exception
            myret = False
            MessageBox.Show("Validate::" & ex.Message)
        End Try
        Return myret

    End Function
    Public Sub loaddata(ByVal id As Long)
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
        'ProgressReport(4, "InitData")
        '2 Dataset 
        '1 contains All tx except Completed
        'the other only contains Completed

        DS = New DataSet
        Dim mymessage As String = String.Empty
        sb.Clear()
        'Admin checking first
        'sb.Append("select * from pricechangehd ph where (ph.creator = '" & HelperClass1.UserId & "')")

        'myuser = HelperClass1.UserId.ToLower
        'myuser = "as\dlie"
        'myuser = "as\elai"
        'myuser = "as\rlo"
        'myuser = "AS\afok".ToLower
        'myuser = "as\weho"
        'myuser = "AS\shxu".ToLower
        'myuser = "as\jdai"
        'myuser = "AS\SCHAN".ToLower
        'myuser = "as\Dlam"
        'sb.Append("select distinct h.* ,o.officersebname::text as username,o2.officersebname::text as validatorname,o3.officersebname::text as cc1name ,o4.officersebname::text as cc2name,o5.officersebname::text as cc3name,o6.officersebname::text as cc4name" &
        '          " from doc.header h" &
        '          " left join officerseb o on lower(o.userid) = h.userid" &
        '          " left join officerseb o2 on lower(o2.userid) = h.validator" &
        '          " left join officerseb o3 on lower(o3.userid) = h.cc1" &
        '          " left join officerseb o4 on lower(o4.userid) = h.cc2" &
        '          " left join officerseb o5 on lower(o5.userid) = h.cc3" &
        '          " left join officerseb o6 on lower(o6.userid) = h.cc4  where h.id = " & headerid & ";") 'add join with ofsebid
        sb.Append("select distinct h.* ,o.username::text as username,o2.officersebname::text as validatorname,o3.officersebname::text as cc1name ,o4.officersebname::text as cc2name,o5.officersebname::text as cc3name,o6.officersebname::text as cc4name" &
                 " from doc.header h" &
                 " left join doc.user o on lower(o.userid) = h.userid" &
                 " left join officerseb o2 on lower(o2.userid) = h.validator" &
                 " left join officerseb o3 on lower(o3.userid) = h.cc1" &
                 " left join officerseb o4 on lower(o4.userid) = h.cc2" &
                 " left join officerseb o5 on lower(o5.userid) = h.cc3" &
                 " left join officerseb o6 on lower(o6.userid) = h.cc4  where h.id = " & headerid & ";") 'add join with ofsebid
        'sb.Append("select vd.id,vd.vendorcode,vd.documentid,vd.status,case vd.status when 1 then 'NEW'  when 2 then 'COMPLETED' else 'NOT-ASSIGNED' end as statusname,vd.headerid,v.vendorname::text,v.shortname::text,d.*,vr.version,gt.paymentcode,sc.leadtime,sc.sasl,q.nqsu,p.projectname,sa.auditby,sa.audittype,sa.auditgrade,sef.score,sif.myyear,sif.turnovery,sif.turnovery1,sif.turnovery2,sif.turnovery3,sif.turnovery4,sif.ratioy,sif.ratioy1,sif.ratioy2,sif.ratioy3,sif.ratioy4,null as filename,dt.doctypename,dl.levelname,de.expireddate from doc.vendordoc vd" &
        '          " left join doc.header h on h.id = vd.headerid" &
        '          " left join vendor v on v.vendorcode = vd.vendorcode" &
        '          " left join doc.document d on d.id = vd.documentid" &
        '          " left join doc.version vr on vr.documentid = d.id" &
        '          " left join doc.generalcontract gt on gt.documentid = d.id" &
        '          " left join doc.supplychain sc on sc.documentid = d.id" &
        '          " left join doc.qualityappendix q on q.documentid = d.id" &
        '          " left join doc.project p on p.documentid = d.id" &
        '          " left join doc.socialaudit sa on sa.documentid = d.id" &
        '          " left join doc.sef sef on sef.documentid = d.id" &
        '          " left join doc.sif sif on sif.documentid = d.id" &
        '          " left join doc.doctype dt on dt.id = d.doctypeid" &
        '          " left join doc.doclevel dl on dl.id = d.doclevelid" &
        '          " left join doc.docexpired de on de.documentid = d.id" &
        '          " where h.id = " & headerid & ";")
        sb.Append("select vd.id,vd.vendorcode,vd.documentid,vd.status,doc.getstatusdoc(vd.status) as statusname,vd.headerid,v.vendorname::text,v.shortname::text,d.*,vr.version,gt.paymentcode,sc.leadtime,sc.sasl,q.nqsu,p.projectname,sa.auditby,sa.audittype,sa.auditgrade,sa.overallauditresult,sef.score,sif.myyear,sif.turnovery,sif.turnovery1,sif.turnovery2,sif.turnovery3,sif.turnovery4,sif.ratioy,sif.ratioy1,sif.ratioy2,sif.ratioy3,sif.ratioy4,null as filename,dt.doctypename,dl.levelname,de.expireddate from doc.vendordoc vd" &
                 " left join doc.header h on h.id = vd.headerid" &
                 " left join vendor v on v.vendorcode = vd.vendorcode" &
                 " left join doc.document d on d.id = vd.documentid" &
                 " left join doc.version vr on vr.documentid = d.id" &
                 " left join doc.generalcontract gt on gt.documentid = d.id" &
                 " left join doc.supplychain sc on sc.documentid = d.id" &
                 " left join doc.qualityappendix q on q.documentid = d.id" &
                 " left join doc.project p on p.documentid = d.id" &
                 " left join doc.socialaudit sa on sa.documentid = d.id" &
                 " left join doc.sef sef on sef.documentid = d.id" &
                 " left join doc.sif sif on sif.documentid = d.id" &
                 " left join doc.doctype dt on dt.id = d.doctypeid" &
                 " left join doc.doclevel dl on dl.id = d.doclevelid" &
                 " left join doc.docexpired de on de.documentid = d.id" &
                 " where h.id = " & headerid & ";")
        If HelperClass1.UserInfo.IsAdmin Then
            sb.Append("select null::text as shortname union all (select distinct shortname::text from vendor  where not shortname isnull order by shortname);")
            sb.Append("select shortname::text,vendorcode from vendor where not shortname isnull order by shortname;")
            sb.Append("select null as vendorcode,''::text as description,''::text as vendorname union all (select vendorcode, vendorcode::text || ' - ' || vendorname::text as description,vendorname::text from vendor order by vendorname);")
        Else
            sb.Append(String.Format("select null::text as shortname union all (select distinct shortname::text from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid" &
                     " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where lower(u.userid) ~ '{0}$' and not shortname isnull order by shortname);", HelperClass1.UserInfo.userid.ToLower))
            sb.Append(String.Format("select distinct shortname::text,v.vendorcode as vendorcode from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid" &
                     " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where lower(u.userid) ~ '{0}$' and not shortname isnull order by shortname;", HelperClass1.UserInfo.userid.ToLower))
            sb.Append(String.Format("select null as vendorcode,''::text as description,''::text as vendorname,null::text as shortname union all (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text  from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid" &
                      " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where lower(u.userid) ~ '{0}$' and not shortname isnull  order by vendorname);", HelperClass1.UserInfo.userid.ToLower))

        End If
        sb.Append("select null as id,''::text as doctypename union all (select id,doctypename from doc.doctype order by doctypename);")
        sb.Append("select null as id,''::text as levelname union all (select id,levelname from doc.doclevel order by id);")
        sb.Append("select paymenttermid,payt || ' - ' || details as payt   from paymentterm  order by payt;")
        'sb.Append("select * from officerseb o  where lower(o.userid) = '" & myuser & "' limit 1;")
        'sb.Append("select ''::text as name,'' as userid,null as teamtitleid,'' as officersebname,''::text as teamtitleshortname union all (select distinct teamtitleshortname || ' - ' || officersebname as name,lower(userid) as userid,tt.teamtitleid,officersebname,tt.teamtitleshortname from officerseb o left join teamtitle tt on tt.teamtitleid = o.teamtitleid where teamtitleshortname in ('PD','SPM','PM') and isactive and userid <> 'as\lili2' order by tt.teamtitleid,officersebname);")
        sb.Append("select ''::text as name,'' as userid,null as teamtitleid,'' as officersebname,''::text as teamtitleshortname union all (select distinct teamtitleshortname || ' - ' || officersebname as name,lower(o.userid) as userid,tt.teamtitleid,officersebname,tt.teamtitleshortname from doc.user u left join officerseb o on o.userid = u.userid  left join teamtitle tt on tt.teamtitleid = o.teamtitleid where teamtitleshortname in ('PD','SPM','PM','PO') and o.isactive and o.userid <> 'as\lili2' order by tt.teamtitleid,officersebname);")
        sb.Append("select cvalue from doc.paramhd where paramname = 'docfolder'; ")
        'SIF and IdentitySheet


        sb.Append("select vd.vendorcode,sx.*,sl.name as labelname from  doc.vendordoc vd" &
                  " left join doc.header h on h.id = vd.headerid" &
                  " left join doc.document d on d.id = vd.documentid" &
                  " left join doc.sitx sx on sx.documentid = d.id" &
                  " left join doc.silabel sl on sl.id = sx.labelid" &
                  " where h.id = " & headerid & " and not sx.documentid isnull order by orderline;")

        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                DS.Tables(0).TableName = "Header"
                DS.Tables(1).TableName = "Detail"
                DS.Tables(2).TableName = "ShortName"
                DS.Tables(3).TableName = "ShortNameVendorCode"
                DS.Tables(4).TableName = "VendorName"
                DS.Tables(5).TableName = "DocType"
                DS.Tables(6).TableName = "DocLevel"
                DS.Tables(7).TableName = "PaymentTerm"
                DS.Tables(8).TableName = "User"
                DS.Tables(9).TableName = "Folder"
                DS.Tables(10).TableName = "SITx"

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
                            bsDetail = New BindingSource

                            bsShortname = New BindingSource
                            bsShortnameHelper = New BindingSource
                            bsShortnameVendor = New BindingSource
                            bsVendorname = New BindingSource
                            bsVendornameHelper = New BindingSource
                            bsDocType = New BindingSource
                            bsDocTypeHelper = New BindingSource
                            bsDocLevel = New BindingSource
                            bsPaymentTerm = New BindingSource
                            bsfolder = New BindingSource
                            bsSiTx = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("id")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns(0).AutoIncrement = True
                            DS.Tables(0).Columns(0).AutoIncrementSeed = 0
                            DS.Tables(0).Columns(0).AutoIncrementStep = -1
                            DS.Tables(0).TableName = "Header"


                            Dim pk1(0) As DataColumn
                            pk1(0) = DS.Tables(1).Columns("id")
                            DS.Tables(1).PrimaryKey = pk1
                            DS.Tables(1).Columns(0).AutoIncrement = True
                            DS.Tables(1).Columns(0).AutoIncrementSeed = 0
                            DS.Tables(1).Columns(0).AutoIncrementStep = -1
                            DS.Tables(1).TableName = "Details"

                            Dim rel As DataRelation
                            Dim hcol As DataColumn
                            Dim dcol As DataColumn
                            'create relation ds.table(0) and ds.table(1)
                            hcol = DS.Tables(0).Columns("id") 'id in table header
                            dcol = DS.Tables(1).Columns("headerid") 'headerid in table vendordoc
                            rel = New DataRelation("hdrel", hcol, dcol)
                            

                            DS.Relations.Add(rel)

                            bsheader.DataSource = DS.Tables(0)
                            bsDetail.DataSource = DS.Tables(1)

                            bsShortname.DataSource = New DataView(DS.Tables(2))
                            bsShortnameHelper.DataSource = New DataView(DS.Tables(2))
                            bsShortnameVendor.DataSource = DS.Tables(3)
                            bsVendorname.DataSource = New DataView(DS.Tables(4))
                            bsVendornameHelper.DataSource = New DataView(DS.Tables(4))


                            bsDocType.DataSource = New DataView(DS.Tables(5))
                            bsDocTypeHelper.DataSource = New DataView(DS.Tables(5))

                            bsDocLevel.DataSource = DS.Tables(6)
                            bsPaymentTerm.DataSource = DS.Tables(7)

                            bsfolder.DataSource = DS.Tables(9)

                            Dim DocCol(1) As DataColumn
                            Dim SITxCol(1) As DataColumn
                            DocCol(0) = DS.Tables(1).Columns("vendorcode")
                            DocCol(1) = DS.Tables(1).Columns("documentid")
                            SITxCol(0) = DS.Tables(10).Columns("vendorcode")
                            SITxCol(1) = DS.Tables(10).Columns("documentid")
                            rel = New DataRelation("SIRel", DocCol, SITxCol)
                            Dim FkeyConstraint As ForeignKeyConstraint
                            FkeyConstraint = New ForeignKeyConstraint("fkey", DocCol, SITxCol)
                            FkeyConstraint.DeleteRule = Rule.Cascade
                            FkeyConstraint.UpdateRule = Rule.Cascade

                            DS.Relations.Add(rel)

                            bsSiTx.DataSource = bsDetail
                            bsSiTx.DataMember = "SIRel"

                            TextBox1.DataBindings.Clear()
                            TextBox2.DataBindings.Clear()
                            TextBox3.DataBindings.Clear()
                            TextBox4.DataBindings.Clear()
                            TextBox5.DataBindings.Clear()
                            TextBox6.DataBindings.Clear()
                            TextBox7.DataBindings.Clear()

                            If headerid = 0 Then ' New Record
                                'bsheader.AddNew()
                                Dim drv As DataRowView = bsheader.AddNew()
                                drv.Row.Item("creationdate") = Date.Today
                                drv.Row.Item("userid") = HelperClass1.UserId.ToLower
                                drv.Row.Item("username") = HelperClass1.UserInfo.DisplayName
                                bsheader.EndEdit()
                            Else
                                ' recPosition = bsDetail.Find("vdid", VDID)

                            End If


                            ComboBox1.DataBindings.Clear()

                            ComboBox1.DisplayMember = "shortname"
                            ComboBox1.ValueMember = "shortname"
                            'ComboBox1.SelectedIndex = -1
                            ComboBox1.DataSource = bsShortname
                            'ComboBox1.DataBindings.Add("SelectedValue", bsDetail, "shortname")

                            ComboBox2.DataBindings.Clear()

                            ComboBox2.DisplayMember = "description"
                            ComboBox2.ValueMember = "vendorcode"
                            ComboBox2.DataSource = bsVendorname
                            ComboBox2.DataBindings.Add("Selectedvalue", bsDetail, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged)


                            ComboBox3.DataBindings.Clear()
                            ComboBox3.DataSource = bsDocType
                            ComboBox3.DisplayMember = "doctypename"
                            ComboBox3.ValueMember = "id"
                            ComboBox3.DataBindings.Add("SelectedValue", bsDetail, "doctypeid", True, DataSourceUpdateMode.OnPropertyChanged)

                            ComboBox4.DataBindings.Clear()
                            ComboBox4.DataSource = bsDocLevel
                            ComboBox4.DisplayMember = "levelname"
                            ComboBox4.ValueMember = "id"
                            ComboBox4.DataBindings.Add("SelectedValue", bsDetail, "doclevelid", True, DataSourceUpdateMode.OnPropertyChanged)


                            ComboBox5.DataBindings.Clear()
                            ComboBox5.DataSource = bsPaymentTerm
                            ComboBox5.DisplayMember = "payt"
                            ComboBox5.ValueMember = "paymenttermid"
                            ComboBox5.DataBindings.Add("SelectedValue", bsDetail, "paymentcode", True, DataSourceUpdateMode.OnPropertyChanged)


                            DateTimePicker1.DataBindings.Clear()
                            DateTimePicker2.DataBindings.Clear()
                            DateTimePicker3.DataBindings.Clear()
                            TextBox8.DataBindings.Clear()
                            TextBox9.DataBindings.Clear()
                            TextBox10.DataBindings.Clear()
                            TextBox11.DataBindings.Clear()
                            TextBox12.DataBindings.Clear()
                            TextBox13.DataBindings.Clear()
                            TextBox14.DataBindings.Clear()
                            TextBox15.DataBindings.Clear()
                            'TextBox16.DataBindings.Clear()
                            'TextBox17.DataBindings.Clear()
                            TextBox18.DataBindings.Clear()
                            TextBox19.DataBindings.Clear()
                            TextBox20.DataBindings.Clear()
                            TextBox21.DataBindings.Clear()
                            TextBox22.DataBindings.Clear()
                            TextBox23.DataBindings.Clear()
                            TextBox24.DataBindings.Clear()


                            ComboBox6.DataBindings.Clear()
                            ComboBox6.DataSource = auditbylist
                            ComboBox6.DataBindings.Add("Text", bsDetail, "auditby", True, DataSourceUpdateMode.OnPropertyChanged, "")
                            'ComboBox6.SelectedIndex = -1

                            ComboBox7.DataBindings.Clear()
                            ComboBox7.DataSource = auditTypelist
                            ComboBox7.DataBindings.Add("Text", bsDetail, "audittype", True, DataSourceUpdateMode.OnPropertyChanged, "")

                            ComboBox8.DataBindings.Clear()
                            ComboBox8.DataSource = auditGradelist
                            ComboBox8.DataBindings.Add("Text", bsDetail, "auditgrade", True, DataSourceUpdateMode.OnPropertyChanged, "")

                            TextBox1.DataBindings.Add(New Binding("Text", bsheader, "username", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox2.DataBindings.Add(New Binding("Text", bsheader, "validatorname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox3.DataBindings.Add(New Binding("Text", bsheader, "cc1name", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox4.DataBindings.Add(New Binding("Text", bsheader, "cc2name", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox5.DataBindings.Add(New Binding("Text", bsheader, "cc3name", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox6.DataBindings.Add(New Binding("Text", bsheader, "cc4name", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox7.DataBindings.Add(New Binding("Text", bsheader, "otheremail", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            DateTimePicker1.DataBindings.Add(New Binding("Text", bsheader, "creationdate"))

                            TextBox8.DataBindings.Add(New Binding("Text", bsDetail, "remarks", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox9.DataBindings.Add(New Binding("Text", bsDetail, "version", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox10.DataBindings.Add(New Binding("Text", bsDetail, "filename", True, DataSourceUpdateMode.OnPropertyChanged, "")) 'if no value means update
                            TextBox11.DataBindings.Add(New Binding("Text", bsDetail, "leadtime", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox12.DataBindings.Add(New Binding("Text", bsDetail, "sasl", True, DataSourceUpdateMode.OnPropertyChanged, "", "##0"))
                            TextBox13.DataBindings.Add(New Binding("Text", bsDetail, "nqsu", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
                            TextBox14.DataBindings.Add(New Binding("Text", bsDetail, "projectname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            'TextBox15.DataBindings.Add(New Binding("Text", bsDetail, "auditbyvalue", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox16.DataBindings.Add(New Binding("Text", bsDetail, "audittype", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox17.DataBindings.Add(New Binding("Text", bsDetail, "auditgrade", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox18.DataBindings.Add(New Binding("Text", bsDetail, "score", True, DataSourceUpdateMode.OnPropertyChanged, "", "##0"))
                            DateTimePicker2.DataBindings.Add(New Binding("Text", bsDetail, "docdate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            DateTimePicker3.DataBindings.Add(New Binding("Text", bsDetail, "expireddate", True, DataSourceUpdateMode.OnPropertyChanged, ""))

                            Dim b19 As Binding = New Binding("Text", bsDetail, "turnovery", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b20 As Binding = New Binding("Text", bsDetail, "turnovery1", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b21 As Binding = New Binding("Text", bsDetail, "turnovery2", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b22 As Binding = New Binding("Text", bsDetail, "turnovery3", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b23 As Binding = New Binding("Text", bsDetail, "turnovery4", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00")
                            Dim b24 As Binding = New Binding("Text", bsDetail, "ratioy", True, DataSourceUpdateMode.OnPropertyChanged, "", "##0.00")
                            Dim b25 As Binding = New Binding("Text", bsDetail, "ratioy1", True, DataSourceUpdateMode.OnPropertyChanged, "", "##0.00")
                            Dim b26 As Binding = New Binding("Text", bsDetail, "ratioy2", True, DataSourceUpdateMode.OnPropertyChanged, "", "##0.00")
                            Dim b27 As Binding = New Binding("Text", bsDetail, "ratioy3", True, DataSourceUpdateMode.OnPropertyChanged, "", "##0.00")
                            Dim b28 As Binding = New Binding("Text", bsDetail, "ratioy4", True, DataSourceUpdateMode.OnPropertyChanged, "", "##0.00")
                            TextBox19.DataBindings.Add(b19)
                            TextBox20.DataBindings.Add(b20)
                            TextBox21.DataBindings.Add(b21)
                            TextBox22.DataBindings.Add(b22)
                            TextBox23.DataBindings.Add(b23)
                            TextBox24.DataBindings.Add(b24)
                            TextBox25.DataBindings.Add(b25)
                            TextBox26.DataBindings.Add(b26)
                            TextBox27.DataBindings.Add(b27)
                            TextBox28.DataBindings.Add(b28)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = bsDetail
                            DataGridView1.RowTemplate.Height = 22
                            BindingNavigator1.BindingSource = bsDetail

                            DataGridView2.AutoGenerateColumns = False
                            DataGridView2.DataSource = bsSiTx

                            Select Case isRole
                                Case Role.Admin

                                Case Role.Creator
                                    ToolStripButton5.Visible = False
                                Case Role.Other
                                    ToolStripButton1.Visible = False
                                    ToolStripButton4.Visible = False
                                    ToolStripButton2.Visible = False
                                    ToolStripButton5.Visible = False
                                    ToolStripButton6.Visible = False
                                    DataGridView1.ContextMenuStrip = Nothing
                                Case Role.Validator
                                    ToolStripButton1.Visible = False
                                    ToolStripButton4.Visible = False
                                    DataGridView1.ContextMenuStrip = Nothing
                            End Select
                            '
                            TextBox15.Enabled = False
                            If ComboBox6.SelectedValue = "Intertek" Then
                                'TextBox15.Enabled = True
                            End If
                            'Renew
                            If Not IsNothing(RenewDRV) Then
                                ToolStripButton1.Visible = False
                                ToolStripButton4.Visible = False
                                ToolStripButton5.Visible = False
                                DataGridView1.ContextMenuStrip = Nothing

                                If headerid = 0 Then 'Renew
                                    Dim mydrv = bsDetail.AddNew()

                                    'bsDetail.Item("doctypeid") = RenewDRV.Item("doctypeid")

                                    mydrv.Row.Item("doctypename") = RenewDRV.Item("doctypename")
                                    mydrv.Row.Item("doctypeid") = RenewDRV.Item("doctypeid")


                                    mydrv.Row.Item("headerid") = 0
                                    mydrv.Row.Item("userid") = bsheader.Current.row.item("userid")
                                    mydrv.Row.Item("uploaddate") = Date.Today
                                    mydrv.Row.Item("docdate") = Date.Today
                                    mydrv.Row.Item("levelname") = "Supplier"
                                    mydrv.Row.Item("doclevelid") = 1
                                    mydrv.row.item("vendorcode") = RenewDRV.Item("vendorcode")
                                    mydrv.row.item("shortname") = RenewDRV.Item("shortname")
                                    mydrv.row.item("vendorname") = RenewDRV.Item("suppliername")
                                    'vendorcode
                                    'shortname

                                    'Sync with bsDocType
                                    Dim itemfound As Integer = bsDocType.Find("id", RenewDRV.Item("doctypeid"))
                                    bsDocType.Position = itemfound
                                    itemfound = bsDocLevel.Find("levelname", mydrv.row.item("levelname"))
                                    bsDocLevel.Position = itemfound
                                    itemfound = bsVendorname.Find("vendorcode", mydrv.row.item("vendorcode"))
                                    bsVendorname.Position = itemfound
                                    'bsDocType.Position = bsDocTypeHelper.Position
                                    'enabledTextBox()
                                Else 'Non Renew


                                    Dim itemfound As Integer = bsDetail.Find("id", RenewDRV.Item("vdid"))

                                    bsDetail.Position = itemfound

                                    bsDetail.Current.item("status") = 4
                                    bsDetail.Current.Item("statusname") = "Non-Renewable"
                                    bsDetail.EndEdit()
                                    TextBox1.Enabled = False
                                    TextBox2.Enabled = False
                                    TextBox3.Enabled = False
                                    TextBox4.Enabled = False
                                    TextBox5.Enabled = False
                                    TextBox6.Enabled = False
                                    TextBox7.Enabled = False
                                    TextBox9.Enabled = False
                                    TextBox10.Enabled = False
                                    ComboBox2.Enabled = False
                                    ComboBox3.Enabled = False
                                    ComboBox4.Enabled = False
                                    ComboBox5.Enabled = False
                                    Button1.Enabled = False
                                    Button2.Enabled = False
                                    Button3.Enabled = False
                                    Button4.Enabled = False
                                    Button5.Enabled = False
                                    Button6.Enabled = False
                                    Button7.Enabled = False
                                    Button8.Enabled = False
                                    Button9.Enabled = False
                                    Button10.Enabled = False
                                    Button5.Enabled = False
                                    DateTimePicker1.Enabled = False
                                    DateTimePicker2.Enabled = False
                                    DateTimePicker3.Enabled = False
                                    DataGridView1.Invalidate()

                                End If
                            Else
                                recPosition = bsDetail.Find("id", VDID)
                                If recPosition >= 0 Then
                                    bsDetail.Position = recPosition
                                End If

                            End If

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


    Private Sub bsheader_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles bsheader.ListChanged

        TextBox1.Enabled = Not IsNothing(bsheader.Current)
        TextBox2.Enabled = Not IsNothing(bsheader.Current)
        TextBox3.Enabled = Not IsNothing(bsheader.Current)
        TextBox4.Enabled = Not IsNothing(bsheader.Current)
        TextBox5.Enabled = Not IsNothing(bsheader.Current)
        TextBox6.Enabled = Not IsNothing(bsheader.Current)
        TextBox7.Enabled = Not IsNothing(bsheader.Current)
    End Sub

    Private Sub bsDetail_CurrentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles bsDetail.CurrentChanged
        bsDetail.EndEdit()
    End Sub



    Private Sub bsDetail_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles bsDetail.ListChanged
        TextBox8.Enabled = Not IsNothing(bsDetail.Current)
        TextBox9.Enabled = Not IsNothing(bsDetail.Current)
        TextBox10.Enabled = Not IsNothing(bsDetail.Current)
        TextBox10.ReadOnly = True
        TextBox11.Enabled = Not IsNothing(bsDetail.Current)
        TextBox12.Enabled = Not IsNothing(bsDetail.Current)
        TextBox13.Enabled = Not IsNothing(bsDetail.Current)
        TextBox14.Enabled = Not IsNothing(bsDetail.Current)
        TextBox15.Enabled = Not IsNothing(bsDetail.Current)
        TextBox16.Enabled = Not IsNothing(bsDetail.Current)
        TextBox17.Enabled = Not IsNothing(bsDetail.Current)
        TextBox18.Enabled = Not IsNothing(bsDetail.Current)
        TextBox19.Enabled = Not IsNothing(bsDetail.Current)
        TextBox20.Enabled = Not IsNothing(bsDetail.Current)
        TextBox21.Enabled = Not IsNothing(bsDetail.Current)
        TextBox22.Enabled = Not IsNothing(bsDetail.Current)
        TextBox23.Enabled = Not IsNothing(bsDetail.Current)
        TextBox24.Enabled = Not IsNothing(bsDetail.Current)
        RadioButton1.Enabled = Not IsNothing(bsDetail.Current)
        RadioButton2.Enabled = Not IsNothing(bsDetail.Current)

        If Not IsNothing(bsDetail.Current) Then
            ComboBox1.Enabled = RadioButton1.Checked
            ComboBox2.Enabled = RadioButton2.Checked
        Else
            ComboBox1.Enabled = Not IsNothing(bsDetail.Current)
            ComboBox2.Enabled = Not IsNothing(bsDetail.Current)

        End If
        ComboBox3.Enabled = Not IsNothing(bsDetail.Current)
        ComboBox4.Enabled = Not IsNothing(bsDetail.Current)
        ComboBox5.Enabled = Not IsNothing(bsDetail.Current)
        DateTimePicker2.Enabled = Not IsNothing(bsDetail.Current)
        DateTimePicker3.Enabled = Not IsNothing(bsDetail.Current)
        Button6.Enabled = Not IsNothing(bsDetail.Current)
        Button8.Enabled = Not IsNothing(bsDetail.Current)
        Button9.Enabled = Not IsNothing(bsDetail.Current)
        Button10.Enabled = Not IsNothing(bsDetail.Current)
    End Sub


    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Try
            bsheader.EndEdit()
            Dim mydrv As DataRowView = bsDetail.AddNew()

            Dim myhddrv As DataRowView = bsheader.Current()
            mydrv.Row.Item("headerid") = myhddrv.Row.Item("id")
            mydrv.Row.Item("userid") = bsheader.Current.row.item("userid")
            mydrv.Row.Item("uploaddate") = Date.Today
            mydrv.Row.Item("docdate") = Date.Today
            mydrv.Row.Item("levelname") = "Supplier"
            mydrv.Row.Item("doclevelid") = 1

            mydrv.Row.Item("documentid") = mydrv.Row.Item("id")
            'bsDetail.EndEdit()
            'ComboBox1.SelectedIndex = -1
            'ComboBox2.SelectedIndex = -1
            'ComboBox4.SelectedIndex = 1
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        ComboBox1.Enabled = RadioButton1.Checked
        ComboBox2.Enabled = RadioButton2.Checked
        Button7.Enabled = RadioButton1.Checked
        Button8.Enabled = RadioButton2.Checked

        'If RadioButton2.Checked Then
        '    ComboBox1.SelectedIndex = -1
        'Else
        '    ComboBox2.SelectedIndex = -1
        'End If
    End Sub


    Private Sub ComboBox3_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        'General Contract
        enabledTextBox()


    End Sub

    Private Sub enabledTextBox()
        GroupBox1.Enabled = False
        GroupBox2.Enabled = False
        GroupBox3.Enabled = False
        GroupBox5.Enabled = False
        GroupBox6.Enabled = False
        GroupBox7.Enabled = False
        GroupBox9.Enabled = False
        Button12.Enabled = False

        If Not IsNothing(ComboBox3.SelectedValue) Then
            Dim drv As DataRowView = ComboBox3.SelectedItem

            TabControl1.SelectedTab = TabPage1
            Try
                Dim myid As DataRowView = bsDetail.Current
                If Not IsNothing(bsSiTx.Current) Then
                    If Not (myid.Row.Item("doctypeid") = SIF Or myid.Row.Item("doctypeid") = IDENTITY_SHEET) Then
                        bsSiTx.Filter = String.Format("documentid = {0}", myid.Row.Item("documentid"))
                        For Each mydrv As DataRowView In bsSiTx.List
                            mydrv.Delete()
                        Next
                        bsSiTx.Filter = ""
                    End If
                    
                End If
            Catch ex As Exception

            End Try
            

            Select Case drv.Row.Item("doctypename") 'ComboBox3.SelectedText
                Case "Contract: General Contract"
                    GroupBox1.Enabled = True

                    'General Contract
                    'ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""

                    ComboBox6.SelectedIndex = -1
                    ComboBox7.SelectedIndex = -1
                    ComboBox8.SelectedIndex = -1


                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""

                    '

                Case "Contract: Supply Chain Appendix"
                    GroupBox2.Enabled = True
                    'General Contract
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    'TextBox11.Text = ""
                    'TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""
                    ComboBox6.SelectedIndex = -1
                    ComboBox7.SelectedIndex = -1
                    ComboBox8.SelectedIndex = -1

                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""
                Case "Contract: Quality Appendix"
                    GroupBox3.Enabled = True
                    'General Contract
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    'TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""
                    ComboBox6.SelectedIndex = -1
                    ComboBox7.SelectedIndex = -1
                    ComboBox8.SelectedIndex = -1

                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""
                Case "Social Audit"
                    GroupBox5.Enabled = True
                    'General Contract
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    'TextBox15.Text = ""
                    'TextBox16.Text = ""
                    'TextBox17.Text = ""

                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""
                Case "SEF"
                    GroupBox6.Enabled = True
                    'General Contract
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""
                    ComboBox6.SelectedIndex = -1
                    ComboBox7.SelectedIndex = -1
                    ComboBox8.SelectedIndex = -1

                    'SEF
                    'TextBox18.Text = ""

                    'SIF
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""
                Case "SIF"
                    TabControl1.SelectedTab = TabPage2
                    Button12.Enabled = True
                    GroupBox7.Enabled = True
                    GroupBox9.Enabled = True
                    'General Contract
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""
                    ComboBox6.SelectedIndex = -1
                    ComboBox7.SelectedIndex = -1
                    ComboBox8.SelectedIndex = -1

                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    'TextBox19.Text = ""
                    'TextBox20.Text = ""
                    'TextBox21.Text = ""
                    'TextBox22.Text = ""
                    'TextBox23.Text = ""
                    'TextBox24.Text = ""
                    'TextBox25.Text = ""
                    'TextBox26.Text = ""
                    'TextBox27.Text = ""
                    'TextBox28.Text = ""
                Case "Identity sheet"
                    TabControl1.SelectedTab = TabPage2
                    GroupBox9.Enabled = True
                    Button12.Enabled = True

                    'General Contract
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""
                    ComboBox6.SelectedIndex = -1
                    ComboBox7.SelectedIndex = -1
                    ComboBox8.SelectedIndex = -1

                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    'TextBox19.Text = ""
                    'TextBox20.Text = ""
                    'TextBox21.Text = ""
                    'TextBox22.Text = ""
                    'TextBox23.Text = ""
                    'TextBox24.Text = ""
                    'TextBox25.Text = ""
                    'TextBox26.Text = ""
                    'TextBox27.Text = ""
                    'TextBox28.Text = ""
                Case Else
                    ComboBox5.SelectedIndex = -1


                    'supply chain Appendix
                    TextBox11.Text = ""
                    TextBox12.Text = ""

                    'Quality Appendix
                    TextBox13.Text = ""

                    'Social Audit
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""
                    ComboBox6.SelectedIndex = -1
                    ComboBox7.SelectedIndex = -1
                    ComboBox8.SelectedIndex = -1

                    'SEF
                    TextBox18.Text = ""

                    'SIF
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""
            End Select
        End If
    End Sub
    Private Sub ComboBox3_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectionChangeCommitted
        'Call checkcombobox3()
        ' enabledTextBox()
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)

        '1. Force the Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        If Not IsNothing(ComboBox3.SelectedValue) Then
            Dim drv As DataRowView = ComboBox3.SelectedItem
            'If drv.Row.Item("doctypename") <> "" Then
            If Not IsNothing(bsDetail.Current) Then
                Dim mydrv As DataRowView = bsDetail.Current
                mydrv.Row.BeginEdit()
                'If IsDBNull(mydrv.Row.Item("doctypename")) Then
                '    'mydrv.Row.Item("doctypeid") = drv.Row.Item("id")
                '    mydrv.Row.Item("doctypename") = drv.Row.Item("doctypename")
                'Else
                '    If Not mydrv.Row.Item("doctypename") = drv.Row.Item("doctypename") Then
                '        'mydrv.Row.Item("doctypeid") = drv.Row.Item("id")
                '        mydrv.Row.Item("doctypename") = drv.Row.Item("doctypename")
                '    End If
                'End If

                mydrv.Row.Item("doctypename") = drv.Row.Item("doctypename")
                bsDetail.EndEdit()
                DataGridView1.Invalidate()
            End If
        End If

    End Sub
    Private Sub ComboBox4_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        GroupBox4.Enabled = False
        If Not IsNothing(ComboBox4.SelectedValue) Then
            Dim drv As DataRowView = ComboBox4.SelectedItem
            Dim mydrv As DataRowView = bsDetail.Current

            Select Case drv.Row.Item("levelname")
                Case "Product / Project"
                    GroupBox4.Enabled = True
                Case Else
                    'mydrv.Row.BeginEdit()

                    TextBox14.Text = ""
            End Select
            'mydrv.Row.Item("levelname") = drv.Item(1)
        End If
    End Sub
    Private Sub ComboBox4_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectionChangeCommitted
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)

        '1. Force the Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        If Not IsNothing(ComboBox4.SelectedValue) Then
            Dim drv As DataRowView = ComboBox4.SelectedItem
            'If drv.Row.Item("levelname") <> "" Then
            If Not IsNothing(bsDetail.Current) Then
                Dim mydrv As DataRowView = bsDetail.Current
                mydrv.Row.BeginEdit()
                If IsDBNull(mydrv.Row.Item("levelname")) Then
                    mydrv.Row.Item("levelname") = drv.Row.Item("levelname")
                Else
                    If Not mydrv.Row.Item("levelname") = drv.Row.Item("levelname") Then
                        mydrv.Row.Item("levelname") = drv.Row.Item("levelname")
                    End If
                End If
                If mydrv.Row.Item("levelname") <> "Product and Project" Then
                    'mydrv.Row.Item("projectname") = ""
                    TextBox14.Text = ""
                End If
                bsDetail.EndEdit()
                DataGridView1.Invalidate()
            End If
            'End If
            'GroupBox4.Enabled = False
            Select Case drv.Row.Item("levelname")
                Case "Product / Project"
                    GroupBox4.Enabled = True
                Case Else
                    TextBox14.Text = ""
            End Select


        End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

        'If Not IsNothing(bsDetail.Current) Then
        '    Dim drvd As DataRowView = bsDetail.Current
        '    If Not drvd.Row.RowState = DataRowState.Unchanged Then
        '        ComboBox1.SelectedIndex = -1
        '        If Not IsNothing(bsDetail.Current) Then
        '            Dim mydrv As DataRowView = bsDetail.Current
        '            If IsDBNull(mydrv.Row.Item("vendorname")) Then
        '                mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
        '                mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")

        '            Else
        '                If Not mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname") Then
        '                    mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
        '                    mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
        '                    mydrv.Row.Item("shortname") = ""
        '                End If
        '            End If
        '            DataGridView1.Invalidate()
        '        End If
        '    End If
        'End If

    End Sub
    Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectionChangeCommitted
        'checkcombobox2()

        Dim myobj As ComboBox = DirectCast(sender, ComboBox)

        '1. Force the Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        If Not IsNothing(ComboBox2.SelectedItem) Then
            Dim drv As DataRowView = ComboBox2.SelectedItem
            'If drv.Row.Item("description") <> "" Then
            ComboBox1.SelectedIndex = -1 'Set Combobox for Short Name is blank
            If Not IsNothing(bsDetail.Current) Then
                Dim mydrv As DataRowView = bsDetail.Current
                mydrv.Row.BeginEdit()

                'If IsDBNull(mydrv.Row.Item("vendorname")) Then
                '    mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
                '    mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")


                'Else
                '    If Not mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname") Then
                '        mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
                '        mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")

                '        'mydrv.Row.Item("shortname") = ""
                '        mydrv.Row.Item("shortname") = DBNull.Value
                '    End If
                'End If

                mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
                mydrv.Row.Item("shortname") = DBNull.Value

                bsDetail.EndEdit()
                DataGridView1.Invalidate()
            End If
            'End If
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        'checkcombobox1()
        'Dim myobj As ComboBox = DirectCast(sender, ComboBox)

        ''1. Force the Combobox to commit the value 
        'For Each binding As Binding In myobj.DataBindings
        '    binding.WriteValue()
        '    binding.ReadValue()
        'Next

        'No binding for ComboBox1

        If Not IsNothing(ComboBox1.SelectedItem) Then
            Dim drv As DataRowView = ComboBox1.SelectedItem
            'If drv.Row.Item("shortname") <> "" Then
            ComboBox2.SelectedIndex = -1
            Dim mydrv As DataRowView = bsDetail.Current
            mydrv.Row.BeginEdit()
            mydrv.Row.Item("vendorcode") = DBNull.Value
            mydrv.Row.Item("vendorname") = DBNull.Value
            mydrv.Row.Item("shortname") = drv.Row.Item("shortname")
            'End If
            bsDetail.EndEdit()
            DataGridView1.Invalidate()
        End If

    End Sub


    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If Not IsNothing(bsDetail.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                'Dim mydrv As DataRowView = bsDetail.Current
                'mydrv.Row.CancelEdit()
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    bsDetail.RemoveAt(drv.Index)
                Next
                'bsDetail.RemoveCurrent()
            End If
        End If

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim openfiledialog1 As New OpenFileDialog
        If openfiledialog1.ShowDialog = DialogResult.OK Then
            Dim mydrv As DataRowView = bsDetail.Current
            mydrv.Row.Item("docname") = IO.Path.GetFileName(openfiledialog1.FileName)
            mydrv.Row.Item("docext") = IO.Path.GetExtension(openfiledialog1.FileName)
            'mydrv.Row.Item("userid") = "as\dlie"
            'mydrv.Row.Item("uploaddate") = Date.Today.AddDays(-1)
            TextBox10.Text = openfiledialog1.FileName
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click
        Dim myobj = CType(sender, Button)
        Dim bs As New BindingSource
        bs.DataSource = DS.Tables(8)
        Dim myform As New FormGetValidator(bs)
        bs.Filter = ""
        If myobj.Name = "Button1" Or myobj.Name = "Button2" Then
            bs.Filter = "teamtitleshortname <> 'PO'"
        End If
        If myform.ShowDialog = DialogResult.OK Then
            Dim drv As DataRowView = bs.Current
            Dim myrowhd As DataRowView = bsheader.Current

            Select Case myobj.Name
                Case "Button1"
                    TextBox2.Text = drv.Row.Item("name")
                    validatorid = drv.Row.Item("userid")
                    myrowhd.Row.Item("validator") = IIf(validatorid.ToLower = "", DBNull.Value, validatorid.ToLower)
                    myrowhd.Row.Item("validatorname") = drv.Row.Item("name")
                Case "Button2"
                    TextBox3.Text = drv.Row.Item("name")
                    cc1id = drv.Row.Item("userid")
                    myrowhd.Row.Item("cc1") = IIf(cc1id.ToLower = "", DBNull.Value, cc1id.ToLower)
                    myrowhd.Row.Item("cc1name") = drv.Row.Item("name")
                Case "Button3"
                    TextBox4.Text = drv.Row.Item("name")
                    cc2id = drv.Row.Item("userid")
                    myrowhd.Row.Item("cc2") = IIf(cc2id.ToLower = "", DBNull.Value, cc2id.ToLower)
                    myrowhd.Row.Item("cc2name") = drv.Row.Item("name")
                Case "Button4"
                    TextBox5.Text = drv.Row.Item("name")
                    cc3id = drv.Row.Item("userid")
                    myrowhd.Row.Item("cc3") = IIf(cc3id.ToLower = "", DBNull.Value, cc3id.ToLower)
                    myrowhd.Row.Item("cc3name") = drv.Row.Item("name")
                Case "Button5"
                    TextBox6.Text = drv.Row.Item("name")
                    cc4id = drv.Row.Item("userid")
                    myrowhd.Row.Item("cc4") = IIf(cc4id.ToLower = "", DBNull.Value, cc4id.ToLower)
                    myrowhd.Row.Item("cc4name") = drv.Row.Item("name")
            End Select
        End If
    End Sub



    Private Sub Button8_Click1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Handles Button7.Click, Button8.Click, Button9.Click, Button10.Click, Button11.Click
        Dim myobj As Button = CType(sender, Button)
        If Not IsNothing(bsDetail.Current) Then
            Dim drv As DataRowView = bsDetail.Current
            drv.Row.BeginEdit()
            Select Case myobj.Name
                Case "Button7"
                    ComboBox1.SelectedIndex = -1
                    drv.Row.Item("shortname") = DBNull.Value
                Case "Button8"
                    ComboBox2.SelectedIndex = -1
                    'drv.Row.Item("vendorcode") = DBNull.Value
                    drv.Row.Item("vendorname") = DBNull.Value
                Case "Button9"
                    ComboBox3.SelectedIndex = -1
                    drv.Row.Item("doctypename") = DBNull.Value
                    'drv.Row.Item("doctypeid") = DBNull.Value
                    ComboBox5.SelectedIndex = -1
                    'drv.Row.Item("paymentcode") = DBNull.Value
                    TextBox11.Text = ""
                    TextBox12.Text = ""
                    TextBox13.Text = ""

                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""
                    TextBox18.Text = ""
                    TextBox19.Text = ""
                    TextBox20.Text = ""
                    TextBox21.Text = ""
                    TextBox22.Text = ""
                    TextBox23.Text = ""
                    TextBox24.Text = ""
                    TextBox25.Text = ""
                    TextBox26.Text = ""
                    TextBox27.Text = ""
                    TextBox28.Text = ""



                Case "Button10"
                    ComboBox4.SelectedIndex = -1
                    drv.Row.Item("levelname") = DBNull.Value
                    TextBox14.Text = ""


                Case "Button11"
                    ComboBox5.SelectedIndex = -1
                    'drv.Row.Item("paymentcode") = DBNull.Value

            End Select
            DataGridView1.Invalidate()
        End If

    End Sub

    Private Sub TextBox14_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        DataGridView1.Invalidate()
    End Sub

    Private Sub onTextBoxBindingParse(ByVal sender As Object, ByVal e As ConvertEventArgs)
        If (IsDBNull(e.Value)) Then
            e.Value = String.Empty
        ElseIf (e.Value = "") Then
            e.Value = DBNull.Value
        End If
    End Sub




    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click, Button8.Click, Button9.Click
        Dim myobj As Button = CType(sender, Button)
        'Dim myform As Object
        Try


            Select Case myobj.Name
                Case "Button7"
                    Dim myform = New FormHelper(bsShortnameHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "shortname"
                    If Not myform.ShowDialog = DialogResult.OK Then
                        ComboBox1.SelectedItem = -1
                    Else
                        'checkcombobox1()
                        Dim drv As DataRowView = bsShortnameHelper.Current

                        'If drv.Row.Item("shortname") <> "" Then
                        ComboBox2.SelectedIndex = -1
                        Dim mydrv As DataRowView = bsDetail.Current
                        'mydrv.Row.BeginEdit()
                        'mydrv.Row.Item("vendorcode") = DBNull.Value
                        mydrv.Row.Item("vendorname") = DBNull.Value
                        mydrv.Row.Item("shortname") = drv.Row.Item("shortname")
                        'Sync with Combobox BindingSource
                        Dim itemfound As Integer = bsShortname.Find("shortname", drv.Row.Item("shortname"))
                        If itemfound <= 0 Then
                            MessageBox.Show("Please check your Vendor Code!")
                        End If
                        bsShortname.Position = itemfound
                    End If

                Case "Button8"
                    Dim myform = New FormHelper(bsVendornameHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "description"
                    If Not myform.ShowDialog = DialogResult.OK Then
                        ' ComboBox2.SelectedIndex = -1
                    Else
                        'checkcombobox2()
                        Dim drv As DataRowView = bsVendornameHelper.Current
                        Dim mydrv As DataRowView = bsDetail.Current

                        mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
                        mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
                        mydrv.Row.EndEdit()
                        'Sync with Combobox BindingSource
                        Dim itemfound As Integer = bsVendorname.Find("vendorcode", drv.Row.Item("vendorcode"))
                        bsVendorname.Position = itemfound

                    End If

                Case "Button9"
                    Dim myform = New FormHelper(bsDocTypeHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "doctypename"
                    If Not myform.ShowDialog = DialogResult.OK Then
                        ComboBox2.SelectedIndex = -1
                    Else
                        'checkcombobox3()
                        Dim drv As DataRowView = bsDocTypeHelper.Current
                        Dim mydrv As DataRowView = bsDetail.Current
                        mydrv.Row.Item("doctypename") = drv.Row.Item("doctypename")
                        mydrv.Row.Item("doctypeid") = drv.Row.Item("id")

                        'Sync with bsDocType
                        Dim itemfound As Integer = bsDocType.Find("id", drv.Row.Item("id"))
                        bsDocType.Position = itemfound
                        'bsDocType.Position = bsDocTypeHelper.Position
                        enabledTextBox()
                    End If

            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        DataGridView1.Invalidate()

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(bsDetail.Current) Then
            Dim drv As DataRowView = bsDetail.Current
            Dim mydownload = New SaveFileDialog
            If Not IsDBNull(drv.Row.Item("docname")) Then
                mydownload.FileName = drv.Row.Item("docname")
                mydownload.DefaultExt = drv.Row.Item("docext")
                Dim mysource = DS.Tables(9).Rows(0).Item("cvalue") & "\" & drv.Row.Item("documentid") & drv.Row.Item("docext")
                If mydownload.ShowDialog = DialogResult.OK Then
                    'MessageBox.Show(mydownload.FileName)
                    Try
                        FileIO.FileSystem.CopyFile(mysource, mydownload.FileName, True)
                        If MessageBox.Show("Show in folder?", "Locate Folder", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                            'Dim myrow As DataRowView = BS.Current                       
                            Dim p As New System.Diagnostics.Process
                            p.StartInfo.FileName = "explorer.exe"
                            p.StartInfo.Arguments = String.Format("{0},""{1}\{2}""", "/select", IO.Path.GetDirectoryName(mydownload.FileName), IO.Path.GetFileName(mydownload.FileName))
                            'Process.Start("explorer.exe", "/select," & "C:\temp\Documents\Forwarder\""" & DbAdapter1.validfilename(myfolder) & """")
                            'Process.Start("explorer.exe", "/select," & mybasefolder & """" & DbAdapter1.validfilename(myfolder) & """")
                            p.Start()
                        End If
                    Catch ex As Exception

                    End Try


                End If
            Else
                MessageBox.Show("No File Found!")
            End If


        End If
    End Sub

    Private Sub DateTimePicker3_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DateTimePicker3.Validating
        If Not DateTimePicker3.Checked Then
            Dim drv As DataRowView = bsDetail.Current
            If Not IsNothing(drv) Then
                drv.Row.Item("expireddate") = DBNull.Value
            End If

        End If
    End Sub

    Private Sub AddRecordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddRecordToolStripMenuItem.Click
        ToolStripButton1.PerformClick()
    End Sub

    Private Sub DeleteRecordToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteRecordToolStripMenuItem1.Click
        ToolStripButton4.PerformClick()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        If MessageBox.Show("Validate this task?", "Validate", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.OK Then
            For Each drv As DataRowView In bsDetail.List
                drv.Item("status") = 2
                drv.Item("statusname") = "Completed"
            Next
            DataGridView1.Invalidate()
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim ImportFile = New OpenFileDialog
        ImportFile.FileName = "Select File"
        If ImportFile.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim mydrv As DataRowView = bsDetail.Current
            Dim myobj As Object
            If mydrv.Row.Item("doctypeid") = SIF Then
                myobj = New ClassImportSIF(bsSiTx, bsDetail, ImportFile.FileName)
            Else
                myobj = New ClassImportIdentity(bsSiTx, bsDetail, ImportFile.FileName)
            End If
            If Not myobj.getrecord() Then
                MessageBox.Show(myobj.errormessage)
            Else
                
                DataGridView2.Invalidate()
                bsSiTx.Position = 0
            End If
        End If
    End Sub

    Private Sub DoImportSIFIdentity(ByRef bsSiTx As BindingSource, ByVal bsdetail As BindingSource, ByVal p2 As String)
        Dim dtdrv As DataRowView = bsdetail.Current

    End Sub


    Private Sub ComboBox6_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectionChangeCommitted
        TextBox15.Enabled = False
        If ComboBox6.SelectedItem = "Intertek" Then
            'TextBox15.Enabled = True
        End If
    End Sub

    Private Sub bsDetail_PositionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles bsDetail.PositionChanged

        If Not IsNothing(bsDetail.Current) Then
            Dim dr As DataRowView = bsDetail.Current
            TextBox15.Enabled = False
            If Not IsDBNull(dr.Item("auditby")) Then
                If dr.Item("auditby") = "Intertek" Then
                    'TextBox15.Enabled = True
                End If

            End If
        End If
    End Sub

End Class