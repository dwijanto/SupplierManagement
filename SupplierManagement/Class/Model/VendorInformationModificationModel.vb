Imports Npgsql
Imports System.Text

Public Class VendorInformationModificationModel
    Implements IModel

    Dim myadapter As DbAdapter = DbAdapter.getInstance
    Dim myhelper As HelperClass = HelperClass.getInstance

    Dim sb As New StringBuilder
    Private _vendorcode As Long



    Public ReadOnly Property sortField As String Implements IModel.sortField
        Get
            Return "id"
        End Get
    End Property

    Public ReadOnly Property tablename As String Implements IModel.tablename
        Get
            Return "doc.vendorinfmodi"
        End Get
    End Property

    Public Sub New()
    End Sub

    Public Sub New(ByVal vendorcode As Long)
        _vendorcode = vendorcode
    End Sub
    Public Function LoadData1(ByVal DS As System.Data.DataSet) As Boolean Implements IModel.LoadData
        Return True
    End Function

    Public Function LoadData(ByVal DS As System.Data.DataSet, ByVal hdid As Long) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            sb = New StringBuilder

            Dim mycriteria = String.Format("where v.vendorcode = {0}", _vendorcode)
            'Header 00
            sb.Append(String.Format("select u.*,v.vendorcode::text || ' - ' || v.vendorname::text as vendorcodename,v.vendorname::text,v.shortname::text,spm.username as spmusername,pd.username as pdusername,adb.username as dbusername,afc.username as fcusername,avp.username as vpusername ,app.email as appemail, doc.getmodifiedfield(u.id) as modifiedfield from {0} u left join vendor v on v.vendorcode = u.vendorcode" &
                                    " left join doc.user spm on lower(spm.userid) = lower(u.approvaldept)" &
                                    " left join doc.user pd on lower(pd.userid) = lower(u.approvaldept2)" &
                                    " left join doc.user adb on lower(adb.userid) = lower(u.approvaldb)" &
                                    " left join doc.user afc on lower(afc.userid) = lower(u.approvalfc)" &
                                    " left join doc.user avp on lower(avp.userid) = lower(u.approvalvp)" &
                                    " left join doc.user app on app.username = u.applicantname where u.id = {1};", tablename, hdid))
            'Detail 01
            sb.Append(String.Format("select u.*,m.sensitivitylevel,m.remarks from {0} u " &
                                    " left join doc.modificationtype m on m.id = u.fieldid" &
                                    " where u.hdid = {1} order by linenumber;", "doc.vendorinfmodidt", hdid))
            'Vendor 02
            sb.Append(String.Format("select distinct vf.*,v.shortname::text,v.vendorname,true as status from doc.vendorfactory vf" &
                                " left join vendor v on v.vendorcode = vf.vendorcode {0};", mycriteria))
            'VendorContact 03
            sb.Append(String.Format("select vc.*,v.shortname::text,v.vendorname from doc.vendorcontact vc" &
                                " left join vendor v on v.vendorcode = vc.vendorcode " &
                                " left join doc.contact c on c.id = vc.contactid" &
                                " {0} and c.isecoqualitycontact = true;", mycriteria))
            'Contact 04
            sb.Append(String.Format("select distinct c.* from doc.contact c" &
                                    " left join doc.vendorcontact vc on vc.contactid = c.id" &
                                    " left join vendor v on v.vendorcode = vc.vendorcode {0} and isecoqualitycontact = true;", mycriteria))


            'Document 05
            sb.Append(String.Format("with d as (select dt.paramdtid as id,dt.cvalue as doctype from doc.paramdt dt" &
                                    " left join doc.paramhd hd on hd.paramhdid = dt.paramhdid" &
                                    " where hd.paramname  = 'vendorinfomodiattachmenttype'" &
                                    " order by dt.ivalue)select false as download,va.*,'' as docfullname,d.doctype from doc.vendorinfmodiattachment va" &
                                    " left join d on d.id = va.doctypeid" &
                                    " where va.vendorinfmodiid = {0} order by va.id;", hdid))


            'DocumentType 06  dt.nvalue for type of email approval
            sb.Append(String.Format("select dt.paramdtid as id,dt.cvalue as doctype,dt.nvalue,dt.ivalue from doc.paramdt dt" &
                                    " left join doc.paramhd hd on hd.paramhdid = dt.paramhdid" &
                                    " where hd.paramname  = 'vendorinfomodiattachmenttype'" &
                                    " order by dt.cvalue;"))

            'FileSourceFullPath 07
            sb.Append("select h.cvalue,h.paramname from doc.paramhd h where h.paramname = 'vendorinfomodiattachmentfolder' order by ivalue;")

            'VendorinfomodiAction 08
            'doc.getstatusvendorinfmodi(vim.status)::text as statusname
            sb.Append(String.Format("select *,doc.getstatusvendorinfmodi(status)::text as statusname from doc.vendorinfomodiaction where vendorinfomodiid = {0} order by id desc;", hdid))
            Dim sqlstr = sb.ToString

            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS)
            
            myret = True
        End Using
        Return myret
    End Function

    Function LoadData(ByRef DS As DataSet, ByVal Criteria As String) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            sb = New StringBuilder

            'sb.Append(String.Format("select u.*,v.vendorname::text,adb.username as dbusername,afc.username as fcusername,avp.username as vpusername ,doc.getstatusvendorinfmodi(u.status) as lateststatus,cr.username as creatorname from {0} u left join vendor v on v.vendorcode = u.vendorcode" &
            '                        " left join doc.user adb on adb.userid = u.approvaldb" &
            '                        " left join doc.user afc on afc.userid = u.approvalfc" &
            '                        " left join doc.user avp on avp.userid = u.approvalvp " &
            '                        " left join doc.user cr on cr.userid = u.creator {1};", tablename, Criteria))
            sb.Append(String.Format("with q as (select u.*,v.vendorname::text,v.shortname::text,adb.username as dbusername,afc.username as fcusername,avp.username as vpusername ,doc.getstatusvendorinfmodi(u.status) as lateststatus,cr.username as creatorname, doc.getmodifiedfield(u.id) as modifiedfield from {0} u left join vendor v on v.vendorcode = u.vendorcode" &
                                   " left join doc.user adb on lower(adb.userid) = lower(u.approvaldb)" &
                                   " left join doc.user afc on lower(afc.userid) = lower(u.approvalfc)" &
                                   " left join doc.user avp on lower(avp.userid) = lower(u.approvalvp) " &
                                   " left join doc.user cr on lower(cr.userid) = lower(u.creator) ) select * from q {1};", tablename, Criteria))
            Dim sqlstr = sb.ToString

            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS)

            myret = True
        End Using
        Return myret
    End Function

    Public Function GetReportQuery(ByVal criteria) As String
        Dim sb As New StringBuilder
        'sb.Append(String.Format("with q as (select u.*,v.vendorname::text,adb.username as dbusername,afc.username as fcusername,avp.username as vpusername ,doc.getstatusvendorinfmodi(u.status) as lateststatus,cr.username as creatorname from {0} u left join vendor v on v.vendorcode = u.vendorcode" &
        '                          " left join doc.user adb on adb.userid = u.approvaldb" &
        '                          " left join doc.user afc on afc.userid = u.approvalfc" &
        '                          " left join doc.user avp on avp.userid = u.approvalvp " &
        '                          " left join doc.user cr on cr.userid = u.creator ) select * from q {1};", tablename, criteria))
        sb.Append(String.Format("with q as (select u.suppliermodificationid,u.vendorcode as vendorcode,v.vendorname::text,v.shortname::text,u.applicantname," &
                                " case m.informationtype when 1 then 'Basic Information' when 2 then 'Bank Information' else '' end as modificationtype," &
                                " m.modifytype,dt.newvalue,m.sensitivitylevel,u.applicantdate,doc.getstatusvendorinfmodi(u.status) as lateststatus,doc.getlateststatusmodifieddate(u.id,u.status) as lateststatusmodifieddate,cr.username as creatorname " &
                                " from doc.vendorinfmodi u left join doc.vendorinfmodidt dt on dt.id = u.id left join doc.modificationtype m on m.id = dt.fieldid" &
                                " left join vendor v on v.vendorcode = u.vendorcode left join doc.user cr on cr.userid = u.creator ) select * from q {0}", criteria))
        Return sb.ToString
    End Function

    Public Function LoadUserTask(ByRef DS As DataSet, ByVal userid As String, ByVal criteria As String) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            sb = New StringBuilder

            'sb.Append(String.Format("select vim.*,v.vendorname::text,v.shortname::text,doc.getstatusvendorinfmodi(vim.status)::text as statusname, doc.getmodifiedfield(vim.id) as modifiedfield from doc.vendorinfmodi vim" &
            '                        " left join vendor v on v.vendorcode = vim.vendorcode" &
            '                        " left join doc.user u on lower(u.userid) = '{0}'" &
            '                        " where status < 13 and (" &
            '                        " (vim.creator = lower(u.userid) and (status >= 8 and status <= 12 ))" &
            '                        " or (vim.approvaldept = lower(u.userid) and (status = 1 or status = 2))" &
            '                        " or (vim.approvaldept2 = lower(u.userid) and status = 3 )" &
            '                        " or (vim.approvaldb = lower(u.userid) and status = 4 )" &
            '                        " or (vim.approvalfc = lower(u.userid) and status = 5 )" &
            '                        " or (vim.approvalvp = lower(u.userid) and status = 6)) {1} ;", userid.ToLower, criteria))
            'sb.Append(String.Format("select vim.*,v.vendorname::text,v.shortname::text,doc.getstatusvendorinfmodi(vim.status)::text as statusname, doc.getmodifiedfield(vim.id) as modifiedfield from doc.vendorinfmodi vim" &
            '                        " left join vendor v on v.vendorcode = vim.vendorcode" &
            '                        " left join doc.user u on lower(u.userid) = '{0}'" &
            '                        " where status < 13 and (" &
            '                        " (vim.creator = lower(u.userid) and (status >= 8 and status <= 12 ))" &
            '                        " or (vim.approvaldept = lower(u.userid) and (status = 1 or status = 2))" &
            '                        " or (vim.approvaldept2 = lower(u.userid) and status = 3 )" &
            '                        " or (vim.approvaldb = lower(u.userid) and (status = 4 or status = 7) )" &
            '                        " or (vim.approvalfc = lower(u.userid) and status = 5 )" &
            '                        " or (vim.approvalvp = lower(u.userid) and status = 6)) {1} ;", userid.ToLower, criteria))


            'sb.Append(String.Format("select vim.*,v.vendorname::text,v.shortname::text,doc.getstatusvendorinfmodi(vim.status)::text as statusname,doc.getmodifiedfield(vim.id) as modifiedfield from doc.vendorinfmodi vim" &
            '                        " left join vendor v on v.vendorcode = vim.vendorcode" &
            '                        " left join doc.user u on lower(u.userid) = '{0}'" &
            '                        " where status < 13 and ( (vim.creator = lower(u.userid) and (status >= 8 and status <= 12 )) " &
            '                        " or (vim.approvaldept = lower(u.userid) and (status = 1 or status = 2)) or (vim.approvaldept2 = lower(u.userid) and status = 3 ) " &
            '                        " or (lower(u.userid) in (select lower(u.userid) from doc.groupuser gu" &
            '                        " left join doc.user u on u.id = gu.userid" &
            '                        " left join doc.groupauth g on g.groupid = gu.groupid" &
            '                        " where groupname = 'Approval Database')" &
            '                        " and (status = 4 or status = 7 or (status = 6 and sensitivitylevel = 2 ) or (status = 6 and sensitivitylevel = 1 and turnovervalue < 5000000)) ) or (vim.approvalfc = lower(u.userid) and status = 5 ) " &
            '                        " or (vim.approvalvp = lower(u.userid) and status = 6)) {1};", userid.ToLower, criteria))
            sb.Append(String.Format("select vim.*,v.vendorname::text,v.shortname::text,doc.getstatusvendorinfmodi(vim.status)::text as statusname,doc.getmodifiedfield(vim.id) as modifiedfield from doc.vendorinfmodi vim" &
                                    " left join vendor v on v.vendorcode = vim.vendorcode" &
                                    " left join doc.user u on lower(u.userid) = '{0}'" &
                                    " where status < 13 and ( (lower(vim.creator) = lower(u.userid) and (status >= 8 and status <= 12 )) " &
                                    " or (vim.approvaldept = lower(u.userid) and (status = 1 or status = 2)) or (vim.approvaldept2 = lower(u.userid) and status = 3 ) " &
                                    " or (lower(u.userid) in (select lower(u.userid) from doc.groupuser gu" &
                                    " left join doc.user u on u.id = gu.userid" &
                                    " left join doc.groupauth g on g.groupid = gu.groupid" &
                                    " where groupname = 'Approval Database')" &
                                    " and (status = 4 or status = 7 or (status = 6 and sensitivitylevel = 2 ) or (status = 3 and sensitivitylevel = 3 )  or (status = 6 and sensitivitylevel = 1 and turnovervalue < 5000000)) ) or (vim.approvalfc = lower(u.userid) and status = 5 ) " &
                                    " or (vim.approvalvp = lower(u.userid) and status = 6)) {1};", userid.ToLower, criteria))

            'sb.Append(String.Format("select vim.*,v.vendorname::text,v.shortname::text,doc.getstatusvendorinfmodi(vim.status)::text as statusname, doc.getmodifiedfield(vim.id) as modifiedfield from doc.vendorinfmodi vim" &
            '                       " left join vendor v on v.vendorcode = vim.vendorcode" &
            '                       " left join doc.user u on lower(u.userid) = '{0}'" &
            '                       " where status >= 3 and " &
            '                       " ((vim.applicantname = u.username and (status >= 8 and status <= 14))" &
            '                       " or (vim.approvaldept = lower(u.userid) and (status >= 3 or status <= 14))" &
            '                       " or (vim.approvaldept2 = lower(u.userid) and (status >=4 and status <=14))" &
            '                       " or (vim.approvaldb = lower(u.userid) and (status >=5 and status <=14))" &
            '                       " or (vim.approvalfc = lower(u.userid) and (status >=6 and status <=14))" &
            '                       " or (vim.approvalvp = lower(u.userid) and (status >=7 and status <=14))) {1};", userid.ToLower, criteria))
            'sb.Append(String.Format("select vim.*,v.vendorname::text,v.shortname::text,doc.getstatusvendorinfmodi(vim.status)::text as statusname, doc.getmodifiedfield(vim.id) as modifiedfield from doc.vendorinfmodi vim" &
            '                       " left join vendor v on v.vendorcode = vim.vendorcode" &
            '                       " left join doc.user u on lower(u.userid) = '{0}'" &
            '                       " where ((lower(vim.creator) = lower(u.userid) or vim.applicantname = u.username) and status >= 1)  " &
            '                       " or (lower(vim.approvaldept) = lower(u.userid) and (status >= 3 or status <= 14))" &
            '                       " or (lower(vim.approvaldept2) = lower(u.userid) and (status >=4 and status <=14))" &
            '                       " or (lower(vim.approvaldb) = lower(u.userid) and (status >=5 and status <=14))" &
            '                       " or (lower(vim.approvalfc) = lower(u.userid) and (status >=6 and status <=14))" &
            '                       " or (lower(vim.approvalvp) = lower(u.userid) and (status >=7 and status <=14)) {1};", userid.ToLower, criteria))
            sb.Append(String.Format("select vim.*,v.vendorname::text,v.shortname::text,doc.getstatusvendorinfmodi(vim.status)::text as statusname, doc.getmodifiedfield(vim.id) as modifiedfield from doc.vendorinfmodi vim" &
                                   " left join vendor v on v.vendorcode = vim.vendorcode" &
                                   " left join doc.user u on lower(u.userid) = '{0}'" &
                                   " where ((lower(vim.creator) = lower(u.userid) or vim.applicantname = u.username) and status >= 1)  " &
                                   " or (lower(vim.approvaldept) = lower(u.userid) and (status >= 3 or status <= 14))" &
                                   " or (lower(vim.approvaldept2) = lower(u.userid) and (status >=4 and status <=14))" &
                                   " or (lower(u.userid) in (select lower(u.userid) from doc.groupuser gu" &
                                    " left join doc.user u on u.id = gu.userid" &
                                    " left join doc.groupauth g on g.groupid = gu.groupid" &
                                    " where groupname = 'Approval Database') and ((status >=5 and status <=14)  ))" &
                                   " or (lower(vim.approvalfc) = lower(u.userid) and (status >=6 and status <=14))" &
                                   " or (lower(vim.approvalvp) = lower(u.userid) and (status >=7 and status <=14)) {1};", userid.ToLower, criteria))
            Dim sqlstr = sb.ToString

            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS)

            myret = True
        End Using
        Return myret
    End Function

    Public Function LoadUserTask(ByRef DS As DataSet, ByVal criteria As String) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            sb = New StringBuilder

            sb.Append(String.Format("select vim.*,v.vendorname::text,v.shortname::text,doc.getstatusvendorinfmodi(vim.status)::text as statusname, doc.getmodifiedfield(vim.id) as modifiedfield from doc.vendorinfmodi vim" &
                                    " left join vendor v on v.vendorcode = vim.vendorcode" &
                                    " where (status <= 12) {0};", criteria))
            sb.Append(String.Format("select vim.*,v.vendorname::text,v.shortname::text,doc.getstatusvendorinfmodi(vim.status)::text as statusname, doc.getmodifiedfield(vim.id) as modifiedfield from doc.vendorinfmodi vim" &
                                   " left join vendor v on v.vendorcode = vim.vendorcode" &
                                   " where (status > 12) {0};", criteria))
            Dim sqlstr = sb.ToString

            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS)

            myret = True
        End Using
        Return myret
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myadapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myadapter.getConnection

            conn.Open()
            mytransaction = conn.BeginTransaction

            'Header
            Dim sqlstr = "doc.sp_updatevendorinfmodi"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "suppliermodificationid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "applicantname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "applicantdate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familycode").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "subfamilycode").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "ismissing").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "currency").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "turnovertype").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "yearreference").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovervalue").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "ecoqualitycontactname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "ecoqualitycontactemail").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "sensitivitylevel").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvaldept").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvaldept2").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvaldb").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvalfc").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvalvp").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current            
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_insertvendorinfmodi"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("ivendorcode", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "applicantname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "applicantdate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familycode").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "subfamilycode").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "ismissing").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "currency").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "turnovertype").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "yearreference").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovervalue").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "ecoqualitycontactname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "ecoqualitycontactemail").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "sensitivitylevel").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvaldept").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvaldept2").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvaldb").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvalfc").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvalvp").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "creator").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("isuppliermodificationid", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "suppliermodificationid").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.Parameters.Add("iid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_deletevendorinfmodi"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(0))

            'Detail
            sqlstr = "doc.sp_updatevendorinfmodidt"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "hdid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "fieldid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "oldvalue").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "newvalue").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_insertvendorinfmodidt"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "hdid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "fieldid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "oldvalue").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "newvalue").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_deletevendorinfmodidt"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(1))


            'Document 
            sqlstr = "doc.sp_updatevendorinfmodiattachment"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorinfmodiid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "docdate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docext").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doctypeid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remarks").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_insertvendorinfmodiattachment"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorinfmodiid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "docdate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docext").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doctypeid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remarks").SourceVersion = DataRowVersion.Current            
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_deletevendorinfmodiattachment"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(5))


            RemoveHandler dataadapter.RowUpdated, AddressOf myadapter.onRowInsertUpdate


            'Contact
            sqlstr = "doc.sp_updatecontact"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "contactname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "title").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "officeph").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "factoryph").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "officemb").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "factorymb").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isecoqualitycontact").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName
            'dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current


            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_insertcontact"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "contactname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "title").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "officeph").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "factoryph").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "officemb").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "factorymb").SourceVersion = DataRowVersion.Current            
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isecoqualitycontact").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_deletecontact"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(4))


            ''Vendor Contact
            'sqlstr = "doc.sp_updatevendorcontact"
            'dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)            
            'dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            'dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "contactid").SourceVersion = DataRowVersion.Current
            'dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            'dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_insertvendorcontact"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "contactid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName            
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_deletevendorcontact"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction            
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(3))

            'Action 
            'AssetPurchaseAction
            AddHandler dataadapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf myadapter.onRowInsertUpdate)
            sqlstr = "doc.sp_insertvendorinfomodiaction"
            dataadapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorinfomodiid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remark").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure
            mye.ra = dataadapter.Update(mye.dataset.Tables("VendorInfoModiAction"))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret




    End Function



End Class
