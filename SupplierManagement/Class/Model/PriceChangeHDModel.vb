Imports Npgsql
Imports System.Text
Imports SupplierManagement.PublicClass 'for HelperClass
Public Class PriceChangeHDModel
    Implements IModel

    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Public Property shortname As String
    Private sqlstr As String = String.Empty

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "pricechangehd"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "pricechangehdid"
        End Get
    End Property

    Public ReadOnly Property FilterField
        Get
            Return Nothing
        End Get
    End Property

    Public Function GetQuery() As String
        Dim mysql = String.Format("select actionid as ""ActionId"",shortname ,priority as ""Priority (H/M/L)"",situation as ""Situation"",target as ""Target"",proposal as ""Proposal/Action Plan"",responsibleperson as ""Responsible Person"",startdate  as ""Start Date"",enddate as ""Expected End Date"",result as ""Actual Result / Evidence / Status"",finishdate as ""Actual Finish Date"",status as ""Status (Not Started / On-going / Delay/ Closed)"",uploaddate as ""Update Date"",modifiedby as ""Update By"" from ({0}) as foo", sqlstr)
        Return mysql
    End Function

    Public Function GetQueryALL() As String
        Dim mysql = String.Format("select actionid as ""ActionId"",shortname ,priority as ""Priority (H/M/L)"",situation as ""Situation"",target as ""Target"",proposal as ""Proposal/Action Plan"",responsibleperson as ""Responsible Person"",startdate  as ""Start Date"",enddate as ""Expected End Date"",result as ""Actual Result / Evidence / Status"",finishdate as ""Actual Finish Date"",status as ""Status (Not Started / On-going / Delay/ Closed)"",uploaddate as ""Update Date"",modifiedby as ""Update By"" from doc.actionplan order by shortname, uploaddate")
        Return mysql
    End Function

    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()

            sqlstr = String.Format("select u.* from {0} u order by {1}", TableName, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function LoadData(ByVal DS As DataSet, ByVal Criteria As String) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim mysb As New StringBuilder
            mysb.Append(String.Format("with vr as (select distinct first_value(id) over (partition by shortname,actionid order by id desc) as id from doc.actionplan)"))
            mysb.Append(String.Format("(select null::text as spacer,ac.id,ac.documentid,ac.priority,ac.situation,ac.target,ac.proposal,ac.responsibleperson,ac.startdate,ac.enddate,ac.result,ac.finishdate,ac.status,ac.actionid,modifiedby,uploaddate,ac.shortname,ac.validator,ac.validatorflag,ac.statuschangedate from doc.actionplan ac inner join vr on vr.id = ac.id  {0} order by id asc)", Criteria))

            sqlstr = mysb.ToString 'String.Format("select u.* from {0} u {1} order by {2}", TableName, Criteria, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function LoadDataVendorUser(ByVal DS As DataSet, ByVal Criteria As String) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim mysb As New StringBuilder
            mysb.Append(String.Format("with vu as (select shortname from doc.getvendoruser('{0}')), vr as (select distinct first_value(id) over (partition by shortname,actionid order by id desc) as id from doc.actionplan)", HelperClass1.UserInfo.userid & "$"))
            mysb.Append(String.Format("(select null::text as spacer,ac.id,ac.documentid,ac.priority,ac.situation,ac.target,ac.proposal,ac.responsibleperson,ac.startdate,ac.enddate,ac.result,ac.finishdate,ac.status,ac.actionid,modifiedby,uploaddate,ac.shortname,ac.validator,ac.validatorflag,ac.statuschangedate from doc.actionplan ac inner join vr on vr.id = ac.id {0} and shortname in (select shortname from vu) order by id asc)", Criteria))

            sqlstr = mysb.ToString 'String.Format("select u.* from {0} u {1} order by {2}", TableName, Criteria, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function
    Public Function LoadData(ByVal DS As DataSet, ByVal Criteria As String, ByVal shortname As String) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim mysb As New StringBuilder
            mysb.Append(String.Format("with vr as (select distinct first_value(id) over (partition by shortname,actionid order by id desc) as id from doc.actionplan where trim(shortname) in ({0}))", shortname))
            mysb.Append(String.Format("(select null::text as spacer,ac.id,ac.documentid,ac.priority,ac.situation,ac.target,ac.proposal,ac.responsibleperson,ac.startdate,ac.enddate,ac.result,ac.finishdate,ac.status,ac.actionid,modifiedby,uploaddate,ac.shortname,ac.validator,ac.validatorflag,ac.statuschangedate from doc.actionplan ac inner join vr on vr.id = ac.id  {0} order by id asc)", Criteria))

            sqlstr = mysb.ToString 'String.Format("select u.* from {0} u {1} order by {2}", TableName, Criteria, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function getStatusBS() As BindingSource
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Dim myds As New DataSet
        Dim myBS As New BindingSource
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select distinct u.status from {0} u order by u.status;", TableName)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(myds, TableName)
            myret = True
        End Using
        myBS.DataSource = myds.Tables(0)
        Return myBS
    End Function


    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myadapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myadapter.getConnection
            conn.Open()
            mytransaction = conn.BeginTransaction
            'ActionPlan
            Dim sqlstr As String
            sqlstr = "doc.sp_deleteactionplan"
            dataadapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_insertactionplan"
            dataadapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "priority").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "situation").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "target").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "proposal").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "responsibleperson").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "startdate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "enddate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "result").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "finishdate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "status").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "uploaddate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "validatorflag").SourceVersion = DataRowVersion.Current

            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_insertactionplan"
            dataadapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "priority").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "situation").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "target").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "proposal").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "responsibleperson").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "startdate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "enddate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "result").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "finishdate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "status").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "uploaddate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "validatorflag").SourceVersion = DataRowVersion.Current

            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(TableName))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function
End Class
