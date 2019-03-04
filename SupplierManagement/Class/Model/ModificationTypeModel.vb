Imports Npgsql
Imports System.Text

Public Class ModificationTypeModel
    Implements IModel
    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Public Function LoadData(ByVal DS As System.Data.DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Dim SB As New StringBuilder
        Using conn As Object = myadapter.getConnection
            conn.Open()
            SB.Append(String.Format("select u.*,case u.informationtype when 1 then 'Basic Information' when 2 then 'Bank Information' end as informationtypename from  doc.modificationtype u order by {0};", sortField))
            SB.Append(String.Format("with dtl as (select d.* from doc.paramdt d left join doc.paramhd h on h.paramhdid = d.paramhdid where h.paramname = 'vendorinfomodiattachmenttype' order by d.ivalue) select u.*,dtl.cvalue as doctypename from doc.modificationtypedtl u left join dtl on dtl.paramdtid = u.doctypeid ;"))
            SB.Append("select d.* from doc.paramdt d left join doc.paramhd h on h.paramhdid = d.paramhdid where h.paramname = 'vendorinfomodiattachmenttype' order by d.ivalue")
            Dim sqlstr = SB.ToString
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
            'ModificationType HD
            Dim sqlstr As String
            sqlstr = "doc.sp_deletemodificationtype"
            dataadapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_insertmodificationtype"
            dataadapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "informationtype").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifytype").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "sensitivitylevel").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineorder").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "remarks").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_updatemodificationtype"
            dataadapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "informationtype").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifytype").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "sensitivitylevel").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineorder").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "remarks").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(0))


            'DTL
            sqlstr = "doc.sp_deletemodificationtypedtl"
            dataadapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_insertmodificationtypedtl"
            dataadapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "modificationtypeid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doctypeid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_updatemodificationtypedtl"
            dataadapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "modificationtypeid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doctypeid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(1))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Public ReadOnly Property sortField As String Implements IModel.sortField
        Get
            Return "lineorder"
        End Get
    End Property

    Public ReadOnly Property tablename As String Implements IModel.tablename
        Get
            Return "doc.modificationtype"
        End Get
    End Property

    Public ReadOnly Property FilterField
        Get
            Return Nothing
        End Get
    End Property

    Public Function GetModificationTypeBS() As BindingSource
        Dim sqlstr = String.Format("select u.*,case u.informationtype when 1 then 'Basic Information' else 'Bank Information ' end  || ' - ' || u.modifytype as description from {0} u order by {1}", tablename, sortField)
        Dim DS As New DataSet
        Dim bs As New BindingSource
        myadapter.TbgetDataSet(sqlstr, DS)
        bs.DataSource = DS.Tables(0)
        Return bs
    End Function

    Public Function GetDocTypeId(ByVal modificationtype As Integer) As Hashtable
        Dim sb As New StringBuilder
        Dim myret As New Hashtable
        Dim sqlstr = String.Format("with dtl as (select d.* from doc.paramdt d left join doc.paramhd h on h.paramhdid = d.paramhdid where h.paramname = 'vendorinfomodiattachmenttype' order by d.ivalue) " &
                                   " select u.*,dtl.cvalue as doctypename from doc.modificationtypedtl u left join dtl on dtl.paramdtid = u.doctypeid   where modificationtypeid = :modificationtype")
        Dim ds As New DataSet
        Dim bs As New BindingSource
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("modificationtype", modificationtype)
        If myadapter.TbgetDataSet(sqlstr, ds, myparam) Then
            bs.DataSource = ds.Tables(0)
            For Each drv In bs.List
                myret.Add(drv.row.item("doctypeid"), drv.row.item("doctypename"))
            Next
        End If
        Return myret
    End Function
    Public Function GetRemarks(ByVal id As Integer) As String
        Dim sb As New StringBuilder
        Dim myret As String = String.Empty
        Dim sqlstr = String.Format("select informationtype::text || '|' || case when remarks isnull then '' else remarks end from doc.modificationtype u  where id = :id")
        Dim ds As New DataSet
        Dim bs As New BindingSource
        Dim myparam(0) As NpgsqlParameter
        myparam(0) = New NpgsqlParameter("id", id)

        If myadapter.ExecuteScalar(sqlstr, myret, myparam) Then
           
        End If       
        Return myret
    End Function
End Class
