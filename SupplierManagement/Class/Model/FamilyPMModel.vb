Imports Npgsql

Public Class FamilyPMModel
    Implements IModel
    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Public ReadOnly Property TableName
        Get
            Return "FamilyPM"
        End Get
    End Property

    Public ReadOnly Property SortField
        Get
            Return "familyid"
        End Get
    End Property

    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("(select 0::bigint as familyid,null::bigint as pmid,'To be define'::text as familyname,'To be define'::character varying as familydesc,'To be define'::character varying as pm,'To be define'::character varying as spm,'To be define'::text as familydesc2) union all " &
                                       "(select fp.*,f.familyname::text,f.familyid::text || ' - ' || f.familyname::text as familydesc,mu.username as pm,mu2.username as spm, f.familyid::text || ' - ' || f.familyname::text || ' - ' || mu.username as familydesc2   from doc.familypm fp " &
                         " left join family f on f.familyid = fp.familyid" &
                         " left join officerseb o on o.ofsebid = fp.pmid " &
                         " left join masteruser mu on mu.id = o.muid " &
                         " left join officerseb spm on spm.ofsebid = o.parent " &
                         " left join masteruser mu2 on mu2.id = spm.muid " &
                         " order by {0})", SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
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

            'Update
            Dim sqlstr = "doc.sp_updatefamilypm"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pmid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_insertfamilypm"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)            
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pmid").SourceVersion = DataRowVersion.Current            
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_deletefamilypm"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(TableName))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function


    Public ReadOnly Property sortField1 As String Implements IModel.sortField
        Get
            Return True
        End Get
    End Property

    Public ReadOnly Property tablename1 As String Implements IModel.tablename
        Get
            Return True
        End Get
    End Property
End Class
