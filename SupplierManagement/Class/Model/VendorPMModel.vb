﻿Imports Npgsql

Public Class VendorPMModel
    Implements IModel
    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Public ReadOnly Property TableName
        Get
            Return "VENDORPM"
        End Get
    End Property

    Public ReadOnly Property SortField
        Get
            Return "vendorcode"
        End Get
    End Property

    Public ReadOnly Property FilterField
        Get
            Return "[vendorcode] like '*{0}*' or [vendorname] like '*{0}*' or [shortname] like '*{0}*' or [pm] like '*{0}*'  or [spm] like '*{0}*'"
        End Get
    End Property
    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            'Dim sqlstr = String.Format("select vp.id,vp.pmid,vp.vendorcode::text,vp.pmeffectivedate,vp.spmeffectivedate,mu.username as pm,mus.username as spm,v.vendorname::text,v.shortname::text  from doc.vendorpm vp " &
            '                           " left join vendor v on v.vendorcode = vp.vendorcode" &
            '                           " left join officerseb o on o.ofsebid = vp.pmid" &
            '                           " left join masteruser mu on mu.id = o.muid" &
            '                           " left join officerseb spm on spm.ofsebid = o.parent" &
            '                           " left join masteruser mus on mus.id = spm.muid order by {0};", SortField)
            Dim sqlstr = String.Format("select vp.id,vp.pmid,vp.vendorcode::text, coalesce(vp.pmeffectivedate,(select applicantdate from doc.vendorinfmodi where vendorcode = vp.vendorcode limit 1))as pmeffectivedate," &
                                       " coalesce(vp.spmeffectivedate,(select applicantdate from doc.vendorinfmodi where vendorcode = vp.vendorcode limit 1))as spmeffectivedate,mu.username as pm,mus.username as spm,v.vendorname::text,v.shortname3::text as shortname from doc.vendorpm vp " &
                                       " left join vendor v on v.vendorcode = vp.vendorcode" &
                                       " left join officerseb o on o.ofsebid = vp.pmid" &
                                       " left join masteruser mu on mu.id = o.muid" &
                                       " left join officerseb spm on spm.ofsebid = o.parent" &
                                       " left join masteruser mus on mus.id = spm.muid order by {0};", SortField)

           
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
            Dim sqlstr = "doc.sp_updatevendorpm"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pmid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "pmeffectivedate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "spmeffectivedate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_insertvendorpm"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pmid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "pmeffectivedate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "spmeffectivedate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_deletevendorpm"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.Input
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
