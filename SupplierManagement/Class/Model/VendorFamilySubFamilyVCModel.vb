Imports Npgsql
Imports System.Text

Public Class VendorFamilySubFamilyVCModel
    Implements IModel
    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Public ReadOnly Property TableName
        Get
            Return "vendorfamilysubfamilyvc"
        End Get
    End Property

    Public ReadOnly Property SortField
        Get
            Return "id"
        End Get
    End Property

    Public ReadOnly Property FilterField
        Get
            Return "[vendorname] like '*{0}*' or [description] like '*{0}*' or [familyvc] like '*{0}*' or [subfamilyvc] like '*{0}*'"
        End Get
    End Property
    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select * from doc.vendorfamilysubfamilyvc vfvc" &
                                       " left join doc.familyvc fvc on fvc.familyfc = vfvc.familyvc" &
                                       " left join doc.subfamilyvc sfvc on sfvc.subfamilyvc = vfvc.subfamilyvc" &
                                       " order by {0};", SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function
    Public Function LoadData(ByVal DS As DataSet, ByVal vendorcode As Long) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select vfvc.*,fvc.description as familyvcdescription,sfvc.description as subfamilyvcdescription from doc.vendorfamilysubfamilyvc vfvc" &
                                       " left join doc.familyvc fvc on fvc.familyvc = vfvc.familyvc" &
                                       " left join doc.subfamilyvc sfvc on sfvc.subfamilyvc = vfvc.subfamilyvc" &
                                       " where vendorcode = {1}" &
                                       " order by {0};", SortField, vendorcode)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function GetFamilyCode(ByVal vendorcode As Long) As String
        Dim result As New StringBuilder
        Dim DS As New DataSet
        LoadData(DS, vendorcode)
        For i = 0 To DS.Tables(0).Rows.Count - 1
            If result.Length > 1 Then result.Append(",")
            result.Append(DS.Tables(0).Rows(i).Item("familyvc"))
        Next
        Return result.ToString
    End Function

    Public Function GetSubFamilyCode(ByVal vendorcode As Long) As String
        Dim result As New StringBuilder
        Dim DS As New DataSet
        LoadData(DS, vendorcode)
        For i = 0 To DS.Tables(0).Rows.Count - 1
            If result.Length > 1 Then result.Append(",")
            result.Append(DS.Tables(0).Rows(i).Item("subfamilyvc"))
        Next
        Return result.ToString
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
            Dim sqlstr = "doc.sp_updatevendorfamilysubfamilyvc"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyvc").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "subfamilyvc").SourceVersion = DataRowVersion.Current

            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_insertvendorfamilysubfamilyvc"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyvc").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "subfamilyvc").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "doc.sp_deletevendorfamilysubfamilyvc"
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
