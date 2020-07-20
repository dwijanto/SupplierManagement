Imports Npgsql

Public Class FamilyModel
    Implements IModel

    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "Family"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "familyid"
        End Get
    End Property

    Public Function getFamilyBS() As BindingSource
        Dim sqlstr As String = String.Format("select 0::integer as familyid,null::text as familyname,null::text as familydesc union all (select familyid,familyname::text ,familyid::text || ' - ' || familyname::text as familydesc from family order by {0});", SortField)
        Dim DS As New DataSet
        Dim bs As New BindingSource
        If myadapter.TbgetDataSet(sqlstr, DS) Then
            bs.DataSource = DS.Tables(0)
        End If
        Return bs
    End Function

    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select familyid,familyname::text ,familyid::text || ' - ' || familyname::text as familydesc from family order by {0};", SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Return Nothing
    End Function

  

End Class
