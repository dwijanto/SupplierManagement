Imports Npgsql
Public Class OfficerSEBModel
    Implements IModel
    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Public Function LoadData(ByVal DS As System.Data.DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select * from  {0} order by {1};", tablename, sortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            myret = True
        End Using
        Return myret
    End Function

    Public Function GetValidator() As BindingSource
        '"select ''::text as name,'' as userid,null as teamtitleid,'' as officersebname,''::text as teamtitleshortname union all (select distinct teamtitleshortname || ' - ' || officersebname as name,lower(o.userid) as userid,tt.teamtitleid,officersebname,tt.teamtitleshortname from doc.user u left join officerseb o on o.userid = u.userid  left join teamtitle tt on tt.teamtitleid = o.teamtitleid where teamtitleshortname in ('PD','SPM','PM','PO') and o.isactive and o.userid <> 'as\lili2' order by tt.teamtitleid,officersebname);"
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Dim ds As New DataSet
        Dim bs As New BindingSource
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select ''::text as name,'' as userid,null as teamtitleid,'' as officersebname,''::text as teamtitleshortname union all (select distinct teamtitleshortname || ' - ' || officersebname as name,lower(o.userid) as userid,tt.teamtitleid,officersebname,tt.teamtitleshortname from doc.user u left join officerseb o on o.userid = u.userid  left join teamtitle tt on tt.teamtitleid = o.teamtitleid where teamtitleshortname in ('PD','SPM','PM','PCL') and u.isactive and o.userid <> 'as\lili2' order by tt.teamtitleid,officersebname);")
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(ds, tablename)
            bs.DataSource = ds.Tables(0)
            myret = True
        End Using
        Return bs
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Return Nothing
    End Function

    Public ReadOnly Property sortField As String Implements IModel.sortField
        Get
            Return "ofsebid"
        End Get
    End Property

    Public ReadOnly Property tablename As String Implements IModel.tablename
        Get
            Return "officerseb"
        End Get
    End Property
End Class
