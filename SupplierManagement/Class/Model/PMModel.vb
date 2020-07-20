Imports Npgsql

Public Class PMModel
    Implements IModel
    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Public ReadOnly Property TableName
        Get
            Return "PM"
        End Get
    End Property

    Public ReadOnly Property SortField
        Get
            Return "PM"
        End Get
    End Property

    Public Function getPMBS() As BindingSource
        Dim sqlstr As String = String.Format("select 0 as pmid,null::text as pm,null::text as smp union all (select o.ofsebid as pmid,mu.username as pm ,mu2.username as spm from officerseb o" &
                  " left join masteruser mu on mu.id = o.muid" &
                  " left join teamtitle tt on tt.teamtitleid = o.teamtitleid" &
                  " left join officerseb spm on spm.ofsebid = o.parent" &
                  " left join masteruser mu2 on mu2.id = spm.muid" &
                  " where teamtitleshortname in ('PM','PCL') and mu.isactive order by pm);")
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
            Dim sqlstr = "select o.ofsebid as pmid,mu.username as pm ,mu2.username as spm from officerseb o" &
                  " left join masteruser mu on mu.id = o.muid" &
                  " left join teamtitle tt on tt.teamtitleid = o.teamtitleid" &
                  " left join officerseb spm on spm.ofsebid = o.parent" &
                  " left join masteruser mu2 on mu2.id = spm.muid" &
                  " where teamtitleshortname in ('PM','PCL') and mu.isactive order by pm;"
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
