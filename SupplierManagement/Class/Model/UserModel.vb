Imports Npgsql
Public Class UserModel
    Implements IModel

    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Public Function getApplicantBS() As BindingSource
        Dim sqlstr As String = "select ''::text as username union all (select u.username from doc.user u " &
                 " left join officerseb o on o.userid = u.userid left join masteruser mu on mu.id = o.muid" &
                 " left join teamtitle tt on tt.teamtitleid = o.teamtitleid where teamtitleshortname in ('PD','SPM','PM','PO','WMF','PCL') and u.isactive  order by u.username);"
        Dim DS As New DataSet
        Dim bs As New BindingSource
        If myadapter.TbgetDataSet(sqlstr, DS) Then
            bs.DataSource = DS.Tables(0)
        End If
        Return bs
    End Function

    Public Function LoadData(ByVal DS As System.Data.DataSet) As Boolean Implements IModel.LoadData
        Return False
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Return Nothing
    End Function

    Public ReadOnly Property sortField As String Implements IModel.sortField
        Get
            Return "username"
        End Get
    End Property

    Public ReadOnly Property tablename As String Implements IModel.tablename
        Get
            Return "doc.user"
        End Get
    End Property

    Public Function getApprovalDirectorBS() As BindingSource
        Dim sqlstr As String = "select u.username,u.userid from doc.user u" &
                               " left join doc.groupuser gu on gu.userid = u.id" &
                               " left join doc.groupauth ga on ga.groupid = gu.groupid" &
                               " where groupname = 'Approval Director'" &
                               " order by username"
        Return getBindingsource(sqlstr)
    End Function

    Public Function getCCDBBS() As BindingSource
        Dim sqlstr As String = "select u.username,u.userid from doc.user u" &
                               " left join doc.groupuser gu on gu.userid = u.id" &
                               " left join doc.groupauth ga on ga.groupid = gu.groupid" &
                               " where groupname = 'Approval Database CC'" &
                               " order by username limit 1"
        Return getBindingsource(sqlstr)
    End Function

    Public Function getApprovalDBBS() As BindingSource
        Dim sqlstr As String = "select u.username,u.userid from doc.user u" &
                               " left join doc.groupuser gu on gu.userid = u.id" &
                               " left join doc.groupauth ga on ga.groupid = gu.groupid" &
                               " where groupname = 'Approval Database'" &
                               " order by username limit 1"
        Return getBindingsource(sqlstr)
    End Function

    Public Function getApprovalDBListBS() As BindingSource
        Dim sqlstr As String = "select u.username,u.userid,u.email from doc.user u" &
                               " left join doc.groupuser gu on gu.userid = u.id" &
                               " left join doc.groupauth ga on ga.groupid = gu.groupid" &
                               " where groupname = 'Approval Database'" &
                               " order by username"
        Return getBindingsource(sqlstr)
    End Function

    Public Function getApprovalDBVIMBS() As BindingSource
        Dim sqlstr As String = "select u.username,u.userid from doc.user u" &
                               " left join doc.groupuser gu on gu.userid = u.id" &
                               " left join doc.groupauth ga on ga.groupid = gu.groupid" &
                               " where groupname =  'Approval Database (Vendor Modification)'" &
                               " order by username limit 1"
        Return getBindingsource(sqlstr)
    End Function

    Public Function getApprovalFCBS() As BindingSource
        Dim sqlstr As String = "select u.username,u.userid from doc.user u" &
                               " left join doc.groupuser gu on gu.userid = u.id" &
                               " left join doc.groupauth ga on ga.groupid = gu.groupid" &
                               " where groupname = 'Approval Financial Controller'" &
                               " order by username limit 1"
        Return getBindingsource(sqlstr)
    End Function

    Public Function getApprovalVPBS() As BindingSource
        Dim sqlstr As String = "select u.username,u.userid from doc.user u" &
                               " left join doc.groupuser gu on gu.userid = u.id" &
                               " left join doc.groupauth ga on ga.groupid = gu.groupid" &
                               " where groupname = 'Approval VP'" &
                               " order by username limit 1"
        Return getBindingsource(sqlstr)
    End Function

    Public Function getApprovalDeptBS() As BindingSource
        Dim sqlstr As String = "select u.username,u.userid from doc.user u" &
                               " left join doc.groupuser gu on gu.userid = u.id" &
                               " left join doc.groupauth ga on ga.groupid = gu.groupid" &
                               " where groupname = 'Approval Dept'" &
                               " order by username"
        Return getBindingsource(sqlstr)
    End Function

    Private Function getBindingsource(ByVal sqlstr As String) As BindingSource
        Dim DS As New DataSet
        Dim bs As New BindingSource
        If myadapter.TbgetDataSet(sqlstr, DS) Then
            bs.DataSource = DS.Tables(0)
        End If
        Return bs
    End Function
End Class
