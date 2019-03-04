Imports Npgsql

Public Class VendorModel
    Implements IModel

    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "Vendor"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "vendorcode"
        End Get
    End Property
    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            'Dim sqlstr = String.Format("select v.*,v.vendorcode::text || ' - ' || v.vendorname::text as vendordesc from vendor v order by {0}", SortField)
            Dim vendorfamilysql = String.Format("select vf.id,vf.vendorcode,vf.familyid,f.familyname, vf.familyid ||  ' - ' || f.familyname as familydesc,v.vendorname::text,v.shortname::text,vf.vendorcode || ' - ' || v.vendorname::text || (" &
                                    " case when vf.familyid isnull then ' '  else  ' - '  ||  vf.familyid ||  ' - ' || f.familyname  end )  || (case when mu.username isnull then ' ' else ' - ' || mu.username end ) as description from doc.vendorfamily vf" &
                                    " left join vendor v on v.vendorcode = vf.vendorcode left join family f on f.familyid = vf.familyid left join doc.familypm fpm on fpm.familyid = vf.familyid" &
                                    " left join officerseb o on o.ofsebid = fpm.pmid left join masteruser mu on mu.id = o.muid order by vf.vendorcode,vf.familyid;")
            Dim sqlstr = String.Format("select v.*,v.vendorcode::text || ' - ' || v.vendorname::text as vendordesc from vendor v order by {0}; {1}", SortField, vendorfamilysql)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function LoadDataVendorFamily(ByVal DS As DataSet) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select vf.id,vf.vendorcode,vf.familyid,vf.vendorcode || ' - ' || v.vendorname::text || (" &
                                    " case when vf.familyid isnull then ' '  else  ' - '  ||  vf.familyid ||  ' - ' || f.familyname  end )  || (case when mu.username isnull then ' ' else ' - ' || mu.username end ) as description from doc.vendorfamily vf" &
                                    " left join vendor v on v.vendorcode = vf.vendorcode left join family f on f.familyid = vf.familyid left join doc.familypm fpm on fpm.familyid = vf.familyid" &
                                    " left join officerseb o on o.ofsebid = fpm.pmid left join masteruser mu on mu.id = o.muid order by vf.vendorcode,vf.familyid")
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, "VendorFamily")
            myret = True
        End Using
        Return myret
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Return Nothing
    End Function


    Public Function GetShortNameBS() As BindingSource
        Dim MyDS As New DataSet
        Dim myBindingSource As New BindingSource
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select ''::text as shortname  union all (select distinct trim(shortname) as shortname from vendor where not shortname isnull order by shortname)")
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(MyDS, TableName)
            myret = True
        End Using
        myBindingSource.DataSource = MyDS.Tables(0)
        Return myBindingSource
    End Function

    Public Function GetShortNameUserBS(ByVal userid As String) As BindingSource
        Dim MyDS As New DataSet
        Dim myBindingSource As New BindingSource
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("with su as (select * from doc.getvendoruser('{0}')) select ''::text as shortname  union all  (select shortname from su order by shortname)", userid)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(MyDS, TableName)
            myret = True
        End Using
        myBindingSource.DataSource = MyDS.Tables(0)
        Return myBindingSource
    End Function

    Public Function GetVendorCodeBS() As BindingSource
        Dim MyDS As New DataSet
        Dim myBindingSource As New BindingSource
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select 0 as vendorcode,''::text as displayname,''::text as vendorname  union all (select vendorcode, vendorcode::text || ' - ' || trim(vendorname) as displayname,vendorname from vendor order by vendorname)")
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(MyDS, TableName)
            myret = True
        End Using
        myBindingSource.DataSource = MyDS.Tables(0)
        Return myBindingSource
    End Function

    Public Function GetVendorCodeShortnameBS() As BindingSource
        Dim MyDS As New DataSet
        Dim myBindingSource As New BindingSource
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select 0 as vendorcode,''::text as displayname,''::text as vendorname  union all (select vendorcode, vendorcode::text || ' - ' || trim(vendorname)  || (case when shortname isnull then '' else  ' - ' || shortname end) as displayname ,vendorname from vendor order by vendorname)")
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(MyDS, TableName)
            myret = True
        End Using
        myBindingSource.DataSource = MyDS.Tables(0)
        Return myBindingSource
    End Function

    Public Function GetVendorCodeUserBS(ByVal userid As String) As BindingSource
        Dim MyDS As New DataSet
        Dim myBindingSource As New BindingSource
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            'Dim sqlstr = String.Format("select 0 as vendorcode,''::text as displayname,''::text as vendorname  union all (select vendorcode, vendorcode::text || ' - ' || trim(vendorname) as displayname,vendorname from vendor order by vendorname)")
            Dim sqlstr = String.Format("with vu as (select * from doc.getvendoruser('{0}')) select 0 as vendorcode,''::text as displayname,''::text as vendorname  union all  (select vendorcode, description as displayname,vendorname from vu order by vendorname)", userid)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(MyDS, TableName)
            myret = True
        End Using
        myBindingSource.DataSource = MyDS.Tables(0)
        Return myBindingSource
    End Function

    Public Function GetVendorECOContactName(ByVal vendorcode As Long) As String
        Dim myret As String = String.Empty
        Dim sqlstr = String.Format("select c.contactname from doc.vendorcontact vc" &
                     " left join doc.contact c on c.id = vc.contactid" &
                     " where(isecoqualitycontact = True And vendorcode = {0});", vendorcode)
        myadapter.ExecuteScalar(sqlstr, myret)
        Return myret
    End Function

    Public Function GetVendorECOContactEmail(ByVal vendorcode As Long) As String
        Dim myret As String = String.Empty
        Dim sqlstr = String.Format("select c.email from doc.vendorcontact vc" &
                     " left join doc.contact c on c.id = vc.contactid" &
                     " where(isecoqualitycontact = True And vendorcode = {0});", vendorcode)
        myadapter.ExecuteScalar(sqlstr, myret)
        Return myret
    End Function
End Class
