Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass

Public Class VendorAdapter
    Private sb As StringBuilder
    Public DS As DataSet
    Public BS As BindingSource
    Public Property shortname

    Public Function LoadSupplierSearchPhoto() As Boolean
        DS = New DataSet
        BS = New BindingSource
        sb = New StringBuilder

        sb.Clear()
        'sb.Append(String.Format("select vendorcode::text,vendorname::text,shortname::text from vendor v order by vendorcode;"))
        sb.Append(String.Format("with s as (select v.shortname3 as shortname,vp.*,case when fp then 1 else 0 end as fpint,case when cp then 1 else 0 end as cpint,case when general then 1 else 0 end as generalint from doc.vendorphoto vp left join vendor v on v.vendorcode = vp.vendorcode)" &
                                " select v.vendorcode::text,vendorname::text,v.shortname3::text as shortname,sum(fpint) as countfp,sum(cpint) as countcp,sum(generalint) as countgeneral from vendor v " &
                                " left join s on s.shortname = v.shortname3" &
                                " group by v.vendorcode,vendorname,v.shortname3" &
                                " order by v.vendorcode;"))
        If DbAdapter1.TbgetDataSet(sb.ToString, DS) Then
            Dim pk(0) As DataColumn
            DS.Tables(0).PrimaryKey = pk
            BS.DataSource = DS.Tables(0)
        End If
        Return True
    End Function
End Class
