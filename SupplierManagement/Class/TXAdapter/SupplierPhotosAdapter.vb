Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class SupplierPhotosAdapter
    Private sb As StringBuilder
    Public DS As DataSet
    Public BS As BindingSource
    Private shortname As String
    Private DRV As DataRowView
    Public PhotoFolder As String

    Public Sub New(ByVal drv As DataRowView)
        Me.drv = drv
    End Sub
    Public Sub New()
       
    End Sub
    Public Function LoadSupplierPhoto() As Boolean
        DS = New DataSet
        BS = New BindingSource
        sb = New StringBuilder

        sb.Clear()
        sb.Append(String.Format("select vp.*,case phototype when 1 then 'Factory' else 'Product' end as phototypename,null::text as fullpath from doc.vendorphoto vp " &
                                " left join vendor v on v.vendorcode = vp.vendorcode" &
                                " where v.shortname = '{0}' order by id;", DRV.Row.Item("shortname")))
        sb.Append("select cvalue from doc.paramhd where paramname = 'supplierphotos';")
        If DbAdapter1.TbgetDataSet(sb.ToString, DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).PrimaryKey = pk
            DS.Tables(0).Columns("id").AutoIncrement = True
            DS.Tables(0).Columns("id").AutoIncrementSeed = 0
            DS.Tables(0).Columns("id").AutoIncrementStep = -1

            BS.DataSource = DS.Tables(0)
            PhotoFolder = DS.Tables(1).Rows(0).Item(0)

        End If
        Return True
    End Function

    Public Function GetSupplierPhoto(ByVal shortname As String, ByVal producttype As String) As Boolean
        DS = New DataSet
        BS = New BindingSource
        sb = New StringBuilder
        Dim myret As Boolean = False
        sb.Clear()
        sb.Append(String.Format("select vp.*,case phototype when 1 then 'Factory' else 'Product' end as phototypename,null::text as fullpath from doc.vendorphoto vp " &
                                " left join vendor v on v.vendorcode = vp.vendorcode" &
                                " where v.shortname = '{0}' and {1} order by phototype,lineorder{1},createddate desc,description;", shortname, producttype))
        sb.Append("select cvalue from doc.paramhd where paramname = 'supplierphotos';")
        If DbAdapter1.TbgetDataSet(sb.ToString, DS) Then
            Dim pk(0) As DataColumn
            'pk(0) = DS.Tables(0).Columns("id")
            'DS.Tables(0).PrimaryKey = pk
            BS.DataSource = DS.Tables(0)
            PhotoFolder = DS.Tables(1).Rows(0).Item(0)
            myret = True
        End If
        Return myret
    End Function

    Public Function save() As Boolean
        Dim myret As Boolean = False
        'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
        Dim ds2 As DataSet
        ds2 = DS.GetChanges

        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If Not DbAdapter1.SupplierPhotosTx(Me, mye) Then
                MessageBox.Show(mye.message)
                Return myret
            End If
            DS.Merge(ds2)
            DS.AcceptChanges()
            MessageBox.Show("Saved.")
            myret = True
        End If
        Return myret
    End Function
End Class
