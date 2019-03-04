Public Class VendorTxAdapter
    Implements IAdapter

    Dim DS As DataSet
    Dim DbAdapter1 As DbAdapter = DbAdapter.getInstance
    Public bs As BindingSource

    Dim VendorModel1 As New VendorModel

    Public ReadOnly Property GetTableVendor As DataTable
        Get
            Return DS.Tables("Vendor").Copy()
        End Get
    End Property

    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim myBS As New BindingSource
            myBS.DataSource = GetTableVendor
            myBS.Sort = SortField
            Return myBS
        End Get
    End Property
    Public ReadOnly Property GetBindingSourceVendorFamily As BindingSource
        Get
            Dim myBS As New BindingSource

            myBS.DataSource = DS.Tables(1).Copy()
            myBS.Sort = "vendorcode,familyid"
            Return myBS
        End Get
    End Property

    Public ReadOnly Property SortField
        Get
            Return VendorModel1.SortField
        End Get
    End Property

    Public Function loaddata() As Boolean Implements IAdapter.loaddata
        Dim myret As Boolean = False
        'VendorModel1 = New VendorModel
        DS = New DataSet
        If VendorModel1.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("vendorcode")
            DS.Tables(0).PrimaryKey = pk
            bs = New BindingSource
            bs.DataSource = DS.Tables(VendorModel1.TableName)
            myret = True
        End If
        Return myret
    End Function

    Public Function save() As Boolean Implements IAdapter.save
        Return True
    End Function

    Public Function GetVendorEcoContactName(ByVal vendorcode As Long) As String
        Return VendorModel1.GetVendorECOContactName(vendorcode)
    End Function
    Public Function GetVendorEcoContactEmail(ByVal vendorcode As Long) As String
        Return VendorModel1.GetVendorECOContactEmail(vendorcode)
    End Function


End Class
