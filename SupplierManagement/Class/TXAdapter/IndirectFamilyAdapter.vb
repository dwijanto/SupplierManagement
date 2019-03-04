Public Class IndirectFamilyAdapter
    Implements IAdapter
    Implements IToolbarAction

    Dim DS As DataSet
    Dim DbAdapter1 As DbAdapter = DbAdapter.getInstance
    Public bs As BindingSource
    Public Model As FamilyCodeModel

    Public ReadOnly Property GetTableFamily As DataTable
        Get
            Return DS.Tables("FamilyCode").Copy()
        End Get
    End Property
    'Public ReadOnly Property SortField As String
    '    Get
    '        Return "familycode"
    '    End Get
    'End Property

   
    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim BS As New BindingSource
            BS.DataSource = GetTableFamily
            BS.Sort = Model.SortField
            Return BS
        End Get
    End Property
    Public Function LoadData() As Boolean Implements IAdapter.loaddata
        Dim myret As Boolean = False
        Model = New FamilyCodeModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).PrimaryKey = pk
            bs = New BindingSource
            bs.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret
    End Function

    Public Function save() As Boolean Implements IAdapter.save
        Throw New System.Exception("Not implemented.")
    End Function

    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return bs.Filter
        End Get
        Set(ByVal value As String)
            bs.Filter = String.Format(Model.FilterField, value)
        End Set
    End Property

    Public Function GetCurrentRecord() As System.Data.DataRowView Implements IToolbarAction.GetCurrentRecord
        Return bs.Current
    End Function

    Public Function GetNewRecord() As System.Data.DataRowView Implements IToolbarAction.GetNewRecord
        Throw New System.Exception("Not implemented.")
    End Function

    Public Sub RemoveAt(ByVal value As Integer) Implements IToolbarAction.RemoveAt
        Throw New System.Exception("Not implemented.")
    End Sub

    Public Function Save1(ByVal mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Throw New System.Exception("Not implemented.")
    End Function
End Class
