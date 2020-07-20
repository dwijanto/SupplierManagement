Imports System.Text

Public Class FamilyAdapter
    Implements IAdapter

    Dim DS As DataSet
    Dim DbAdapter1 As DbAdapter = DbAdapter.getInstance
    Public bs As BindingSource
    Dim FamilyModel1 As New FamilyModel

    Public ReadOnly Property GetTableFamily As DataTable
        Get
            Return DS.Tables("Family").Copy()
        End Get
    End Property
    Public ReadOnly Property SortField As String
        Get
            Return "familyid"
        End Get
    End Property
    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim BS As New BindingSource
            BS.DataSource = GetTableFamily
            BS.Sort = FamilyModel1.SortField
            Return BS
        End Get
    End Property

    Public ReadOnly Property getFamilyBS As BindingSource
        Get
            Return FamilyModel1.getFamilyBS
        End Get
    End Property

    Public Function LoadData() As Boolean Implements IAdapter.loaddata
        Dim myret As Boolean = False
        FamilyModel1 = New FamilyModel
        DS = New DataSet
        If FamilyModel1.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("familyid")
            DS.Tables(0).PrimaryKey = pk
            bs = New BindingSource
            bs.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret
    End Function
    'Public Function loaddata() As Boolean Implements IAdapter.loaddata
    '    Dim myret As Boolean = False
    '    Dim sb As New StringBuilder
    '    sb.Append("select familyid,familyname::text ,familyid::text || ' - ' || familyname::text as familydesc from family order by familyid;")
    '    Dim sqlstr = sb.ToString
    '    DS = New DataSet
    '    If DbAdapter1.TbgetDataSet(sqlstr, DS) Then
    '        'set primary key
    '        DS.Tables(0).TableName = "Family"
    '        Dim pk(0) As DataColumn
    '        pk(0) = DS.Tables(0).Columns("familyid")
    '        DS.Tables(0).PrimaryKey = pk
    '        bs = New BindingSource
    '        bs.DataSource = DS.Tables(0)
    '        myret = True
    '    End If
    '    Return myret
    'End Function

    Public Function save() As Boolean Implements IAdapter.save
        Return True
    End Function
End Class
