Imports System.Text

Public Class PMAdapter
    Implements IAdapter

    Dim DS As DataSet
    Dim DbAdapter1 As DbAdapter = DbAdapter.getInstance
    Public bs As BindingSource

    Dim PMModel1 As New PMModel

    Public ReadOnly Property GetTablePM As DataTable
        Get
            Return DS.Tables("PM").Copy()
        End Get
    End Property

    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim BS As New BindingSource
            BS.DataSource = GetTablePM
            BS.Sort = PMModel1.SortField
            Return BS
        End Get
    End Property

    Public ReadOnly Property getPMBS As BindingSource
        Get
            Return PMModel1.getPMBS
        End Get
    End Property

    Public Function loaddata() As Boolean Implements IAdapter.loaddata
        Dim myret As Boolean = False
        PMModel1 = New PMModel
        DS = New DataSet
        If PMModel1.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("pmid")
            DS.Tables(0).PrimaryKey = pk
            bs = New BindingSource
            bs.DataSource = DS.Tables(PMModel1.TableName)
            myret = True
        End If
        Return myret
    End Function

    Public Function save() As Boolean Implements IAdapter.save
        Return True
    End Function
End Class
