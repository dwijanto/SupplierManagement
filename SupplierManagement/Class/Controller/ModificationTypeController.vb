Public Class ModificationTypeController
    Implements IController
    Implements IToolbarAction

    Public Model As New ModificationTypeModel
    Public BS As BindingSource
    Public BSDTL As BindingSource
    Public DocTypeBS As BindingSource
    Dim DS As DataSet
    Public ReadOnly Property GetTable As System.Data.DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.TableName).Copy()
        End Get
    End Property

    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New ModificationTypeModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).PrimaryKey = pk
            DS.Tables(0).Columns("id").AutoIncrement = True
            DS.Tables(0).Columns("id").AutoIncrementSeed = -1
            DS.Tables(0).Columns("id").AutoIncrementStep = -1

            Dim pk1(0) As DataColumn
            pk1(0) = DS.Tables(1).Columns("id")
            DS.Tables(1).PrimaryKey = pk1
            DS.Tables(1).Columns("id").AutoIncrement = True
            DS.Tables(1).Columns("id").AutoIncrementSeed = -1
            DS.Tables(1).Columns("id").AutoIncrementStep = -1

            'DataRelation 
            Dim rel As DataRelation
            Dim hcol As DataColumn
            Dim dcol As DataColumn
            hcol = DS.Tables(0).Columns("id")
            dcol = DS.Tables(1).Columns("modificationtypeid")
            rel = New DataRelation("hdrel", hcol, dcol)
            DS.Relations.Add(rel)

            'Bindingsource
            BS = New BindingSource
            BSDTL = New BindingSource
            DocTypeBS = New BindingSource

            BS.DataSource = DS.Tables(0)
            BSDTL.DataSource = BS
            BSDTL.DataMember = "hdrel"
            DocTypeBS.DataSource = DS.Tables(2)
            myret = True
        End If
        Return myret
    End Function

    Public Function save() As Boolean Implements IController.save
        Dim myret As Boolean = False
        BS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If

        Return myret
    End Function

    Public Function Save(ByVal mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Dim myret As Boolean = False
        If Model.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function
    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format(Model.FilterField, value)
        End Set
    End Property

    Public Function GetCurrentRecord() As System.Data.DataRowView Implements IToolbarAction.GetCurrentRecord
        Return BS.Current
    End Function

    Public Function GetNewRecord() As System.Data.DataRowView Implements IToolbarAction.GetNewRecord
        Return BS.AddNew
    End Function

    Public Sub RemoveAt(ByVal value As Integer) Implements IToolbarAction.RemoveAt
        BS.RemoveAt(value)
    End Sub


End Class
