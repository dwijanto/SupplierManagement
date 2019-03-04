Imports System.Text

Public Class DataTypeLabelAdapter
    Inherits BaseAdapter
    Implements IAdapter

    Private sb As StringBuilder


    Public Sub New()
        MyBase.New()
    End Sub

    Public Function loaddata() As Boolean Implements IAdapter.loaddata
        DS = New DataSet
        BS = New BindingSource
        sb = New StringBuilder

        sb.Clear()
        sb.Append("select * from doc.datatypelabel")
        If DbAdapter1.TbgetDataSet(sb.ToString, DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).PrimaryKey = pk
            DS.Tables(0).Columns("id").AutoIncrement = True
            DS.Tables(0).Columns("id").AutoIncrementSeed = 0
            DS.Tables(0).Columns("id").AutoIncrementStep = -1

            BS.DataSource = DS.Tables(0)
        End If
        Return True
    End Function


    Public Function save() As Boolean Implements IAdapter.save
        Dim myret As Boolean = False
        Dim ds2 As DataSet
        ds2 = DS.GetChanges

        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If Not DbAdapter1.DataTypeLabelTx(Me, mye) Then
                DS.Merge(ds2)
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
