Imports System.Text
Public Class PaymentTermAdapter
    Implements IAdapter
    Implements IToolbarAction
    Dim DS As DataSet
    Dim DbAdapter1 As DbAdapter = DbAdapter.getInstance
    Public bs As BindingSource

    Dim PaymentTermModel1 As PaymentTermModel

    Public ReadOnly Property GetTable As DataTable
        Get
            Return DS.Tables("PaymentTerm").Copy()
        End Get
    End Property

    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim BS As New BindingSource
            BS.DataSource = GetTable
            BS.Sort = PaymentTermModel1.SortField
            Return BS
        End Get
    End Property
    Public Function loaddata() As Boolean Implements IAdapter.loaddata
        Dim myret As Boolean = False
        PaymentTermModel1 = New PaymentTermModel
        DS = New DataSet
        If PaymentTermModel1.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("paymmenttermid")
            DS.Tables(0).PrimaryKey = pk
            bs = New BindingSource
            bs.DataSource = DS.Tables(PaymentTermModel1.TableName)
            myret = True
        End If
        Return myret
    End Function

    Public Function save() As Boolean Implements IAdapter.save
        Dim myret As Boolean = False
        bs.EndEdit()

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
        Dim PaymentTermModel1 As New PaymentTermModel
        If PaymentTermModel1.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function

    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            bs.Filter = String.Format("[payt] like '*{0}*' or [details] like '*{0}*' or [days] like '*{0}*'", value)
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
