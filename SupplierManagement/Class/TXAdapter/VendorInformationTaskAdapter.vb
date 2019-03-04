Public Class VendorInformationTaskAdapter
    Dim myModel As VendorInformationModificationModel = New VendorInformationModificationModel
    Public bs1 As BindingSource
    Public bs2 As BindingSource
    Public ReadOnly Property getCurrentRecordTx
        Get
            Return bs1.Current
        End Get
    End Property
    Public ReadOnly Property getCurrentRecordHistory
        Get
            Return bs2.Current
        End Get
    End Property

    Public Function loadData(ByRef ds As DataSet, ByVal userid As String, ByVal criteria As String) As Boolean
        Dim myret As Boolean = False
        myret = myModel.LoadUserTask(ds, userid, criteria)
        If myret Then
            initbindingsource(ds)

        End If
        Return myret
    End Function
    Public Function loadData(ByRef ds As DataSet, ByVal criteria As String) As Boolean
        Dim myret As Boolean = False
        myret = myModel.LoadUserTask(ds, criteria)
        If myret Then
            initbindingsource(ds)
        End If
        Return myret
    End Function

    Private Sub initbindingsource(ByVal ds As DataSet)
        bs1 = New BindingSource
        bs1.DataSource = ds.Tables(0)
        bs2 = New BindingSource
        bs2.DataSource = ds.Tables(1)
    End Sub

End Class
