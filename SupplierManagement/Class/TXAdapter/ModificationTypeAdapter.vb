Public Class ModificationTypeAdapter
    Implements IAdapter
    Dim myModel As New ModificationTypeModel
    Public Function loaddata() As Boolean Implements IAdapter.loaddata
        Return Nothing
    End Function

    Public Function save() As Boolean Implements IAdapter.save
        Return Nothing
    End Function

    Public Function GetModificationTypeBS() As BindingSource
        Return myModel.GetModificationTypeBS
    End Function
End Class
