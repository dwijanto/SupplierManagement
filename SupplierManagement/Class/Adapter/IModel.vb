Public Interface IModel
    'Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean
    ReadOnly Property tablename As String
    ReadOnly Property sortField As String
    Function LoadData(ByVal DS As DataSet) As Boolean
    Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean
End Interface
