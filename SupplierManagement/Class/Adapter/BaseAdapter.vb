Public Class BaseAdapter
    Public DS As DataSet
    Public BS As BindingSource
    Public DbAdapter1 As DbAdapter

    Public Sub New()
        DbAdapter1 = DbAdapter.getInstance
    End Sub

End Class
