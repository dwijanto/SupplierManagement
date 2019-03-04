Public Class GroupDataTypeAdapter
    Inherits BaseAdapter
    Implements IAdapter

    Public Sub New()
        MyBase.new()
    End Sub

    Public Function loaddata() As Boolean Implements IAdapter.loaddata
        Return True
    End Function

    Public Function save() As Boolean Implements IAdapter.save
        Return True
    End Function

    Public Function getComboBoxBS() As BindingSource
        Dim sqlstr = "select pdt.ivalue as id,pdt.paramname as groupname from doc.paramdt pdt left join doc.paramhd phd on phd.paramhdid = pdt.paramhdid where phd.paramname = 'groupdatatype';"
        Dim ds As New DataSet
        Dim bs As New BindingSource
        If DbAdapter1.TbgetDataSet(sqlstr, ds) Then
            bs.DataSource = ds.Tables(0)
        End If
        Return bs
    End Function
End Class
