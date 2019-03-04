Imports System.Text

Public Class ClassSelectQuery
    Public Property WithQuery As String
        Get
            Return _withQuerySB.ToString
        End Get
        Set(ByVal value As String)
            If _withQuerySB.Length = 0 Then
                _withQuerySB.Append("With ")
            End If
            _withQuerySB.Append(value)
        End Set
    End Property

    Public Property SelectFrom As String
        Get
            Return _selectFromSB.ToString
        End Get
        Set(ByVal value As String)
            _selectFromSB.Append(value)
        End Set
    End Property

    Public Property JoinCondition As String
        Get
            Return _joinConditionSB.ToString
        End Get
        Set(ByVal value As String)
            _joinConditionSB.Append(value)
        End Set
    End Property

    Public Property WhereCondition As String
        Get
            Return _whereConditionSB.ToString
        End Get
        Set(ByVal value As String)
            _whereConditionSB.Append(value)
        End Set
    End Property

    Public Property GroupBy As String
        Get
            Return _groupbySB.ToString
        End Get
        Set(ByVal value As String)
            _groupbySB.Append(value)
        End Set
    End Property

    Public Property HavingCondition As String
        Get
            Return _havingConditionSB.ToString
        End Get
        Set(ByVal value As String)
            _havingConditionSB.Append(value)
        End Set
    End Property

    Public Property Orderby As String
        Get
            Return _orderbySB.ToString
        End Get
        Set(ByVal value As String)
            _orderbySB.Append(value)
        End Set
    End Property

    Public Property Limit As String
        Get
            Return _limitSB.ToString
        End Get
        Set(ByVal value As String)
            _limitSB.Append(value)
        End Set
    End Property

    Public Property Offset As String
        Get
            Return _offsetSB.ToString
        End Get
        Set(ByVal value As String)
            _offsetSB.Append(value)
        End Set
    End Property

    Public Property Fetch As String
        Get
            Return _fetchSB.ToString
        End Get
        Set(ByVal value As String)
            _fetchSB.Append(value)
        End Set
    End Property
    Public Property ForCondition As String
        Get
            Return _forConditionSB.ToString
        End Get
        Set(ByVal value As String)
            _forConditionSB.Append(value)
        End Set
    End Property

    Private _withQuerySB As StringBuilder
    Public Property _selectFromSB As StringBuilder
    Public Property _joinConditionSB As StringBuilder
    Public Property _whereConditionSB As StringBuilder
    Public Property _groupbySB As StringBuilder
    Public Property _havingConditionSB As StringBuilder
    Public Property _orderbySB As StringBuilder
    Public Property _limitSB As StringBuilder
    Public Property _offsetSB As StringBuilder
    Public Property _fetchSB As StringBuilder
    Public Property _forConditionSB As StringBuilder

    Public Sub New()
        _withQuerySB = New StringBuilder
        _selectFromSB = New StringBuilder
        _joinConditionSB = New StringBuilder
        _whereConditionSB = New StringBuilder
        _groupbySB = New StringBuilder
        _havingConditionSB = New StringBuilder
        _orderbySB = New StringBuilder
        _limitSB = New StringBuilder
        _offsetSB = New StringBuilder
        _fetchSB = New StringBuilder
        _forConditionSB = New StringBuilder
    End Sub

    Public Function getQuery() As String
        Return (String.Format("{0} {1} {2} {3} {4} {5} {6} {7} {8} {9} {10}", WithQuery, SelectFrom, JoinCondition, WhereCondition, GroupBy,
            HavingCondition, Orderby, Limit, Offset, Fetch, ForCondition))
    End Function


End Class
