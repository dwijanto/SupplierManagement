Imports System.Reflection 'for FieldInfo

Public Class EnumHelper
    Public Shared Function GetDescription(ByVal value As [Enum]) As String
        If IsNothing(value) Then Throw New ArgumentNullException("value")
        Dim description As String = value.ToString()
        Dim fieldInfo As FieldInfo = value.GetType().GetField(description)
        Dim attributes() As EnumDescriptionAttribute = _
        CType(fieldInfo.GetCustomAttributes(GetType(EnumDescriptionAttribute), False),  _
        EnumDescriptionAttribute())
        If Not IsNothing(attributes) AndAlso attributes.Length > 0 Then _
            description = attributes(0).Description
        Return description
    End Function
    Public Shared Function ToList(ByVal type As Type) As IList
        If IsNothing(type) Then Throw New ArgumentNullException("type")
        Dim list As New ArrayList
        Dim enumValues As Array = [Enum].GetValues(type)
        For Each value As [Enum] In enumValues
            list.Add(New KeyValuePair(Of [Enum], String)(value, GetDescription(value)))
        Next
        Return list
    End Function
    Public Shared Function ToDict(ByVal type As [Enum]) As IDictionary
        If IsNothing(type) Then Throw New ArgumentNullException("type")
        Dim mydict As New Dictionary(Of String, Integer)

        For Each value In [Enum].GetValues(GetType(Type))
            'list.Add([Enum].GetNames(GetType(value)), [Enum].GetValues(type))
            mydict.Add(value.ToString, DirectCast(value, Integer))
        Next
        Return mydict
    End Function

End Class