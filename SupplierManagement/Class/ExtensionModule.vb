

<AttributeUsage(AttributeTargets.Enum Or AttributeTargets.Field, AllowMultiple:=False)> _
Public NotInheritable Class EnumDescriptionAttribute
    Inherits Attribute
    Private _description As String
    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property
    Public Sub New(ByVal description As String)
        MyBase.New()
        _description = description
    End Sub
End Class

Public Module Extension
    <System.Runtime.CompilerServices.Extension()> _
    Public Function EnumToDictionary(ByVal t As Type) As IDictionary(Of String, Integer)
        If t Is Nothing Then
            Throw New NullReferenceException()
        End If
        If Not t.IsEnum Then
            Throw New InvalidCastException("object is not an Enumeration")
        End If

        Dim names As String() = [Enum].GetNames(t)
        Dim values As Array = [Enum].GetValues(t)

        Dim mydict As Dictionary(Of String, Integer)
        mydict = (From i In Enumerable.Range(0, names.Length)
                Select New With {Key .Key = GetDescription(values(i)),
                                 Key .Value = CInt(values.GetValue(i))}).ToDictionary(Function(x) x.Key, Function(x) x.Value)


        'Return (From i In Enumerable.Range(0, names.Length)
        '        Select New With {
        '            Key .Key = GetDescription(values(i)),
        '            Key .Value = CInt(values.GetValue(i))}
        '        ).ToDictionary(Function(k) k.Key, Function(k) k.Value)
        Return mydict
    End Function

    Public Function GetDescription(ByVal value As [Enum]) As String
        If IsNothing(value) Then Throw New ArgumentNullException("value")
        Dim description As String = value.ToString()
        Dim fieldInfo As Reflection.FieldInfo = value.GetType().GetField(description)
        Dim attributes() As EnumDescriptionAttribute = CType(fieldInfo.GetCustomAttributes(
                GetType(EnumDescriptionAttribute), False), EnumDescriptionAttribute())
        If Not IsNothing(attributes) AndAlso attributes.Length > 0 Then description = attributes(0).Description
        Return description
    End Function
End Module
