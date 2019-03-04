Imports System.Windows.Forms.DataVisualization.Charting
Public Class ClassSEBTO
    Inherits UTODetail
    Implements IUTODetail

    Public Property currentDate As Date

    Private _SEBTOBS As BindingSource
    Private _SEBTOBS1 As BindingSource
    Private _SEBTOBS2 As BindingSource
    Private _SEBTOBS3 As BindingSource
    Private _SEBTOBS4 As BindingSource
    Private _sebtodrv As DataRowView
    Private _sebtodrv1 As DataRowView
    Private _sebtodrv2 As DataRowView
    Private _sebtodrv3 As DataRowView
    Private _sebtodrv4 As DataRowView
    Public Property SEBTOBS As BindingSource
        Get
            Return _SEBTOBS
        End Get
        Set(ByVal value As BindingSource)
            _SEBTOBS = value
            If Not IsNothing(value) Then
                _sebtodrv = value.Current
                validblankdrv(_sebtodrv)
            End If

        End Set
    End Property
    Public Property SEBTOBS1 As BindingSource
        Get
            Return _SEBTOBS1
        End Get
        Set(ByVal value As BindingSource)
            _SEBTOBS1 = value
            If Not IsNothing(value) Then
                _sebtodrv1 = value.Current
                validblankdrv(_sebtodrv1)
            End If

        End Set
    End Property
    Public Property SEBTOBS2 As BindingSource
        Get
            Return _SEBTOBS2
        End Get
        Set(ByVal value As BindingSource)
            _SEBTOBS2 = value
            If Not IsNothing(value) Then
                _sebtodrv2 = value.Current
                validblankdrv(_sebtodrv2)
            End If

        End Set
    End Property
    Public Property SEBTOBS3 As BindingSource
        Get
            Return _SEBTOBS3
        End Get
        Set(ByVal value As BindingSource)
            _SEBTOBS3 = value
            If Not IsNothing(value) Then
                _sebtodrv3 = value.Current
                validblankdrv(_sebtodrv3)
            End If

        End Set
    End Property
    Public Property SEBTOBS4 As BindingSource
        Get
            Return _SEBTOBS
        End Get
        Set(ByVal value As BindingSource)
            _SEBTOBS4 = value
            If Not IsNothing(value) Then
                _sebtodrv4 = value.Current
                validblankdrv(_sebtodrv4)
            End If

        End Set
    End Property
    Public Sub New()
        initLabel()
    End Sub
    'Public Sub ClearValue() Implements IUTODetail.ClearValue
    '    TextBox1.Text = 0
    '    TextBox2.Text = 0
    '    TextBox3.Text = 0
    '    TextBox4.Text = 0
    '    TextBox5.Text = 0
    '    TextBox6.Text = 0
    '    TextBox7.Text = 0
    '    TextBox8.Text = 0
    '    TextBox9.Text = 0
    '    TextBox10.Text = 0
    '    TextBox11.Text = "0%"
    '    TextBox12.Text = "0%"
    '    TextBox13.Text = "0%"
    '    TextBox14.Text = "0%"
    '    TextBox15.Text = "0%"
    '    TextBox16.Text = "0%"
    '    TextBox17.Text = "0%"
    '    TextBox18.Text = "0%"

    '    Dim yval(4) As Double
    '    Dim xval(4) As String
    '    For i = 0 To 4
    '        xval(i) = String.Format("{0}", Year(_currentDate) - i)
    '    Next
    '    For i = 0 To 4
    '        Chart1.Series(0).Points.Clear()
    '    Next

    'End Sub

    Public Overrides Sub DisplayValue() Implements IUTODetail.DisplayValue
        MyBase.displayvalue()
        If Not IsNothing(_sebtodrv) Then
            TextBox1.Text = "" & String.Format("{0:#,##0}", _sebtodrv.Row("qty"))
            TextBox10.Text = "" & String.Format("{0:#,##0}", _sebtodrv.Row("amount"))
        End If
        If Not IsNothing(_sebtodrv1) Then
            TextBox2.Text = "" & String.Format("{0:#,##0}", _sebtodrv1.Row("qty"))
            TextBox9.Text = "" & String.Format("{0:#,##0}", _sebtodrv1.Row("amount"))
        End If
        If Not IsNothing(_sebtodrv2) Then
            TextBox3.Text = "" & String.Format("{0:#,##0}", _sebtodrv2.Row("qty"))
            TextBox8.Text = "" & String.Format("{0:#,##0}", _sebtodrv2.Row("amount"))
        End If
        If Not IsNothing(_sebtodrv3) Then
            TextBox4.Text = "" & String.Format("{0:#,##0}", _sebtodrv3.Row("qty"))
            TextBox7.Text = "" & String.Format("{0:#,##0}", _sebtodrv3.Row("amount"))
        End If
        If Not IsNothing(_sebtodrv4) Then
            TextBox5.Text = "" & String.Format("{0:#,##0}", _sebtodrv4.Row("qty"))
            TextBox6.Text = "" & String.Format("{0:#,##0}", _sebtodrv4.Row("amount"))
        End If

        TextBox12.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox10.Text) / CDbl(TextBox9.Text) - 1))
        TextBox13.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox9.Text) / CDbl(TextBox8.Text) - 1))
        TextBox17.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox8.Text) / CDbl(TextBox7.Text) - 1))
        TextBox15.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox7.Text) / CDbl(TextBox6.Text) - 1))

        TextBox11.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox1.Text) / CDbl(TextBox2.Text) - 1))
        TextBox14.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox2.Text) / CDbl(TextBox3.Text) - 1))
        TextBox18.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox3.Text) / CDbl(TextBox4.Text) - 1))
        TextBox16.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox4.Text) / CDbl(TextBox5.Text) - 1))


        'Chart
        Chart1.Series(0).ChartType = SeriesChartType.Column
        Chart1.Series(0)("PointWidth") = "0.6"
        Chart1.Series(0).Name = ""
        Chart1.Series(0).IsValueShownAsLabel = False
        'Chart1.Series(0)("BarLabelStyle") = "Center"

        'Dim yval As Double() = {30000000, 40000000, 50000000, 10000000, 15000000}
        'Dim xval As String() = {"2014", "2013", "2012", "2011", "2010"}
        Dim yval(4) As Double
        Dim xval(4) As String
        For i = 0 To 4
            xval(i) = String.Format("{0}", Year(_currentDate) - i)
        Next
        yval(0) = CDbl(TextBox10.Text)
        yval(1) = CDbl(TextBox9.Text)
        yval(2) = CDbl(TextBox8.Text)
        yval(3) = CDbl(TextBox7.Text)
        yval(4) = CDbl(TextBox6.Text)
        Chart1.Series(0).Points.DataBindXY(xval, yval)

    End Sub

    Public Sub initLabel() Implements IUTODetail.initLabel
        Label2.Text = "Turnover Total with SEB Asia"
        Label1.Visible = False
        Label23.Text = "QTY"
        Label24.Text = "Purchase Amount (USD)"

        TextBox19.Visible = False
    End Sub
    Private Sub validblankdrv(ByVal sebtodrv As DataRowView)
        If Not IsNothing(sebtodrv) Then
            If IsDBNull(sebtodrv.Row.Item("amount")) Then sebtodrv.Row.Item("amount") = 0
            If IsDBNull(sebtodrv.Row.Item("qty")) Then sebtodrv.Row.Item("qty") = 0
            If IsDBNull(sebtodrv.Row.Item("piavg")) Then sebtodrv.Row.Item("piavg") = 0
            If IsDBNull(sebtodrv.Row.Item("pilkp")) Then sebtodrv.Row.Item("pilkp") = 0
            If IsDBNull(sebtodrv.Row.Item("pistd")) Then sebtodrv.Row.Item("pistd") = 0
            If IsDBNull(sebtodrv.Row.Item("tovariance")) Then sebtodrv.Row.Item("tovariance") = 0
        End If

    End Sub
End Class
