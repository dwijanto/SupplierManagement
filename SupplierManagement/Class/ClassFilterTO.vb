Imports System.Windows.Forms.DataVisualization.Charting
Public Class ClassFilterTO
    Inherits UTODetail
    Implements IUTODetail

    Dim _qtybs As BindingSource
    Dim _qtydrv As Object
    Dim _amountbs As BindingSource
    Dim _amountdrv As Object
    Dim _displayFilter As String

    Public Property currentdate As Date

    Public Property DisplayFilter As String
        Get
            Return _displayFilter
        End Get
        Set(ByVal value As String)
            _displayFilter = value
            TextBox19.Text = value
        End Set
    End Property

    Public Property QTYBS As BindingSource
        Get
            Return _qtybs
        End Get
        Set(ByVal value As BindingSource)
            _qtybs = value
            If Not IsNothing(value) Then
                Try
                    _qtydrv = value.Current
                    validblanksdrv(_qtydrv)
                Catch ex As Exception

                End Try
                
            End If
        End Set
    End Property
    Public Property AmountBS As BindingSource
        Get
            Return _amountbs
        End Get
        Set(ByVal value As BindingSource)
            _amountbs = value
            If Not IsNothing(value) Then
                Try
                    _amountdrv = value.Current
                    validblanksdrv(_amountdrv)
                Catch ex As Exception

                End Try
                
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
    '        xval(i) = String.Format("{0}", Year(_currentdate) - i)
    '    Next
    '    For i = 0 To 4
    '        Chart1.Series(0).Points.Clear()
    '    Next
    'End Sub

    Public Overrides Sub DisplayValue() Implements IUTODetail.DisplayValue
        MyBase.displayvalue()
        If Not IsNothing(QTYBS.Current) Then
            TextBox1.Text = "" & String.Format("{0:#,##0}", _qtydrv.item(1))
            TextBox2.Text = "" & String.Format("{0:#,##0}", _qtydrv.item(2))
            TextBox3.Text = "" & String.Format("{0:#,##0}", _qtydrv.item(3))
            TextBox4.Text = "" & String.Format("{0:#,##0}", _qtydrv.item(4))
            TextBox5.Text = "" & String.Format("{0:#,##0}", _qtydrv.item(5))
        End If
        If Not IsNothing(AmountBS.Current) Then
            TextBox10.Text = "" & String.Format("{0:#,##0}", _amountdrv.item(1))
            TextBox9.Text = "" & String.Format("{0:#,##0}", _amountdrv.item(2))
            TextBox8.Text = "" & String.Format("{0:#,##0}", _amountdrv.item(3))
            TextBox7.Text = "" & String.Format("{0:#,##0}", _amountdrv.item(4))
            TextBox6.Text = "" & String.Format("{0:#,##0}", _amountdrv.item(5))
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
            xval(i) = String.Format("{0}", Year(_currentdate) - i)
        Next
        yval(0) = CDbl(TextBox10.Text)
        yval(1) = CDbl(TextBox9.Text)
        yval(2) = CDbl(TextBox8.Text)
        yval(3) = CDbl(TextBox7.Text)
        yval(4) = CDbl(TextBox6.Text)
        Chart1.Series(0).Points.DataBindXY(xval, yval)
    End Sub

    Public Sub initLabel() Implements IUTODetail.initLabel
        Label2.Text = "By Filter"
        Label24.Text = "Purchase Amount (USD)"
        Label23.Text = "QTY"
        Label1.Visible = True
        TextBox19.Visible = True
    End Sub

    Private Sub validblanksdrv(ByVal drv As Object)
        For i = 1 To 5
            If IsDBNull(drv.item(i)) Then drv.item(i) = 0
        Next
    End Sub

End Class
