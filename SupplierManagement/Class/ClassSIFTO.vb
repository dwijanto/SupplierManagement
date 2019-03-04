Imports System.Windows.Forms.DataVisualization.Charting

Public Class ClassSIFTO
    Inherits UTODetail
    Implements IUTODetail

    'Private _currentDate As Date
    Private _SIFTOBS As BindingSource
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

    Public Property currentDate As Date
    Public Property SIFTOBS As BindingSource
        Get
            Return _SIFTOBS
        End Get
        Set(ByVal value As BindingSource)
            _SIFTOBS = value
            validblank(_SIFTOBS)
        End Set
    End Property

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

    'TextBox1.Text = 0
    'TextBox2.Text = 0
    'TextBox3.Text = 0
    'TextBox4.Text = 0
    'TextBox5.Text = 0

    'TextBox11.Text = "0%"
    'TextBox14.Text = "0%"
    'TextBox18.Text = "0%"
    'TextBox16.Text = "0%"

    'TextBox10.Text = "0%"
    'TextBox9.Text = "0%"
    'TextBox8.Text = "0%"
    'TextBox7.Text = "0%"
    'TextBox6.Text = "0%"

    'Dim yval(4) As Double
    'Dim xval(4) As String
    'For i = 0 To 4
    '    xval(i) = String.Format("{0}", Year(_currentDate) - i)
    'Next
    'For i = 0 To 4
    '    Chart1.Series(0).Points.Clear()
    'Next

    'End Sub
    Public Overrides Sub DisplayValue() Implements IUTODetail.DisplayValue
        MyBase.displayvalue()
        If Not IsNothing(_SIFTOBS.Current) Then
            Dim drv1 As DataRowView = _SIFTOBS.Current
            Dim checkYear As Integer = Year(_currentDate) - drv1.Row("myyear")
            ClearValue()
            Select Case checkYear
                Case 0
                    TextBox1.Text = String.Format("{0:#,##0}", drv1.Row("turnovery"))
                    TextBox2.Text = String.Format("{0:#,##0}", drv1.Row("turnovery1"))
                    TextBox3.Text = String.Format("{0:#,##0}", drv1.Row("turnovery2"))
                    TextBox4.Text = String.Format("{0:#,##0}", drv1.Row("turnovery3"))
                    TextBox5.Text = String.Format("{0:#,##0}", drv1.Row("turnovery4"))

                    If drv1.Row("turnovery1") <> 0 Then TextBox11.Text = ValidText(String.Format("{0:0%}", drv1.Row("turnovery") / drv1.Row("turnovery1") - 1))
                    If drv1.Row("turnovery2") <> 0 Then TextBox14.Text = ValidText(String.Format("{0:0%}", drv1.Row("turnovery1") / drv1.Row("turnovery2") - 1))
                    If drv1.Row("turnovery3") <> 0 Then TextBox18.Text = ValidText(String.Format("{0:0%}", drv1.Row("turnovery2") / drv1.Row("turnovery3") - 1))
                    If drv1.Row("turnovery4") <> 0 Then TextBox16.Text = ValidText(String.Format("{0:0%}", drv1.Row("turnovery3") / drv1.Row("turnovery4") - 1))
                Case 1
                    TextBox2.Text = String.Format("{0:#,##0}", drv1.Row("turnovery"))
                    TextBox3.Text = String.Format("{0:#,##0}", drv1.Row("turnovery1"))
                    TextBox4.Text = String.Format("{0:#,##0}", drv1.Row("turnovery2"))
                    TextBox5.Text = String.Format("{0:#,##0}", drv1.Row("turnovery3"))
                    If drv1.Row("turnovery") <> 0 Then TextBox11.Text = ValidText(String.Format("{0:0%}", 0 / drv1.Row("turnovery") - 1))
                    If (drv1.Row("turnovery1") <> 0) Then TextBox14.Text = ValidText(String.Format("{0:0%}", drv1.Row("turnovery") / drv1.Row("turnovery1") - 1))
                    If (drv1.Row("turnovery2") <> 0) Then TextBox18.Text = ValidText(String.Format("{0:0%}", drv1.Row("turnovery1") / drv1.Row("turnovery2") - 1))
                    If (drv1.Row("turnovery3") <> 0) Then TextBox16.Text = ValidText(String.Format("{0:0%}", drv1.Row("turnovery2") / drv1.Row("turnovery3") - 1))
                Case 2
                    TextBox3.Text = String.Format("{0:#,##0}", drv1.Row("turnovery"))
                    TextBox4.Text = String.Format("{0:#,##0}", drv1.Row("turnovery1"))
                    TextBox5.Text = String.Format("{0:#,##0}", drv1.Row("turnovery2"))
                    TextBox11.Text = "0%"
                    If drv1.Row("turnovery") <> 0 Then TextBox14.Text = ValidText(String.Format("{0:0%}", 0 / drv1.Row("turnovery") - 1))
                    If drv1.Row("turnovery1") <> 0 Then TextBox18.Text = ValidText(String.Format("{0:0%}", drv1.Row("turnovery") / drv1.Row("turnovery1") - 1))
                    If drv1.Row("turnovery2") <> 0 Then TextBox16.Text = ValidText(String.Format("{0:0%}", drv1.Row("turnovery1") / drv1.Row("turnovery2") - 1))
                Case 3
                    TextBox4.Text = String.Format("{0:#,##0}", drv1.Row("turnovery"))
                    TextBox5.Text = String.Format("{0:#,##0}", drv1.Row("turnovery1"))
                    TextBox11.Text = "0%"
                    TextBox14.Text = "0%"
                    If drv1.Row("turnovery") <> 0 Then TextBox18.Text = ValidText(String.Format("{0:0%}", 0 / drv1.Row("turnovery") - 1))
                    If drv1.Row("turnovery1") <> 0 Then TextBox16.Text = ValidText(String.Format("{0:0%}", drv1.Row("turnovery") / drv1.Row("turnovery1") - 1))
                Case 4
                    TextBox5.Text = String.Format("{0:#,##0}", drv1.Row("turnovery"))
                    TextBox11.Text = "0%"
                    TextBox14.Text = "0%"
                    TextBox18.Text = "0%"
                    If drv1.Row("turnovery") <> 0 Then TextBox16.Text = ValidText(String.Format("{0:0%}", 0 / drv1.Row("turnovery") - 1))
            End Select
            If Not IsNothing(_sebtodrv) Then
                'If drv1.Row("turnovery") <> 0 Then
                If TextBox1.Text <> 0 Then
                    'TextBox10.Text = ValidText(String.Format("{0:0.0%}", IIf(drv1.Row("turnovery") = 0, 0, _sebtodrv.Row.Item("amount") / drv1.Row("turnovery"))))
                    TextBox10.Text = ValidText(String.Format("{0:0.0%}", _sebtodrv.Row.Item("amount") / CDbl(TextBox1.Text)))
                End If

            End If
            If Not IsNothing(_sebtodrv1) Then
                'If drv1.Row("turnovery1") <> 0 Then
                If TextBox2.Text <> 0 Then
                    'TextBox9.Text = ValidText(String.Format("{0:0.0%}", IIf(drv1.Row("turnovery1") = 0, 0, _sebtodrv1.Row.Item("amount") / drv1.Row("turnovery1"))))
                    TextBox9.Text = ValidText(String.Format("{0:0.0%}", _sebtodrv1.Row.Item("amount") / CDbl(TextBox2.Text)))
                End If

            End If
            If Not IsNothing(_sebtodrv2) Then
                'If drv1.Row("turnovery2") <> 0 Then
                If TextBox3.Text <> 0 Then
                    'TextBox8.Text = ValidText(String.Format("{0:0.0%}", IIf(drv1.Row("turnovery2") = 0, 0, _sebtodrv2.Row.Item("amount") / drv1.Row("turnovery2"))))
                    TextBox8.Text = ValidText(String.Format("{0:0.0%}", _sebtodrv2.Row.Item("amount") / CDbl(TextBox3.Text)))
                End If

            End If
            If Not IsNothing(_sebtodrv3) Then
                'If drv1.Row("turnovery3") <> 0 Then
                If TextBox4.Text <> 0 Then
                    'TextBox7.Text = ValidText(String.Format("{0:0.0%}", IIf(drv1.Row("turnovery3") = 0, 0, _sebtodrv3.Row.Item("amount") / drv1.Row("turnovery3"))))
                    TextBox7.Text = ValidText(String.Format("{0:0.0%}", _sebtodrv3.Row.Item("amount") / CDbl(TextBox4.Text)))
                End If

            End If
            If Not IsNothing(_sebtodrv4) Then
                'If drv1.Row("turnovery4") <> 0 Then
                If TextBox5.Text <> 0 Then
                    'TextBox6.Text = ValidText(String.Format("{0:0.0%}", IIf(drv1.Row("turnovery4") = 0, 0, _sebtodrv4.Row.Item("amount") / drv1.Row("turnovery4"))))
                    TextBox6.Text = ValidText(String.Format("{0:0.0%}", _sebtodrv4.Row.Item("amount") / CDbl(TextBox5.Text)))
                End If

            End If

            TextBox12.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox10.Text.Replace("%", "")) / CDbl(TextBox9.Text.Replace("%", "")) - 1))
            TextBox13.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox9.Text.Replace("%", "")) / CDbl(TextBox8.Text.Replace("%", "")) - 1))
            TextBox17.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox8.Text.Replace("%", "")) / CDbl(TextBox7.Text.Replace("%", "")) - 1))
            TextBox15.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox7.Text.Replace("%", "")) / CDbl(TextBox6.Text.Replace("%", "")) - 1))

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
            yval(0) = CDbl(TextBox1.Text)
            yval(1) = CDbl(TextBox2.Text)
            yval(2) = CDbl(TextBox3.Text)
            yval(3) = CDbl(TextBox4.Text)
            yval(4) = CDbl(TextBox5.Text)
            Chart1.Series(0).Points.DataBindXY(xval, yval)
            'Chart1.ChartAreas(0).AxisY.IsLogarithmic = True
            'Chart1.ChartAreas(0).AxisY.Minimum = [Double].NaN
            'Chart1.ChartAreas(0).AxisY.Maximum = [Double].NaN


            'For i = 0 To 4
            '    Chart1.Series(0).Points.AddXY(xval(i), yval(i))
            'Next

        End If
    End Sub

    Public Sub initLabel() Implements IUTODetail.initLabel
        Label2.Text = "Supplier total (full year) From SIF"
        Label1.Visible = False
        Label23.Text = "Purchase Amount (USD)"
        Label24.Text = "%"

        TextBox19.Visible = False
    End Sub

    Private Sub validblank(ByRef SIFTOBS As BindingSource)
        If Not IsNothing(SIFTOBS) Then
            If Not IsNothing(SIFTOBS.Current) Then
                Dim drv As DataRowView = SIFTOBS.Current

                If IsDBNull(drv.Row.Item("turnovery")) Then drv.Row.Item("turnovery") = 0
                If IsDBNull(drv.Row.Item("turnovery1")) Then drv.Row.Item("turnovery1") = 0
                If IsDBNull(drv.Row.Item("turnovery2")) Then drv.Row.Item("turnovery2") = 0
                If IsDBNull(drv.Row.Item("turnovery3")) Then drv.Row.Item("turnovery3") = 0
                If IsDBNull(drv.Row.Item("turnovery4")) Then drv.Row.Item("turnovery4") = 0
                drv.EndEdit()
            End If
        End If
        

    End Sub

    Private Sub validblankdrv(ByVal sebtodrv As DataRowView)
        If Not IsNothing(sebtodrv) Then
            If IsDBNull(sebtodrv.Row.Item("amount")) Then sebtodrv.Row.Item("amount") = 0
        End If

    End Sub





End Class
Interface IUTODetail
    Sub initLabel()
    Sub DisplayValue()

End Interface