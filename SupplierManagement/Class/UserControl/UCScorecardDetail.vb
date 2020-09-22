Public Class UCScorecardDetail
    Implements IUTODetail

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

    Private _SEBTOBSFP As BindingSource
    Private _SEBTOBSFP1 As BindingSource
    Private _SEBTOBSFP2 As BindingSource
    Private _SEBTOBSFP3 As BindingSource
    Private _SEBTOBSFP4 As BindingSource


    Private _sebtodrvfp As DataRowView
    Private _sebtodrvfp1 As DataRowView
    Private _sebtodrvfp2 As DataRowView
    Private _sebtodrvfp3 As DataRowView
    Private _sebtodrvfp4 As DataRowView

    Private _NQSUdrv As DataRowView
    Dim _NQSUBS As BindingSource
    Dim _logisticsBS As BindingSource
    Dim _Logisticsdrv As DataRowView
    Dim _logisticsBS1 As BindingSource
    Dim _logisticsBS2 As BindingSource
    Dim _logisticsBS3 As BindingSource
    Dim _logisticsBS4 As BindingSource
    Dim _Logisticsdrv1 As DataRowView
    Dim _Logisticsdrv2 As DataRowView
    Dim _Logisticsdrv3 As DataRowView
    Dim _Logisticsdrv4 As DataRowView
    Dim _pdbs As BindingSource
    Dim _pdbs1 As BindingSource
    Dim _pdbs2 As BindingSource
    Dim _pdbs3 As BindingSource
    Dim _pdbs4 As BindingSource
    Dim _pddrv As Object
    Dim _pddrv1 As Object
    Dim _pddrv2 As Object
    Dim _pddrv3 As Object
    Dim _pddrv4 As Object

    Public Property SEBTOBS As BindingSource
        Get
            Return _SEBTOBS
        End Get
        Set(ByVal value As BindingSource)
            _SEBTOBS = value

            If Not IsNothing(value) Then
                _sebtodrv = value.Current
                'validblankdrv(_sebtodrv)
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
                'validblankdrv(_sebtodrv1)
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
                'validblankdrv(_sebtodrv2)
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
                'validblankdrv(_sebtodrv3)
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
                'validblankdrv(_sebtodrv4)
            End If

        End Set
    End Property

    Public Property SEBTOBSFP As BindingSource
        Get
            Return _SEBTOBSFP
        End Get
        Set(ByVal value As BindingSource)
            _SEBTOBSFP = value

            If Not IsNothing(value) Then
                _sebtodrvfp = value.Current
                'validblankdrv(_sebtodrv)
            End If

        End Set
    End Property
    Public Property SEBTOBSFP1 As BindingSource
        Get
            Return _SEBTOBSFP1
        End Get
        Set(ByVal value As BindingSource)
            _SEBTOBSFP1 = value
            If Not IsNothing(value) Then
                _sebtodrvfp1 = value.Current
                'validblankdrv(_sebtodrv1)
            End If

        End Set
    End Property
    Public Property SEBTOBSFP2 As BindingSource
        Get
            Return _SEBTOBSFP2
        End Get
        Set(ByVal value As BindingSource)
            _SEBTOBSFP2 = value
            If Not IsNothing(value) Then
                _sebtodrvfp2 = value.Current
                'validblankdrv(_sebtodrv2)
            End If

        End Set
    End Property
    Public Property SEBTOBSFP3 As BindingSource
        Get
            Return _SEBTOBSFP3
        End Get
        Set(ByVal value As BindingSource)
            _SEBTOBSFP3 = value
            If Not IsNothing(value) Then
                _sebtodrvfp3 = value.Current
                'validblankdrv(_sebtodrv3)
            End If

        End Set
    End Property
    Public Property SEBTOBSFP4 As BindingSource
        Get
            Return _SEBTOBSFP
        End Get
        Set(ByVal value As BindingSource)
            _SEBTOBSFP4 = value
            If Not IsNothing(value) Then
                _sebtodrvfp4 = value.Current
                'validblankdrv(_sebtodrv4)
            End If

        End Set
    End Property

    Public Property NQSUBS As BindingSource
        Get
            Return _NQSUBS
        End Get
        Set(ByVal value As BindingSource)
            _NQSUBS = value
            If Not IsNothing(value) Then
                _NQSUdrv = value.Current
                'validblankdrv(_sebtodrv4)
            End If

        End Set
    End Property
    Public Property LogisticsBS As BindingSource
        Get
            Return _logisticsBS
        End Get
        Set(ByVal value As BindingSource)
            _logisticsBS = value
            If Not IsNothing(value) Then
                _Logisticsdrv = value.Current               
            End If
        End Set
    End Property

    Public Property LogisticsBS1 As BindingSource
        Get
            Return _logisticsBS1
        End Get
        Set(ByVal value As BindingSource)
            _logisticsBS1 = value
            If Not IsNothing(value) Then
                _Logisticsdrv1 = value.Current
            End If
        End Set
    End Property

    Public Property LogisticsBS2 As BindingSource
        Get
            Return _logisticsBS2
        End Get
        Set(ByVal value As BindingSource)
            _logisticsBS2 = value
            If Not IsNothing(value) Then
                _Logisticsdrv2 = value.Current
            End If
        End Set
    End Property

    Public Property LogisticsBS3 As BindingSource
        Get
            Return _logisticsBS3
        End Get
        Set(ByVal value As BindingSource)
            _logisticsBS3 = value
            If Not IsNothing(value) Then
                _Logisticsdrv3 = value.Current
            End If
        End Set
    End Property

    Public Property LogisticsBS4 As BindingSource
        Get
            Return _logisticsBS4
        End Get
        Set(ByVal value As BindingSource)
            _logisticsBS4 = value
            If Not IsNothing(value) Then
                _Logisticsdrv4 = value.Current
            End If
        End Set
    End Property

    Public Property PDBS As BindingSource
        Get
            Return _pdbs
        End Get
        Set(ByVal value As BindingSource)
            _pdbs = value
            If Not IsNothing(value) Then
                _pddrv = value.Current
            End If
        End Set
    End Property

    Public Property PDBS1 As BindingSource
        Get
            Return _pdbs1
        End Get
        Set(ByVal value As BindingSource)
            _pdbs1 = value
            If Not IsNothing(value) Then
                _pddrv1 = value.Current
            End If
        End Set
    End Property

    Public Property PDBS2 As BindingSource
        Get
            Return _pdbs2
        End Get
        Set(ByVal value As BindingSource)
            _pdbs2 = value
            If Not IsNothing(value) Then
                _pddrv2 = value.Current
            End If
        End Set
    End Property

    Public Property PDBS3 As BindingSource
        Get
            Return _pdbs3
        End Get
        Set(ByVal value As BindingSource)
            _pdbs3 = value
            If Not IsNothing(value) Then
                _pddrv3 = value.Current
            End If
        End Set
    End Property

    Public Property PDBS4 As BindingSource
        Get
            Return _pdbs4
        End Get
        Set(ByVal value As BindingSource)
            _pdbs4 = value
            If Not IsNothing(value) Then
                _pddrv4 = value.Current
            End If
        End Set
    End Property

    'Public Sub DisplayValue() Implements IUTODetail.DisplayValue
    '    If Not IsNothing(_sebtodrv) Then
    '        TextBox1.Text = "" & String.Format("{0:#,##0}", _sebtodrv.Row("qty"))
    '        TextBox10.Text = "" & String.Format("{0:#,##0}", _sebtodrv.Row("amount"))
    '    End If
    '    If Not IsNothing(_sebtodrv1) Then
    '        TextBox2.Text = "" & String.Format("{0:#,##0}", _sebtodrv1.Row("qty"))
    '        TextBox9.Text = "" & String.Format("{0:#,##0}", _sebtodrv1.Row("amount"))
    '    End If
    '    If Not IsNothing(_sebtodrv2) Then
    '        TextBox3.Text = "" & String.Format("{0:#,##0}", _sebtodrv2.Row("qty"))
    '        TextBox8.Text = "" & String.Format("{0:#,##0}", _sebtodrv2.Row("amount"))
    '    End If
    '    If Not IsNothing(_sebtodrv3) Then
    '        TextBox4.Text = "" & String.Format("{0:#,##0}", _sebtodrv3.Row("qty"))
    '        TextBox7.Text = "" & String.Format("{0:#,##0}", _sebtodrv3.Row("amount"))
    '    End If
    '    If Not IsNothing(_sebtodrv4) Then
    '        TextBox5.Text = "" & String.Format("{0:#,##0}", _sebtodrv4.Row("qty"))
    '        TextBox6.Text = "" & String.Format("{0:#,##0}", _sebtodrv4.Row("amount"))
    '    End If

    '    TextBox12.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox10.Text) / CDbl(TextBox9.Text) - 1))
    '    TextBox13.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox9.Text) / CDbl(TextBox8.Text) - 1))
    '    TextBox17.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox8.Text) / CDbl(TextBox7.Text) - 1))
    '    TextBox15.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox7.Text) / CDbl(TextBox6.Text) - 1))

    '    TextBox11.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox1.Text) / CDbl(TextBox2.Text) - 1))
    '    TextBox14.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox2.Text) / CDbl(TextBox3.Text) - 1))
    '    TextBox18.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox3.Text) / CDbl(TextBox4.Text) - 1))
    '    TextBox16.Text = ValidText(String.Format("{0:0%}", CDbl(TextBox4.Text) / CDbl(TextBox5.Text) - 1))


    '    'Chart
    '    Chart1.Series(0).ChartType = SeriesChartType.Column
    '    Chart1.Series(0)("PointWidth") = "0.6"
    '    Chart1.Series(0).Name = ""
    '    Chart1.Series(0).IsValueShownAsLabel = False
    '    'Chart1.Series(0)("BarLabelStyle") = "Center"

    '    'Dim yval As Double() = {30000000, 40000000, 50000000, 10000000, 15000000}
    '    'Dim xval As String() = {"2014", "2013", "2012", "2011", "2010"}
    '    Dim yval(4) As Double
    '    Dim xval(4) As String
    '    For i = 0 To 4
    '        xval(i) = String.Format("{0}", Year(_currentDate) - i)
    '    Next
    '    yval(0) = CDbl(TextBox10.Text)
    '    yval(1) = CDbl(TextBox9.Text)
    '    yval(2) = CDbl(TextBox8.Text)
    '    yval(3) = CDbl(TextBox7.Text)
    '    yval(4) = CDbl(TextBox6.Text)
    '    Chart1.Series(0).Points.DataBindXY(xval, yval)

    'End Sub

    Public Sub DisplayValue() Implements IUTODetail.DisplayValue
        'Scorecard
        ClearValue()
        'FPCP
        If Not IsNothing(_sebtodrv) Then
            TextBox1.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv.Row("piavg"))
            TextBox10.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv.Row("pilkp"))
            TextBox85.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv.Row("piavgfx"))
        End If

        If Not IsNothing(_sebtodrv1) Then
            TextBox2.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv1.Row("piavg"))
            TextBox9.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv1.Row("pilkp"))
            TextBox84.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv1.Row("piavgfx"))
        End If
        If Not IsNothing(_sebtodrv2) Then
            TextBox3.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv2.Row("piavg"))
            TextBox8.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv2.Row("pilkp"))
            TextBox83.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv2.Row("piavgfx"))
        End If
        If Not IsNothing(_sebtodrv3) Then
            TextBox4.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv3.Row("piavg"))
            TextBox7.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv3.Row("pilkp"))
            TextBox82.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv3.Row("piavgfx"))
        End If
        If Not IsNothing(_sebtodrv4) Then
            TextBox5.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv4.Row("piavg"))
            TextBox6.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv4.Row("pilkp"))
            TextBox81.Text = "" & String.Format("{0:#,##0.00}", _sebtodrv4.Row("piavgfx"))
        End If

        'FP
        If Not IsNothing(_sebtodrvfp) Then
            TextBox20.Text = "" & String.Format("{0:#,##0.00}", _sebtodrvfp.Row("pistd"))
            TextBox15.Text = "" & String.Format("{0:#,##0}", _sebtodrvfp.Row("tovariance"))
            TextBox80.Text = "" & String.Format("{0:#,##0.00}", _sebtodrvfp.Row("pistdfx"))
            TextBox90.Text = "" & String.Format("{0:#,##0}", _sebtodrvfp.Row("tovariancefx"))
        End If
        If Not IsNothing(_sebtodrvfp1) Then
            TextBox19.Text = "" & String.Format("{0:#,##0.00}", _sebtodrvfp1.Row("pistd"))
            TextBox14.Text = "" & String.Format("{0:#,##0}", _sebtodrvfp1.Row("tovariance"))
            TextBox79.Text = "" & String.Format("{0:#,##0.00}", _sebtodrvfp1.Row("pistdfx"))
            TextBox89.Text = "" & String.Format("{0:#,##0}", _sebtodrvfp1.Row("tovariancefx"))
        End If
        If Not IsNothing(_sebtodrvfp2) Then
            TextBox18.Text = "" & String.Format("{0:#,##0.00}", _sebtodrvfp2.Row("pistd"))
            TextBox13.Text = "" & String.Format("{0:#,##0}", _sebtodrvfp2.Row("tovariance"))
            TextBox78.Text = "" & String.Format("{0:#,##0.00}", _sebtodrvfp2.Row("pistdfx"))
            TextBox88.Text = "" & String.Format("{0:#,##0}", _sebtodrvfp2.Row("tovariancefx"))
        End If
        If Not IsNothing(_sebtodrvfp3) Then
            TextBox17.Text = "" & String.Format("{0:#,##0.00}", _sebtodrvfp3.Row("pistd"))
            TextBox12.Text = "" & String.Format("{0:#,##0}", _sebtodrvfp3.Row("tovariance"))
            TextBox77.Text = "" & String.Format("{0:#,##0.00}", _sebtodrvfp3.Row("pistdfx"))
            TextBox87.Text = "" & String.Format("{0:#,##0}", _sebtodrvfp3.Row("tovariancefx"))
        End If
        If Not IsNothing(_sebtodrvfp4) Then
            TextBox16.Text = "" & String.Format("{0:#,##0.00}", _sebtodrvfp4.Row("pistd"))
            TextBox11.Text = "" & String.Format("{0:#,##0}", _sebtodrvfp4.Row("tovariance"))
            TextBox76.Text = "" & String.Format("{0:#,##0.00}", _sebtodrvfp4.Row("pistdfx"))
            TextBox86.Text = "" & String.Format("{0:#,##0}", _sebtodrvfp4.Row("tovariancefx"))
        End If

        If Not IsNothing(_NQSUdrv) Then
            TextBox25.Text = "" & String.Format("{0:#,##0}", _NQSUdrv.Row("amount1"))
            TextBox24.Text = "" & String.Format("{0:#,##0}", _NQSUdrv.Row("amount2"))
            TextBox23.Text = "" & String.Format("{0:#,##0}", _NQSUdrv.Row("amount3"))
            TextBox22.Text = "" & String.Format("{0:#,##0}", _NQSUdrv.Row("amount4"))
            TextBox21.Text = "" & String.Format("{0:#,##0}", _NQSUdrv.Row("amount5"))
        End If

        If Not IsNothing(_Logisticsdrv) Then
            TextBox45.Text = "" & String.Format("{0:##0%}", _Logisticsdrv.Row("sasl"))
            TextBox40.Text = "" & String.Format("{0:#,##0%}", _Logisticsdrv.Row("ssl"))
            TextBox35.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv.Row("lt"))
            TextBox30.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv.Row("orderno"))
            TextBox50.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv.Row("shipment"))
            TextBox75.Text = "" & String.Format("{0:#,##0%}", _Logisticsdrv.Row("sslnet"))
        End If
        If Not IsNothing(_Logisticsdrv1) Then
            TextBox44.Text = "" & String.Format("{0:##0%}", _Logisticsdrv1.Row("sasl"))
            TextBox39.Text = "" & String.Format("{0:#,##0%}", _Logisticsdrv1.Row("ssl"))
            TextBox34.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv1.Row("lt"))
            TextBox29.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv1.Row("orderno"))
            TextBox49.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv1.Row("shipment"))
            TextBox74.Text = "" & String.Format("{0:#,##0%}", _Logisticsdrv1.Row("sslnet"))
        End If
        If Not IsNothing(_Logisticsdrv2) Then
            TextBox43.Text = "" & String.Format("{0:##0%}", _Logisticsdrv2.Row("sasl"))
            TextBox38.Text = "" & String.Format("{0:#,##0%}", _Logisticsdrv2.Row("ssl"))
            TextBox33.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv2.Row("lt"))
            TextBox28.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv2.Row("orderno"))
            TextBox48.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv2.Row("shipment"))
            TextBox73.Text = "" & String.Format("{0:#,##0%}", _Logisticsdrv2.Row("sslnet"))
        End If
        If Not IsNothing(_Logisticsdrv3) Then
            TextBox42.Text = "" & String.Format("{0:##0%}", _Logisticsdrv3.Row("sasl"))
            TextBox37.Text = "" & String.Format("{0:#,##0%}", _Logisticsdrv3.Row("ssl"))
            TextBox32.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv3.Row("lt"))
            TextBox27.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv3.Row("orderno"))
            TextBox47.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv3.Row("shipment"))
            TextBox72.Text = "" & String.Format("{0:#,##0%}", _Logisticsdrv3.Row("sslnet"))
        End If
        If Not IsNothing(_Logisticsdrv4) Then
            TextBox41.Text = "" & String.Format("{0:##0%}", _Logisticsdrv4.Row("sasl"))
            TextBox36.Text = "" & String.Format("{0:#,##0%}", _Logisticsdrv4.Row("ssl"))
            TextBox31.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv4.Row("lt"))
            TextBox26.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv4.Row("orderno"))
            TextBox46.Text = "" & String.Format("{0:#,##0}", _Logisticsdrv4.Row("shipment"))
            TextBox71.Text = "" & String.Format("{0:#,##0%}", _Logisticsdrv4.Row("sslnet"))
        End If
        If Not IsNothing(_pddrv) Then
            TextBox70.Text = "" & String.Format("{0:##0}", _pddrv.Row("rp4"))
            TextBox65.Text = "" & String.Format("{0:#,##0}", _pddrv.Row("rp5"))
            TextBox60.Text = "" & String.Format("{0:#,##0}", _pddrv.Row("noofproject"))
            TextBox55.Text = "" & String.Format("{0:#,##0%}", _pddrv.Row("projectontime"))

        End If
        If Not IsNothing(_pddrv1) Then
            TextBox69.Text = "" & String.Format("{0:##0}", _pddrv1.Row("rp4"))
            TextBox64.Text = "" & String.Format("{0:#,##0}", _pddrv1.Row("rp5"))
            TextBox59.Text = "" & String.Format("{0:#,##0}", _pddrv1.Row("noofproject"))
            TextBox54.Text = "" & String.Format("{0:#,##0%}", _pddrv1.Row("projectontime"))

        End If
        If Not IsNothing(_pddrv2) Then
            TextBox68.Text = "" & String.Format("{0:##0}", _pddrv2.Row("rp4"))
            TextBox63.Text = "" & String.Format("{0:#,##0}", _pddrv2.Row("rp5"))
            TextBox58.Text = "" & String.Format("{0:#,##0}", _pddrv2.Row("noofproject"))
            TextBox53.Text = "" & String.Format("{0:#,##0%}", _pddrv2.Row("projectontime"))

        End If
        If Not IsNothing(_pddrv3) Then
            TextBox67.Text = "" & String.Format("{0:##0}", _pddrv3.Row("rp4"))
            TextBox62.Text = "" & String.Format("{0:#,##0}", _pddrv3.Row("rp5"))
            TextBox57.Text = "" & String.Format("{0:#,##0}", _pddrv3.Row("noofproject"))
            TextBox52.Text = "" & String.Format("{0:#,##0%}", _pddrv3.Row("projectontime"))

        End If
        If Not IsNothing(_pddrv4) Then
            TextBox66.Text = "" & String.Format("{0:##0}", _pddrv4.Row("rp4"))
            TextBox61.Text = "" & String.Format("{0:#,##0}", _pddrv4.Row("rp5"))
            TextBox56.Text = "" & String.Format("{0:#,##0}", _pddrv4.Row("noofproject"))
            TextBox51.Text = "" & String.Format("{0:#,##0%}", _pddrv4.Row("projectontime"))

        End If
    End Sub

    Public Sub initLabel() Implements IUTODetail.initLabel

    End Sub

    Public Sub ClearValue()
        For Each obj As Control In Me.Controls
            If obj.GetType() Is GetType(TextBox) Then
                Dim Txt As TextBox = CType(obj, TextBox)
                Txt.Text = ""
            End If           
        Next
    End Sub

    Private Function validateText(ByVal p1 As String) As String
        If p1 = "" Then
            p1 = 0
        End If
        Return p1
    End Function

    Private Function CheckDBNUll(ByVal p1 As Object) As Object
        Dim MyRet As Decimal = 0
        If IsDBNull(p1) Then
            Return MyRet
        Else
            Return p1
        End If    
    End Function

End Class
