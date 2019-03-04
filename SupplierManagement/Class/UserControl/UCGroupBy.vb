Imports System.Windows.Forms.DataVisualization.Charting
Public Class UCGroupBy

    Dim _bs As BindingSource
    Public CurrentDate As Date

    Public Property bs As BindingSource
        Get
            Return _bs
        End Get
        Set(ByVal value As BindingSource)
            _bs = value
            displayDataGrid()
        End Set
    End Property

    Private Sub displayDataGrid()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = _bs
        'Dim drv As DataView = bs.List
        'Chart
        Dim i As Integer = 0
        Try
            For l = 0 To Chart1.Series.Count - 1
                Chart1.Series.Remove(Chart1.Series(0))
            Next
            If Not IsNothing(bs) Then
                For Each drv As DataRowView In bs.List
                    'If i > 0 Then
                    Dim myseries As New Series
                    Chart1.Series.Add(myseries)
                    'End If
                    Chart1.Series(i).ChartType = SeriesChartType.StackedColumn
                    Chart1.Series(i)("PointWidth") = "0.6"
                    Chart1.Series(i).Name = "" & drv.Row.Item("sbu")
                    Chart1.Series(i).IsValueShownAsLabel = False
                    Chart1.Series(i).ShadowColor = Color.DarkGray
                    Chart1.Series(i).ShadowOffset = 2
                    'Chart1.Series(0)("BarLabelStyle") = "Center"

                    'Dim yval As Double() = {30000000, 40000000, 50000000, 10000000, 15000000}
                    'Dim xval As String() = {"2014", "2013", "2012", "2011", "2010"}



                    Dim yval(4) As Double
                    Dim xval(4) As String
                    For j = 0 To 4
                        xval(j) = String.Format("{0}", Year(CurrentDate) - j)
                        If IsDBNull(drv.Row.Item(j + 2)) Then
                            drv.Row.Item(j + 2) = 0
                        End If
                    Next

                    yval(0) = CDbl(drv.Row.Item("amount1"))
                    yval(1) = CDbl(drv.Row.Item("amount2"))
                    yval(2) = CDbl(drv.Row.Item("amount3"))
                    yval(3) = CDbl(drv.Row.Item("amount4"))
                    yval(4) = CDbl(drv.Row.Item("amount5"))
                    Chart1.Series(i).Points.DataBindXY(xval, yval)
                    i = i + 1
                Next
            End If
            Chart1.ChartAreas(0).AxisY.LabelStyle.Format = "#,##0"

        Catch ex As Exception
            MessageBox.Show(String.Format("{0}::{1} {2}.", Me.Name, System.Reflection.MethodInfo.GetCurrentMethod(), ex.Message))
        End Try



    End Sub
    Public Sub ClearValue()
        bs = Nothing
    End Sub
End Class
