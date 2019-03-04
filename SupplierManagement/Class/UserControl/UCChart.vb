Imports System.Windows.Forms.DataVisualization.Charting
Public Class UCChart

    Dim _DS As New DataSet
    Public CurrentDate As Date

    Public Property DS As DataSet
        Get
            Return _DS
        End Get
        Set(ByVal value As DataSet)
            _DS = value
            displayChart()
        End Set
    End Property

    Private Sub displayChart()

        'Dim drv As DataView = bs.List
        'Chart
        Dim i As Integer = 0
        Try
            For l = 0 To Chart1.Series.Count - 1
                Chart1.Series.Remove(Chart1.Series(0))
            Next
            If IsNothing(DS) Then
                Exit Sub
            End If
            Dim bs As New BindingSource
            Dim xBS As New BindingSource
            bs.DataSource = DS.Tables(0)
            xBS.DataSource = DS.Tables(1)
            Chart1.ChartAreas(0).AxisY.LabelStyle.Format = "#,##0"
            If Not IsNothing(bs) Then

                For Each drv As DataRowView In bs.List

                    'If i > 0 Then
                    Dim myseries As New Series
                    Chart1.Series.Add(myseries)
                    'End If
                    Chart1.Series(i).ChartType = SeriesChartType.Column
                    Chart1.Series(i)("PointWidth") = "0.6"
                    Chart1.Series(i).Name = "" & Trim(drv.Row.Item(0))
                    Chart1.Series(i).IsValueShownAsLabel = False
                    Chart1.Series(i).ShadowColor = Color.DarkGray
                    Chart1.Series(i).ShadowOffset = 2
                    Chart1.Series(i).IsValueShownAsLabel = CheckBox1.Checked
                    Chart1.Series(i).LabelFormat = "#,##0"
                    Chart1.Series(i).LabelForeColor = Color.Red
                    Chart1.Series(0).SmartLabelStyle.Enabled = True
                    'Chart1.Series(0)("BarLabelStyle") = "Center"

                    'Dim yval As Double() = {30000000, 40000000, 50000000, 10000000, 15000000}
                    'Dim xval As String() = {"2014", "2013", "2012", "2011", "2010"}



                    Dim yval(xBS.Count - 1) As Double
                    Dim xval(xBS.Count - 1) As String

                    For j = 0 To xBS.Count - 1
                        xval(j) = xBS.Item(j).row(0)
                        'If IsDBNull(drv.Row.Item(j + 1)) Then drv.Row.Item(j + 1) = 0
                        'yval(j) = CDbl(drv.Row.Item(j + 1))
                        'yval(j) = Nothing
                        If Not IsDBNull(drv.Row.Item(j + 1)) Then yval(j) = CDbl(drv.Row.Item(j + 1))
                    Next


                    Chart1.Series(i).Points.DataBindXY(xval, yval)
                    i = i + 1
                Next
            End If
            

        Catch ex As Exception
            MessageBox.Show(String.Format("{0}::{1} {2}.", Me.Name, System.Reflection.MethodInfo.GetCurrentMethod(), ex.Message))
        End Try



    End Sub
    Public Sub ClearValue()
        DS = Nothing
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Try
            Chart1.Series(0).IsValueShownAsLabel = CheckBox1.Checked
            'Chart1.Series(0).LabelFormat = "#,##0"

            If CheckBox1.Checked Then
                CheckBox1.Text = "Hide Value Label"
            Else
                CheckBox1.Text = "Show Value Label"
            End If

        Catch ex As Exception

        End Try


    End Sub
End Class
