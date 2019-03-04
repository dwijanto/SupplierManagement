Public Class UTODetail

    Public Overridable Sub displayvalue()
        ClearValue()
    End Sub
    Public Overridable Sub ClearValue()
        TextBox1.Text = 0
        TextBox2.Text = 0
        TextBox3.Text = 0
        TextBox4.Text = 0
        TextBox5.Text = 0
        TextBox6.Text = 0
        TextBox7.Text = 0
        TextBox8.Text = 0
        TextBox9.Text = 0
        TextBox10.Text = 0
        TextBox11.Text = "0%"
        TextBox12.Text = "0%"
        TextBox13.Text = "0%"
        TextBox14.Text = "0%"
        TextBox15.Text = "0%"
        TextBox16.Text = "0%"
        TextBox17.Text = "0%"
        TextBox18.Text = "0%"

        'Dim yval(4) As Double
        'Dim xval(4) As String
        'For i = 0 To 4
        '    xval(i) = String.Format("{0}", Year(_currentDate) - i)
        'Next
        For i = 0 To 4
            Chart1.Series(0).Points.Clear()
        Next
        Chart1.ChartAreas(0).AxisY.LabelStyle.Format = "#,##0"
    End Sub


    Public Function ValidText(ByRef TB As String) As String
        Dim myret As String = TB
        Select Case TB
            Case "Infinity"

                myret = ""
            Case "NaN"
                myret = ""
        End Select
        Return myret
    End Function
End Class
