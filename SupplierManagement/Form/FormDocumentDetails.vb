Public Class FormDocumentDetails
    Dim DRV As DataRowView
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(ByVal bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        DRV = bs.Current
    End Sub



    Private Sub FormDocumentDetails_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            TextBox1.Text = "" & DRV.Row.Item("shortname")
            TextBox2.Text = "" & String.Format("{0} - {1}", DRV.Row.Item("vendorcode"), DRV.Row.Item("vendorname"))
            TextBox3.Text = "" & DRV.Row.Item("doctypename")
            If Not IsDBNull(DRV.Row.Item("docdate")) Then
                TextBox4.Text = "" & String.Format("{0:dd-MMM-yyyy}", CDate(DRV.Row.Item("docdate")))
            End If
            TextBox4.Text = "" & String.Format("{0:dd-MMM-yyyy}", CDate(DRV.Row.Item("docdate")))
            TextBox8.Text = "" & DRV.Row.Item("remarks")
            If TextBox8.TextLength >= 50 Then
                TextBox8.Enabled = True
            End If

            TextBox9.Text = "" & DRV.Row.Item("version")
            If Not IsDBNull(DRV.Row.Item("expireddate")) Then
                TextBox5.Text = "" & String.Format("{0:dd-MMM-yyyy}", CDate(DRV.Row.Item("expireddate")))
            End If

            TextBox6.Text = "" & DRV.Row.Item("levelname")
            TextBox7.Text = "" & DRV.Row.Item("paymentterm")
            TextBox10.Text = "" & DRV.Row.Item("docname")
            TextBox11.Text = "" & DRV.Row.Item("leadtime")
            TextBox12.Text = "" & DRV.Row.Item("sasl")
            TextBox13.Text = "" & DRV.Row.Item("nqsu")
            TextBox14.Text = "" & DRV.Row.Item("projectname")
            TextBox15.Text = "" & DRV.Row.Item("auditby")
            TextBox16.Text = "" & DRV.Row.Item("audittype")
            TextBox17.Text = "" & DRV.Row.Item("auditgrade")
            TextBox18.Text = "" & DRV.Row.Item("score")
            TextBox19.Text = "" & String.Format("{0:#,##0.00}", DRV.Row.Item("turnovery"))
            TextBox20.Text = "" & String.Format("{0:#,##0.00}", DRV.Row.Item("turnovery1"))
            TextBox21.Text = "" & String.Format("{0:#,##0.00}", DRV.Row.Item("turnovery2"))
            TextBox22.Text = "" & String.Format("{0:#,##0.00}", DRV.Row.Item("turnovery3"))
            TextBox23.Text = "" & String.Format("{0:#,##0.00}", DRV.Row.Item("turnovery4"))
            TextBox24.Text = "" & String.Format("{0:#,##0.00}", DRV.Row.Item("ratioy"))
            TextBox25.Text = "" & String.Format("{0:#,##0.00}", DRV.Row.Item("ratioy1"))
            TextBox26.Text = "" & String.Format("{0:#,##0.00}", DRV.Row.Item("ratioy2"))
            TextBox27.Text = "" & String.Format("{0:#,##0.00}", DRV.Row.Item("ratioy3"))
            TextBox28.Text = "" & String.Format("{0:#,##0.00}", DRV.Row.Item("ratioy4"))
            TextBox29.Text = "" & DRV.Row.Item("returnrate")
            Label3.Text = " Year = " & DRV.Row.Item("myyear")
            If TextBox10.TextLength >= 30 Then
                TextBox10.Enabled = True
            End If
            If TextBox14.TextLength >= 30 Then
                TextBox14.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show("")
        End Try
        
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub


End Class