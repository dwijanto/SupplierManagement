Imports System.Text
Public Class UCPriceCMMF
    Public Property VendorQueryType As FormSupplierDashboard.VendorQuery
    Public Property vendorcode As String
    Public Property shortname As String
    Private _bs As BindingSource
    Dim myfilter As String = String.Empty
    Dim textsb As New StringBuilder

    Public Property bs As BindingSource
        Get
            Return _bs
        End Get
        Set(ByVal value As BindingSource)
            _bs = value
            DataGridView1.DataSource = value
        End Set
    End Property

    Dim LKPFilter As String = String.Empty


    Public Function GetLKPFilter() As String
        Dim myfilter = DialogLKPFilter.GetInstance
        Return myfilter.GetLKPFilter
    End Function

    Public Function GetCMMFFilter() As String
        If TextBox1.Text = "" Then
            Return "ALL"
        Else
            Return TextBox1.Text
        End If
    End Function

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'bs = DataGridView1.DataSource
    End Sub
    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        'Dim bs As BindingSource = DataGridView1.DataSource
        If Not IsNothing(bs) Then
            Dim myform As New FormPriceHistory(bs, VendorQueryType, vendorcode, shortname)
            myform.ShowDialog()
        End If
    End Sub




    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'bs = DataGridView1.DataSource

        If IsNothing(bs) Then
            Exit Sub
        End If
        If IsNothing(bs.DataSource) Then
            MessageBox.Show("Nothing to filter. ")
            Exit Sub
        End If
        Dim myform = DialogLKPFilter.GetInstance

        Dim sb As New StringBuilder


        'If myform.ShowDialog = DialogResult.OK Then
        myform.ShowDialog()
        LKPFilter = ""
        If myform.CheckBox1.Checked Then
            sb.Append(String.Format("(lkp > 0)"))
        End If
        If myform.CheckBox2.Checked Then
            If sb.Length > 0 Then
                sb.Append(" or ")
            End If
            sb.Append(String.Format("(lkp_1 > 0)"))
        End If
        If myform.CheckBox3.Checked Then
            If sb.Length > 0 Then
                sb.Append(" or ")
            End If
            sb.Append(String.Format("(lkp_2 > 0)"))
        End If
        If myform.CheckBox4.Checked Then
            If sb.Length > 0 Then
                sb.Append(" or ")
            End If
            sb.Append(String.Format("(lkp_3 > 0)"))
        End If
        If myform.CheckBox5.Checked Then
            If sb.Length > 0 Then
                sb.Append(" or ")
            End If
            sb.Append(String.Format("(lkp_4 > 0)"))
        End If

        If sb.ToString.Length > 0 Then
            LKPFilter = String.Format("({0})", sb.ToString)
        Else
        End If
        bs.Filter = String.Format("{0} {1} {2}", myfilter, IIf(myfilter.Length > 0 And LKPFilter.Length > 0, " and ", ""), LKPFilter)
        ' End If

    End Sub
    Public Sub ResetFilter()
        Dim myform = DialogLKPFilter.GetInstance
        myform.resetFilter()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Not IsNothing(bs) Then
            bs.Filter = ""            
        End If

        textsb.Clear()
        myfilter = ""
        TextBox1.Text = ""
        LKPFilter = ""
        ResetFilter()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        bs = DataGridView1.DataSource

        If IsNothing(bs) Then
            Exit Sub
        End If
        If IsNothing(bs.DataSource) Then
            MessageBox.Show("Nothing to show. ")
            Exit Sub
        End If
        Dim myBindingSource As New BindingSource
        myBindingSource.DataSource = DirectCast(bs.DataSource, DataTable).Copy
        Dim cmmfhelper = DialogCMMFHelper.GetInstance

        cmmfhelper.DataGridView1.AutoGenerateColumns = False
        cmmfhelper.DataGridView1.DataSource = myBindingSource
        cmmfhelper.Column1.DataPropertyName = "cmmfdesc"




        If cmmfhelper.ShowDialog = DialogResult.OK Then
            For Each drv As DataGridViewRow In cmmfhelper.DataGridView1.SelectedRows
                If textsb.Length > 0 Then
                    textsb.Append(",")
                End If
                textsb.Append(String.Format("{0}", DirectCast(drv.Cells.Item(1), DataGridViewTextBoxCell).Value))
            Next
        End If
        TextBox1.Text = textsb.ToString

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim CMMFFilter As New StringBuilder
        If TextBox1.Text.Length > 0 Then
            For Each mytext In TextBox1.Text.Split(",")
                If mytext.Length > 0 Then
                    If CMMFFilter.Length > 0 Then
                        CMMFFilter.Append("or")
                    End If

                    CMMFFilter.Append(String.Format("(cmmf ={0})", mytext))
                End If




            Next
            myfilter = ""
            If CMMFFilter.Length > 0 Then
                myfilter = String.Format("({0})", CMMFFilter.ToString)
            End If

        End If
        Try
            bs.Filter = String.Format("{0} {1} {2}", myfilter, IIf(myfilter.Length > 0 And LKPFilter.Length > 0, " and ", ""), LKPFilter)
        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Debug.Print("hello")
        If Button4.Text = "Show Currency" Then
            Button4.Text = "Hide Currency"
        Else
            Button4.Text = "Show Currency"
        End If
        DataGridView1.Columns("fobcurr").Visible = Not DataGridView1.Columns("fobcurr").Visible
        DataGridView1.Columns("fobcurr1").Visible = Not DataGridView1.Columns("fobcurr1").Visible
        DataGridView1.Columns("fobcurr2").Visible = Not DataGridView1.Columns("fobcurr2").Visible
        DataGridView1.Columns("fobcurr3").Visible = Not DataGridView1.Columns("fobcurr3").Visible
        DataGridView1.Columns("fobcurr4").Visible = Not DataGridView1.Columns("fobcurr4").Visible
        DataGridView1.Columns("lkpcrcy").Visible = Not DataGridView1.Columns("lkpcrcy").Visible
        DataGridView1.Columns("lkpcrcy1").Visible = Not DataGridView1.Columns("lkpcrcy1").Visible
        DataGridView1.Columns("lkpcrcy2").Visible = Not DataGridView1.Columns("lkpcrcy2").Visible
        DataGridView1.Columns("lkpcrcy3").Visible = Not DataGridView1.Columns("lkpcrcy3").Visible
        DataGridView1.Columns("lkpcrcy4").Visible = Not DataGridView1.Columns("lkpcrcy4").Visible
    End Sub
End Class
