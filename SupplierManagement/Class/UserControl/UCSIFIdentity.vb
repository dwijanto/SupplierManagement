Public Class UCSIFIdentity
    Public Property sifidbs As BindingSource
    Public Property sifbs As BindingSource
    Public Property idbs As BindingSource

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub showData()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = sifidbs

        DataGridView2.AutoGenerateColumns = False
        DataGridView2.DataSource = sifbs

        DataGridView3.AutoGenerateColumns = False
        DataGridView3.DataSource = idbs

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""

        TextBox1.DataBindings.Add(New Binding("text", sifbs, "docdate", True, DataSourceUpdateMode.Never, "", "dd-MMM-yyyy"))
        TextBox2.DataBindings.Add(New Binding("text", idbs, "docdate", True, DataSourceUpdateMode.Never, "", "dd-MMM-yyyy"))
        TextBox3.DataBindings.Add(New Binding("text", sifidbs, "doctypename", True))
        TextBox4.DataBindings.Add(New Binding("text", sifidbs, "docdate", True, DataSourceUpdateMode.Never, "", "dd-MMM-yyyy"))
        DataGridView1.Invalidate()
        DataGridView2.Invalidate()
        DataGridView3.Invalidate()

    End Sub
End Class
