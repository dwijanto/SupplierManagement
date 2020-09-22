Imports System.Text

Public Class FormHelperMulti
    Public bs As BindingSource
    Public Property SelectedResult As String
    Public Property Key As Integer
    Dim SB As StringBuilder
    Public SelectedDrv As New List(Of DataRowView)
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByRef bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.bs = bs 'New BindingSource(bs.DataSource, bs.DataMember)
        Me.bs.Filter = ""
    End Sub

    'Private Sub FormHelper_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    '    Dim drv As DataRowView = bs.Current
    '    MessageBox.Show("hello")
    'End Sub

    Private Sub FormGetValidator_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DataGridView1.AutoGenerateColumns = False
        'DataGridView1.AutoGenerateColumns = True
        DataGridView1.DataSource = bs
    End Sub


    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Button1.PerformClick()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Try
            bs.Filter = String.Format("[" & DataGridView1.Columns(0).DataPropertyName & "] like '*{0}*'", TextBox1.Text.Replace("'", "''"))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        bs.Filter = ""
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        SB = New StringBuilder
        For Each item As DataGridViewRow In DataGridView1.SelectedRows
            If SB.Length > 0 Then
                SB.Append(",")
            End If
            Dim mydr As DataRowView = bs.Item(item.Index)
            SB.Append(mydr.Row.Item(Key))
            SelectedDrv.Add(mydr)
        Next
        SelectedResult = SB.ToString
    End Sub
End Class