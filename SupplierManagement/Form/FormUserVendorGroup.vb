Public Class FormUserVendorGroup
    Private bs As BindingSource
    Private username As String
    Dim myFields As String() = {"groupname", "vendorcode", "vendorname", "shortname"}
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Public Sub New(ByVal bs As BindingSource, ByVal username As String)
        InitializeComponent()
        Me.bs = bs
        Me.username = username
    End Sub

    Private Sub FormUserVendorGroup_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DataGridView1.AutoGenerateColumns = False
        bs.Filter = String.Format("username = '{0}'", username)
        DataGridView1.DataSource = bs
    End Sub
    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        ' SPBS.EndEdit()
        bs.Filter = String.Format("username = '{0}'", username)
        If ToolStripTextBox1.Text <> "" And ToolStripComboBox1.SelectedIndex <> -1 Then
            Select Case ToolStripComboBox1.SelectedIndex
                Case 1
                    If Not IsNumeric(ToolStripTextBox1.Text) Then
                        ToolStripTextBox1.Select()
                        SendKeys.Send("{BACKSPACE}")

                    End If
            End Select
            bs.Filter = bs.Filter & " and " & myFields(ToolStripComboBox1.SelectedIndex).ToString & " like '%" & sender.ToString.Replace("'", "''") & "%'"
        End If
    End Sub


End Class