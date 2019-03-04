Imports System.Windows.Forms

Public Class DialogValidator
    Private BSValidator As BindingSource
    Public validator As String = String.Empty
    Public Sub New(ByVal bsValidator)
        InitializeComponent()
        Me.BSValidator = bsValidator

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DialogValidator_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub

    Private Sub initData()
        ComboBox1.DisplayMember = "name"
        ComboBox1.ValueMember = "userid"
        ComboBox1.DataSource = BSValidator
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim mydrv As DataRowView = ComboBox1.SelectedItem
        validator = mydrv.Row.Item("userid")
    End Sub
End Class
