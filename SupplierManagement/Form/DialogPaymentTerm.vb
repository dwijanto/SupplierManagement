Imports System.Windows.Forms

Public Class DialogPaymentTerm

    Dim DRV As DataRowView


    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Sub New(ByVal drv As DataRowView)
        InitializeComponent()
        Me.DRV = drv
    End Sub

    Public Overloads Function Validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(TextBox1, "")
        ErrorProvider1.SetError(TextBox2, "")
        ErrorProvider1.SetError(TextBox3, "")
        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Please fill in the value.")
            myret = False
        End If
        If TextBox2.Text = "" Then
            ErrorProvider1.SetError(TextBox2, "Please fill in the value.")
            myret = False
        End If
        If TextBox3.Text = "" Then
            ErrorProvider1.SetError(TextBox3, "Please fill in the value.")
            myret = False
        End If

        Return myret
    End Function



    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.Validate Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            DRV.EndEdit()
            RaiseEvent RefreshDataGridView(Me, e)
            Me.Close()        
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        DRV.CancelEdit()

        RaiseEvent RefreshDataGridView(Me, e)
        Me.Close()
    End Sub
    Private Sub initData()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()

        TextBox1.DataBindings.Add(New Binding("Text", DRV, "payt", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("Text", DRV, "days", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("Text", DRV, "details", True, DataSourceUpdateMode.OnPropertyChanged, ""))
    End Sub



    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged
        RaiseEvent RefreshDataGridView(Me, e)
    End Sub
End Class
