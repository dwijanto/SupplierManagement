Imports System.Windows.Forms

Public Class DialogQualityAppendixOther

    Private DRV As DataRowView

    Public Sub New(ByRef DRV As DataRowView)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.DRV = DRV
        initialDataRow()
    End Sub
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.validate Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub initialDataRow()
        TextBox1.DataBindings.Add(New Binding("Text", DRV, "nqsu", False, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker1.DataBindings.Add(New Binding("Text", DRV, "otherdate", False, DataSourceUpdateMode.OnPropertyChanged))
        CheckBox1.DataBindings.Add(New Binding("checked", DRV, "status", False, DataSourceUpdateMode.OnPropertyChanged))
    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True

        ErrorProvider1.SetError(TextBox1, "")
        ErrorProvider1.SetError(DateTimePicker1, "")
        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Value cannot be null!")
            myret = False
        End If

        'If DateTimePicker1.Value.Date = Today.Date Then
        '    ErrorProvider1.SetError(DateTimePicker1, "Please change your date!")
        '    myret = False
        'End If
        Return myret
    End Function

End Class
