Imports System.Windows.Forms
Imports System.ComponentModel

Public Class DialogModificationTypeMaster
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
    Dim DRV As DataRowView
    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Public Property InformationType As Integer
        Get
            If RadioButton1.Checked = True Then
                DRV.Row.Item("InformationTypeName") = "Basic Information"
                Return 1
            Else
                DRV.Row.Item("InformationTypeName") = "Bank Information"
                Return 2
            End If
        End Get
        Set(ByVal value As Integer)
            If value = 1 Then
                RadioButton1.Checked = True
            ElseIf value = 2 Then
                RadioButton2.Checked = True
            End If
        End Set
    End Property


    Public Sub New(ByVal drv As DataRowView)
        InitializeComponent()
        Me.DRV = drv
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Me.Validate Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            DRV.EndEdit()
            RaiseEvent RefreshDataGridView(Me, e)
            Me.Close()
        Else
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
        TextBox4.DataBindings.Clear()
        'RichTextBox1.DataBindings.Clear()

        TextBox1.DataBindings.Add(New Binding("Text", DRV, "modifytype", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("Text", DRV, "sensitivitylevel", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("Text", DRV, "lineorder", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox4.DataBindings.Add(New Binding("Text", DRV, "remarks", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        'RichTextBox1.DataBindings.Add(New Binding("Text", DRV, "remarks", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        Me.DataBindings.Add(New Binding("InformationType", DRV, "informationtype", True, DataSourceUpdateMode.OnPropertyChanged))

        'If DRV.Row.RowState = DataRowState.Detached Then
        '    ComboBox1.SelectedIndex = -1
        'End If
    End Sub


    Private Sub Dialog1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        initData()
    End Sub



    Private Sub RadioButton1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        'RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("InformationType"))
        onPropertyChanged("InformationType")
    End Sub

    Private Sub onPropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub

End Class
