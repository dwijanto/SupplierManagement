Imports System.Windows.Forms
Imports System.Threading

Public Class DialogVendorSelection
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim BSVendorHelper As New BindingSource
    Dim VendorModel1 As New VendorModel

    Public VendorCode As Long
    Public VendorDR As DataRow
    Public YearRef As Integer

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        ErrorProvider1.SetError(TextBox1, "")
        If VendorCode > 0 Then
            YearRef = DateTimePicker1.Value.Year
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        Else
            ErrorProvider1.SetError(TextBox1, "Please select vendorcode from button helper.")
        End If

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DialogVendorSelection_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ComboBox1.SelectedIndex = 0
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            'ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        BSVendorHelper = VendorModel1.GetVendorCodeShortnameBS
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try
                Select id
                    '    Case 1
                    '        ToolStripStatusLabel1.Text = message
                    '    Case 2
                    '        ToolStripStatusLabel1.Text = message
                    '    Case 4


                    '    Case 5
                    '        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    '    Case 6
                    '        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        While IsNothing(BSVendorHelper)

        End While
        ErrorProvider1.SetError(TextBox1, "")
        Dim myform = New FormHelper(BSVendorHelper)
        myform.Width = 650
        myform.DataGridView1.Columns(0).DataPropertyName = "displayname"
        myform.DataGridView1.Columns(0).Width = 550
        If myform.ShowDialog() = DialogResult.OK Then
            Dim drv = BSVendorHelper.Current
            VendorCode = drv.row.item("vendorcode")
            TextBox1.Text = VendorCode
            VendorDR = drv.row
        End If
    End Sub
End Class
