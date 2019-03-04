Imports System.Threading
Public Class FormPTEffectiveDate
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myController As PaymentTermEffDateAdapter
    Dim PaymentTermController As PaymentTermAdapter
    Dim VendorController As VendorTxAdapter
    Dim drv As DataRowView = Nothing
    Dim checkbox1 As CheckBox
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        RemoveHandler DialogPaymentTermDate.RefreshDataGridView, AddressOf RefreshDataGridView
        AddHandler DialogPaymentTermDate.RefreshDataGridView, AddressOf RefreshDataGridView

        checkbox1 = New CheckBox
        checkbox1.Text = "Latest Payment Term )"
        Dim host As ToolStripControlHost = New ToolStripControlHost(checkbox1)
        ToolStrip1.Items.Insert(ToolStrip1.Items.Count, host)


    End Sub


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loadData()
    End Sub

    Private Sub loadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        myController = New PaymentTermEffDateAdapter
        PaymentTermController = New PaymentTermAdapter
        VendorController = New VendorTxAdapter
        Try
            ProgressReport(1, "Loading..")
            If myController.LoadData() Then
                ProgressReport(4, "Init Data")
            End If
            PaymentTermController.loaddata()
            VendorController.loaddata()
            ProgressReport(1, "Done.")
        Catch ex As Exception

            ProgressReport(1, ex.Message)
        End Try
    End Sub


    Public Sub showTx(ByVal tx As TxEnum)
        If Not myThread.IsAlive Then
            Select Case tx
                Case TxEnum.NewRecord
                    drv = myController.GetNewRecord
                    Me.drv.BeginEdit()
                    drv.Row.Item("effectivedate") = Date.Today                    
                Case TxEnum.UpdateRecord
                    drv = myController.GetCurrentRecord
                    Me.drv.BeginEdit()
            End Select
            Dim VendorBS = VendorController.GetBindingSource
            Dim VendorHelperBS = VendorController.GetBindingSource
            Dim PaymentTermBS = PaymentTermController.GetBindingSource
            Dim PaymentTermBSHelper = PaymentTermController.GetBindingSource
           

            Dim myform = New DialogPaymentTermDate(drv, VendorBS, VendorHelperBS, PaymentTermBS, PaymentTermBSHelper)
            myform.ShowDialog()
        End If

    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myController.BS
            End Select
        End If
    End Sub


    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        Dim obj As ToolStripTextBox = DirectCast(sender, ToolStripTextBox)
        myController.ApplyFilter = obj.Text
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        loadData()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        showTx(TxEnum.NewRecord)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        showTx(TxEnum.UpdateRecord)
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(myController.GetCurrentRecord) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myController.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ToolStripButton2.PerformClick()
    End Sub

    Private Sub RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        myController.save()
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        Dim mymessage As String = String.Empty
        'MessageBox.Show(checkbox1.Checked)
        Dim sqlstr As String = String.Empty
        Select Case checkbox1.Checked
            Case True
                sqlstr = "with data as (select distinct vendorcode,first_value(paymenttermid) over (partition by vendorcode order by effectivedate desc) as paymenttermid,first_value(effectivedate) over (partition by vendorcode order by effectivedate desc) as effectivedate from doc.paymenttermdate)" &
                         " select data.vendorcode,v.vendorname::text,v.shortname::text,pt.payt,pt.details, data.effectivedate from data" &
                         " left join vendor v on v.vendorcode = data.vendorcode" &
                         " left join paymentterm pt on pt.paymenttermid = data.paymenttermid order by data.vendorcode ;"
            Case False
                sqlstr = "select data.vendorcode,v.vendorname::text,v.shortname::text,pt.payt,pt.details, data.effectivedate from doc.paymenttermdate data" &
                         " left join vendor v on v.vendorcode = data.vendorcode" &
                         " left join paymentterm pt on pt.paymenttermid = data.paymenttermid order by data.vendorcode,data.paymenttermid;"

        End Select



        Dim mysaveform As New SaveFileDialog
        mysaveform.DefaultExt = "xlsx"
        mysaveform.FileName = String.Format("PaymentTermEffectiveDate{0:yyyyMMdd}", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1 'because hidden

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            myreport.Run(Me, e)
        End If
    End Sub

    Private Sub FormattingReport()
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable()
        'Throw New NotImplementedException
    End Sub

End Class