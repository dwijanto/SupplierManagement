Imports System.Threading
Imports SupplierManagement.PublicClass
Public Class FormVIMEmailModification
    Private myAdapter As New VendorInformationModificationAdapter
    Private HDid As Long
    Private _vendorcode As Long
    Private DocAttachmentBS As BindingSource
    Private DataTypeBS As BindingSource
    Private myFullPath As String
    Private BS As BindingSource
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Private DRV As DataRowView
    Dim ApprovalDRV As DataRowView

    Public Event RefreshData()

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Try

            If myAdapter.loaddata(HDid, _vendorcode) Then
                ProgressReport(4, "InitData")
                ProgressReport(1, "Loading Data.Done!")
                ProgressReport(5, "Continuous")
            End If
        Catch ex As Exception
            ProgressReport(1, "Loading Data. Error::" & ex.Message)
            ProgressReport(5, "Continuous")
        End Try
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel1.Text = message
                    Case 4
                        'Check Add or Update
                        BS = myAdapter.getBindingSource
                        DocAttachmentBS = myAdapter.GetDocumentBS

                        DataTypeBS = myAdapter.getDocTypeBS
                        DataTypeBS.Filter = "nvalue = 1"
                        myFullPath = myAdapter.getFileSourceFullPath



                        DRV = DocAttachmentBS.AddNew
                        Dim mydoctypedrv = DataTypeBS.Current
                        DRV.Row.Item("doctypeid") = mydoctypedrv.item("id")

                        BindingControl()

                        'Create New Document for validation

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

    End Sub

    Public Sub loaddata()
        If Not myThread.IsAlive Then
            'ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub


    Public Sub New(ByVal hdid As Long, ByVal vendorcode As Long)

        ' This call is required by the designer.
        InitializeComponent()
        Me.HDid = hdid
        _vendorcode = vendorcode
        ' Add any initialization after the InitializeComponent() call.


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BindingControl()


        ComboBox1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()

        ComboBox1.DataSource = DataTypeBS
        ComboBox1.DisplayMember = "doctype"
        ComboBox1.ValueMember = "id"

        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "doctypeid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox1.DataBindings.Add(New Binding("Text", DRV, "docname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("Text", DRV, "remarks", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        'ComboBox1.SelectedIndex = -1
    End Sub

    Private Sub FormVIMEmailModification_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim openfiledialog1 As New OpenFileDialog
        If openfiledialog1.ShowDialog = DialogResult.OK Then            
            DRV.Row.Item("docfullname") = IO.Path.GetFullPath(openfiledialog1.FileName)
            DRV.Row.Item("docname") = IO.Path.GetFileName(openfiledialog1.FileName)
            DRV.Row.Item("docext") = IO.Path.GetExtension(openfiledialog1.FileName)
            TextBox1.Text = IO.Path.GetFileName(openfiledialog1.FileName)
        End If
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If Me.validate Then
            Dim HeaderDrv = BS.Current

            Dim mydrv = ComboBox1.SelectedItem
            HeaderDrv.row.item("status") = mydrv.Row.Item("ivalue")

            ApprovalDRV = myAdapter.VendorInfoModiActionBS.AddNew
            ApprovalDRV.Row.Item("status") = HeaderDrv.row.Item("status")
            ApprovalDRV.Row.Item("modifiedby") = HelperClass1.UserId
            ApprovalDRV.Row.Item("vendorinfomodiid") = HeaderDrv.row.Item("id")
            ApprovalDRV.EndEdit()

            If SaveRecord() Then
                'send email to next validator
                Dim mySendEmail As FormVendorInformationModification = New FormVendorInformationModification
                mySendEmail.SendEmail(HeaderDrv)
                Me.Close()
                RaiseEvent RefreshData()
            End If
        End If
        
    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        Dim mydrv = ComboBox1.SelectedItem
        DRV.Row.Item("doctypeid") = mydrv.row.item("id")
        DRV.EndEdit()
        ErrorProvider1.SetError(TextBox1, "")
        DRV.Row.RowError = ""
        If IsDBNull(DRV.Row.Item("docfullname")) Then
            DRV.Row.RowError = "File name cannot be blank."
            ErrorProvider1.SetError(TextBox1, "File name cannot be blank.")
            myret = False           
        End If

        Return myret
    End Function

    Private Function SaveRecord() As Boolean
        Dim myret As Boolean = False
        BS.EndEdit()
        DRV.EndEdit()
        ProgressReport(1, "")
        If Me.validate Then
            myret = myAdapter.save()
        Else
            ProgressReport(1, "Error:: Please check details.")
        End If
        Return myret
    End Function
End Class