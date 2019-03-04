Imports System.Threading
Public Class FormVendorFamilySubFamilyVC
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myController As VendorFamilySubFamilyVCAdapter
    Dim FamilyVCController As New FamilyVCAdapter
    Dim SubFamilyVCController As New SubFamilyVCAdapter
    Dim FamilyVCBS As New BindingSource
    Dim SubFamilyVCBS As New BindingSource
    Dim drv As DataRowView = Nothing
    Private vendorcode As Long
    Private VendorName As String
    Private FamilyVCHelper As New BindingSource
    Private SubFamilyVCHelper As New BindingSource


    Public Sub New(ByVal VendorCode, ByVal Vendorname)
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        Me.vendorcode = VendorCode
        Me.VendorName = Vendorname
    End Sub
    'Public Sub New(ByVal VendorCode As Long, ByVal Vendorname As String, ByVal FamilyVCHelper As BindingSource, ByVal SubFamilyVCHelper As BindingSource)
    '    ' This call is required by the designer.
    '    InitializeComponent()
    '    ' Add any initialization after the InitializeComponent() call.
    '    Me.vendorcode = VendorCode
    '    Me.VendorName = Vendorname
    '    Me.FamilyVCHelper = FamilyVCHelper
    '    Me.SubFamilyVCHelper = SubFamilyVCHelper
    'End Sub
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loadData()
    End Sub

    Private Sub loadData()
        Label1.Text = String.Format("{0} - {1}", vendorcode, VendorName)
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        myController = New VendorFamilySubFamilyVCAdapter
        Try
            ProgressReport(1, "Loading..")
            If myController.LoadData(vendorcode) Then
                FamilyVCController.LoadData()
                SubFamilyVCController.LoadData()
                ProgressReport(4, "Init Data")
            End If



            ProgressReport(1, "Loading..Done")
        Catch ex As Exception

            ProgressReport(1, ex.Message)
        End Try
        
    End Sub


    Public Sub showTx(ByVal tx As TxEnum)
        If Not myThread.IsAlive Then
            Select Case tx
                Case TxEnum.NewRecord
                    drv = myController.GetNewRecord
                    drv.Item("vendorcode") = vendorcode
                Case TxEnum.UpdateRecord
                    drv = myController.GetCurrentRecord
            End Select
            Me.drv.BeginEdit()
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
                  
                    FamilyVCBS.DataSource = FamilyVCController.GetBindingSource
                    FamilyVCHelper.DataSource = FamilyVCController.GetBindingSource
                    SubFamilyVCBS.DataSource = SubFamilyVCController.GetBindingSource
                    SubFamilyVCHelper.DataSource = SubFamilyVCController.GetBindingSource

                    ComboBox1.DataBindings.Clear()

                    ComboBox1.DataSource = FamilyVCBS
                    ComboBox1.DisplayMember = "familyvcdesc"
                    ComboBox1.ValueMember = "familyvc"

                    ComboBox1.SelectedIndex = -1
                    ComboBox1.DataBindings.Add("SelectedValue", myController.BS, "familyvc", True, DataSourceUpdateMode.OnPropertyChanged)

                    ComboBox2.DataBindings.Clear()

                    ComboBox2.DataSource = SubFamilyVCBS
                    ComboBox2.DisplayMember = "subfamilyvcdesc"
                    ComboBox2.ValueMember = "subfamilyvc"
                    ComboBox2.SelectedIndex = -1
                    ComboBox2.DataBindings.Add("SelectedValue", myController.BS, "subfamilyvc", True, DataSourceUpdateMode.OnPropertyChanged)

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

    End Sub

    Private Sub RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.Validate()
        myController.save()
    End Sub

    Private Sub DataGridView1_DataBindingComplete(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewBindingCompleteEventArgs) Handles DataGridView1.DataBindingComplete
        Dim mydrv = myController.BS.Current
        ComboBox1.Enabled = Not (IsNothing(mydrv))
        ComboBox2.Enabled = Not (IsNothing(mydrv))
        Button1.Enabled = Not (IsNothing(mydrv))
        Button2.Enabled = Not (IsNothing(mydrv))
        If IsNothing(mydrv) Then
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
        End If
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted, ComboBox2.SelectionChangeCommitted
        Dim drv = myController.BS.Current
        Dim cbdrv1 = ComboBox1.SelectedItem
        If Not IsNothing(cbdrv1) Then
            drv.item("familyvcdescription") = cbdrv1.item("description")
        End If
        Dim cbdrv2 = ComboBox2.SelectedItem
        If Not IsNothing(cbdrv2) Then
            drv.item("subfamilyvcdescription") = cbdrv2.item("description")
        End If

        DataGridView1.Invalidate()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim myform = New FormHelper(FamilyVCHelper)
        myform.DataGridView1.Columns(0).DataPropertyName = "familyvcdesc"
        If myform.ShowDialog = DialogResult.OK Then
            Dim drv As DataRowView = FamilyVCHelper.Current
            Dim mydrv As DataRowView = myController.BS.Current
            mydrv.BeginEdit()
            mydrv.Row.Item("familyvc") = drv.Row.Item("familyvc")
            mydrv.Row.Item("familyvcdescription") = drv.Row.Item("description")
            mydrv.EndEdit()
            'Need bellow code to sync with combobox
            Dim myposition = FamilyVCBS.Find("familyvc", drv.Row.Item("familyvc"))
            FamilyVCBS.Position = myposition
            DataGridView1.Invalidate()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim myform = New FormHelper(SubFamilyVCHelper)
        myform.DataGridView1.Columns(0).DataPropertyName = "subfamilyvcdesc"
        If myform.ShowDialog = DialogResult.OK Then
            Dim drv As DataRowView = SubFamilyVCHelper.Current
            Dim mydrv As DataRowView = myController.BS.Current
            mydrv.BeginEdit()
            mydrv.Row.Item("subfamilyvc") = drv.Row.Item("subfamilyvc")
            mydrv.Row.Item("subfamilyvcdescription") = drv.Row.Item("description")
            mydrv.EndEdit()
            'Need bellow code to sync with combobox
            Dim myposition = SubFamilyVCBS.Find("subfamilyvc", drv.Row.Item("subfamilyvc"))
            SubFamilyVCBS.Position = myposition
            DataGridView1.Invalidate()
        End If
    End Sub
End Class