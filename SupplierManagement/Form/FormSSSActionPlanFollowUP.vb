Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Imports Microsoft.Office.Interop

Public Enum TXModeEnum
    ValidateMode = 1
    RegularMode = 2
End Enum

Public Class FormSSSActionPlanFollowUP


    Dim TxMode As TXModeEnum
    Dim ShortNameHash As New Hashtable
    Dim VendorNameHash As New Hashtable
    Dim HashTable1 As New Hashtable
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myController As New ActionPlanController
    Dim drv As DataRowView = Nothing

    Dim VendorModel1 As New VendorModel
    Dim ShortNameBS As New BindingSource
    Dim VendorcodeBS As New BindingSource
    Dim StatusBS As New BindingSource
    Dim SortNameCriteriaList As StringBuilder
    Dim VendornameCriteriaList As StringBuilder
    Dim StatusCriteriaList As New StringBuilder
    Dim MyC As String
    Dim typeofDate As String() = {"finishdate", "startdate", "enddate"}

    Dim OfficerSEBModel1 As New OfficerSEBModel
    Dim myValidatorBS As BindingSource

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        'RemoveHandler DialogCMMF.RefreshDataGridView, AddressOf RefreshDataGridView
        'AddHandler DialogCMMF.RefreshDataGridView, AddressOf RefreshDataGridView        
        TxMode = TXModeEnum.RegularMode
    End Sub


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'loadData()

        initData()
        ToolStripButton7.PerformClick()

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
        myController = New ActionPlanController

        Try
            ProgressReport(1, "Loading..")
            If SortNameCriteriaList.Length = 0 Then
                If HelperClass1.UserInfo.IsAdmin Then
                    If myController.loaddata(MyC) Then
                        ProgressReport(4, "Init Data")
                    End If
                Else
                    If myController.loaddataVendorUser(MyC) Then
                        ProgressReport(4, "Init Data")
                    End If
                End If

            Else
                If myController.loaddata(MyC, SortNameCriteriaList.ToString) Then
                    ProgressReport(4, "Init Data")
                End If
            End If

            ProgressReport(1, "Done.")

        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
    End Sub


    Public Sub showTx(ByVal tx As TxEnum)
        If IsNothing(myController) Then
            MessageBox.Show("Refresh the query first.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If Not myThread.IsAlive And Not IsNothing(myController) Then
            Select Case tx
                Case TxEnum.NewRecord
                    drv = myController.GetNewRecord
                    Me.drv.BeginEdit()

                Case TxEnum.UpdateRecord
                    drv = myController.GetCurrentRecord
                    Me.drv.BeginEdit()

            End Select

            Dim myform = New DialogActionPlan(drv, TxMode)
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
                    DeleteToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin
            End Select
        End If
    End Sub


    'Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
    '    Dim obj As ToolStripTextBox = DirectCast(sender, ToolStripTextBox)
    '    myController.ApplyFilter = obj.Text
    'End Sub




    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        'Import
        ' Dim mybutton As Button = DirectCast(sender, Button)
        myValidatorBS = New BindingSource
        myValidatorBS.DataSource = OfficerSEBModel1.GetValidator
        Dim getValidator As New DialogValidator(myValidatorBS)
        Dim ImportFile = New OpenFileDialog
        ImportFile.FileName = "Select File"
        If ImportFile.ShowDialog = Windows.Forms.DialogResult.OK Then
            'Dim mydrv As DataRowView = bsDetail.Current
            MyC = " where ac.id < 0"
            myController.loaddata(MyC)

            getValidator.ShowDialog()

            'Dim myobj As New ClassImportActionPlan(myController.BS, ImportFile.FileName)
            Dim myobj As New ClassImportActionPlan(myController.BS, ImportFile.FileName, getValidator.validator)
            If Not myobj.getNewRecord() Then
                MessageBox.Show(myobj.errormessage)

                'Exit Sub
            End If
            DataGridView1.DataSource = myController.BS
            DataGridView1.Invalidate()
        End If

    End Sub
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        TxMode = TXModeEnum.ValidateMode
        If HelperClass1.UserInfo.IsAdmin Then
            MyC = "where status = 'Closed' and validatorflag = false"
        Else
            MyC = String.Format(" where lower(validator) = '{0}' and status = 'Closed' and validatorflag = false", HelperClass1.UserId.ToLower)
        End If
        loadData()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        myController.save()
    End Sub
    Private Sub AddNewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddNewToolStripMenuItem.Click

        showTx(TxEnum.NewRecord)


    End Sub

    Private Sub UpdateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateToolStripMenuItem.Click
        showTx(TxEnum.UpdateRecord)
    End Sub
    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex <> -1 Then
            UpdateToolStripMenuItem.PerformClick()
        End If

    End Sub

    Private Sub RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)
        DataGridView1.Invalidate()
    End Sub

    Private Sub initData()
        ToolStripButton8.Visible = HelperClass1.UserInfo.IsAdmin


        StatusBS = myController.Model.getStatusBS
        If HelperClass1.UserInfo.IsAdmin Then
            ShortNameBS = VendorModel1.GetShortNameBS
            VendorcodeBS = VendorModel1.GetVendorCodeBS
        Else
            ShortNameBS = VendorModel1.GetShortNameUserBS(String.Format("{0}$", HelperClass1.UserInfo.userid))
            VendorcodeBS = VendorModel1.GetVendorCodeUserBS(String.Format("{0}$", HelperClass1.UserInfo.userid))
        End If

        GroupBox1.Enabled = CheckBox1.Checked

        DateTimePicker1.Value = CDate(String.Format("{0:yyyy-MM}-1", Date.Today.AddMonths(-12)))
        DateTimePicker2.Value = Date.Today
        ComboBox1.SelectedIndex = 0
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ShortNameHash.Clear()
        VendorNameHash.Clear()
        HashTable1.Clear()
        SortNameCriteriaList = New StringBuilder
        MyC = " where status not in  ('Closed')"
    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        initData()

        loadData()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim myform = New FormHelper(VendorcodeBS)
        myform.DataGridView1.Columns(0).DataPropertyName = "displayname"
        If myform.ShowDialog = DialogResult.OK Then
            For Each selecteditem As DataGridViewRow In myform.DataGridView1.SelectedRows
                VendorcodeBS.Position = selecteditem.Index
                Dim drv As DataRowView = VendorcodeBS.Current
                'TextBox2.Text = drv.Item("vendorname")
                If Not VendorNameHash.ContainsValue(drv.Item("vendorname")) Then
                    VendorNameHash.Add(drv.Item("vendorname"), drv.Item("vendorname"))
                End If
            Next


            Dim mylist As New StringBuilder
            VendornameCriteriaList = New StringBuilder
            Dim myvalue = VendorNameHash.Values.Cast(Of String).ToArray
            Array.Sort(myvalue)
            For i = 0 To myvalue.Count - 1
                If mylist.Length > 0 Then
                    mylist.Append(",")
                    VendornameCriteriaList.Append(",")
                End If
                mylist.Append(myvalue(i))
                VendornameCriteriaList.Append(String.Format("'{0}'", myvalue(i)))
            Next

            TextBox2.Text = mylist.ToString

        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'show helper

        Dim myform = New FormHelper(ShortNameBS)
        myform.DataGridView1.Columns(0).DataPropertyName = "shortname"
        If myform.ShowDialog = DialogResult.OK Then
            For Each selecteditem As DataGridViewRow In myform.DataGridView1.SelectedRows
                ShortNameBS.Position = selecteditem.Index
                Dim drv As DataRowView = ShortNameBS.Current
                If Not ShortNameHash.ContainsValue(drv.Item("shortname")) Then
                    ShortNameHash.Add(drv.Item("shortname"), drv.Item("shortname"))
                End If
            Next


        End If
        Dim mylist As New StringBuilder
        SortNameCriteriaList = New StringBuilder
        Dim myvalue = ShortNameHash.Values.Cast(Of String).ToArray
        Array.Sort(myvalue)
        For i = 0 To myvalue.Count - 1
            If mylist.Length > 0 Then
                mylist.Append(",")
                SortNameCriteriaList.Append(",")
            End If
            mylist.Append(myvalue(i))
            SortNameCriteriaList.Append(String.Format("'{0}'", myvalue(i)))
        Next

        TextBox1.Text = mylist.ToString
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ' Dim HashTable1 As New Hashtable
        Dim myform = New FormHelper(StatusBS)
        myform.DataGridView1.Columns(0).DataPropertyName = "status"
        If myform.ShowDialog = DialogResult.OK Then

            For Each selecteditem As DataGridViewRow In myform.DataGridView1.SelectedRows
                StatusBS.Position = selecteditem.Index
                Debug.Print("hello")
                Dim drv As DataRowView = StatusBS.Current
                If Not IsDBNull(drv.Item(0)) Then
                    If Not HashTable1.ContainsValue(drv.Item("status")) Then
                        HashTable1.Add(drv.Item("status"), drv.Item("status"))
                    End If
                End If
            Next


        End If
        Dim mylist As New StringBuilder
        StatusCriteriaList = New StringBuilder
        Dim myvalue = HashTable1.Values.Cast(Of String).ToArray
        Array.Sort(myvalue)
        For i = 0 To myvalue.Count - 1
            If mylist.Length > 0 Then
                mylist.Append(",")
                StatusCriteriaList.Append(",")
            End If
            mylist.Append(myvalue(i))
            StatusCriteriaList.Append(String.Format("'{0}'", myvalue(i)))
        Next

        TextBox3.Text = mylist.ToString
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        TxMode = TXModeEnum.RegularMode
        Dim myCriteria As New StringBuilder





        If TextBox1.Text.Length > 0 Then
            myCriteria.Append(String.Format("shortname in ({0})", SortNameCriteriaList))
        End If
        If TextBox2.Text.Length > 0 Then
            If myCriteria.Length > 0 Then
                myCriteria.Append(" and ")
            End If
            'myCriteria.Append(String.Format("vendorname = '{0}'", TextBox2.Text))
            myCriteria.Append(String.Format("shortname in (select shortname from vendor where vendorname in ({0}))", VendornameCriteriaList))
            SortNameCriteriaList.Clear()
            SortNameCriteriaList.Append(String.Format("select shortname from vendor where vendorname in ({0})", VendornameCriteriaList))
        End If

        If TextBox3.Text.Length > 0 Then
            If myCriteria.Length > 0 Then
                myCriteria.Append(" and ")
            End If
            myCriteria.Append(String.Format("status in  ({0})", StatusCriteriaList))
        Else
            If myCriteria.Length > 0 Then
                myCriteria.Append(" and ")
            End If
            myCriteria.Append(String.Format("status not in  ({0})", "'Closed'"))
        End If

        If CheckBox1.Checked Then
            If myCriteria.Length > 0 Then
                myCriteria.Append(" and ")
            End If
            myCriteria.Append(String.Format("to_char({0},'yyyyMM') >= '{1:yyyyMM}' and to_char({0},'yyyyMM') <= '{2:yyyyMM}'", typeofDate(ComboBox1.SelectedIndex), DateTimePicker1.Value.Date, DateTimePicker2.Value.Date))
        End If

        MyC = ""
        If myCriteria.Length > 0 Then
            MyC = String.Format(" where {0}", myCriteria.ToString)
        End If



        loadData()

    End Sub




    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        myController.save()
    End Sub



    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        runreport(Me, New EventArgs)
    End Sub

    Private Sub runreport(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim mymessage As String = String.Empty
        Dim sqlstr = myController.Model.GetQuery
        If sqlstr.Length = 0 Then
            MessageBox.Show("Load data first!")
            Exit Sub
        End If
        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("ActionPlan{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            myreport.Run(Me, e)
        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        Dim osheet = DirectCast(sender, Excel.Worksheet)
        osheet.Cells(1, 12).value = "Status (Not Started / On-going / Delay/ Closed/ Closed - Achieved/ Closed - Not Achieved)"
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        GroupBox1.Enabled = CheckBox1.Checked
    End Sub

    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        Dim mymessage As String = String.Empty
        Dim sqlstr = myController.Model.GetQueryALL

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("ActionPlanAll{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            myreport.Run(Me, e)
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'ToolStripButton4.PerformClick()
        TxMode = TXModeEnum.ValidateMode
        If HelperClass1.UserInfo.IsAdmin Then
            MyC = "where status = 'Closed' and validatorflag = false"
        Else
            MyC = String.Format(" where lower(validator) = '{0}' and status = 'Closed' and validatorflag = false", HelperClass1.UserId.ToLower)
        End If
        loadData()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If Not IsNothing(myController.BS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myController.BS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub FormSSSActionPlanFollowUP_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

    End Sub
End Class