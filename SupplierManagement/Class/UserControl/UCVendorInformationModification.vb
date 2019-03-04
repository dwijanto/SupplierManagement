Imports System.Text

Public Class UCVendorInformationModification
    Private CBS As New BindingSource
    Private VCBS As New BindingSource
    Private VBS As New BindingSource
    Private WithEvents drv As DataRowView
    Private ApplicantBS As New BindingSource
    Dim _suppliermodificationid As String

    Public Event RefreshInterface()
    Private DetailBS As New BindingSource
    Private ModificationTypeBS As BindingSource
    Private DocAttachmentBS As BindingSource
    Private dtdrv As DataRowView
    Public ToCompleteByDBDept As Boolean = False

    Public Sub DisableDataGridViewMenu()
        DataGridView1.ContextMenuStrip = Nothing
    End Sub
    Public Sub DataGridViewEndEdit()
        Me.Validate()
        'With DataGridView1
        '    Dim cm As CurrencyManager = CType(.BindingContext(.DataSource), CurrencyManager)
        '    cm.EndCurrentEdit()
        '    cm.Refresh()
        'End With
    End Sub



    Public Property SupplierModificationID As String
        Get
            Return _suppliermodificationid
        End Get
        Set(ByVal value As String)
            _suppliermodificationid = value
            TextBox1.Text = _suppliermodificationid
        End Set
    End Property


    Public Function ValidateField() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(ComboBox1, "")
        ErrorProvider1.SetError(ComboBox2, "")
        ErrorProvider1.SetError(ComboBox3, "")
        ErrorProvider1.SetError(DataGridView1, "")
        If DetailBS.Count = 0 Then
            ErrorProvider1.SetError(DataGridView1, "Create Item First.")
            myret = False
        End If

        If ComboBox1.SelectedIndex < 1 Then
            ErrorProvider1.SetError(ComboBox1, "Please select from the list")
            myret = False
        End If
        ToCompleteByDBDept = False
        Select Case TextBox11.Text
            Case 0
                drv.Row.Item("approvaldept") = DBNull.Value
                drv.Row.Item("approvaldept2") = DBNull.Value
                'drv.Row.Item("dbusername") = DBNull.Value
                drv.Row.Item("approvaldb") = DBNull.Value
                drv.Row.Item("fcusername") = DBNull.Value
                drv.Row.Item("approvalfc") = DBNull.Value
                drv.Row.Item("vpusername") = DBNull.Value
                drv.Row.Item("approvalvp") = DBNull.Value
                myret = True
            Case 1
                If IsNumeric(TextBox7.Text) Then
                    If TextBox7.Text < 5000000 Then
                        drv.Row.Item("approvalvp") = DBNull.Value
                        ToCompleteByDBDept = True
                    End If
                Else
                    drv.Row.Item("approvalvp") = DBNull.Value
                End If
                If IsNothing(ComboBox2.SelectedItem) Then
                    ErrorProvider1.SetError(ComboBox2, "Please select from the list")
                    myret = False
                End If
                If IsNothing(ComboBox3.SelectedItem) Then
                    ErrorProvider1.SetError(ComboBox3, "Please select from the list")
                    myret = False
                End If
            Case 2
                If IsNothing(ComboBox2.SelectedItem) Then
                    ErrorProvider1.SetError(ComboBox2, "Please select from the list")
                    myret = False
                End If
                If IsNothing(ComboBox3.SelectedItem) Then
                    ErrorProvider1.SetError(ComboBox3, "Please select from the list")
                    myret = False
                End If
                drv.Row.Item("approvalvp") = DBNull.Value
                ToCompleteByDBDept = True
            Case 3
                drv.Row.Item("approvaldept2") = DBNull.Value
                drv.Row.Item("dbusername") = DBNull.Value
                drv.Row.Item("approvaldb") = DBNull.Value
                drv.Row.Item("fcusername") = DBNull.Value
                drv.Row.Item("approvalfc") = DBNull.Value
                drv.Row.Item("vpusername") = DBNull.Value
                drv.Row.Item("approvalvp") = DBNull.Value
                If IsNothing(ComboBox2.SelectedItem) Then
                    ErrorProvider1.SetError(ComboBox2, "Please select from the list")
                    myret = False
                End If

        End Select

        'Check Detail


        Dim myDocAttachment As New Dictionary(Of Integer, String)
        Dim i As Integer = 0
        For Each drv As DataRowView In DocAttachmentBS.List
            If Not myDocAttachment.ContainsKey(drv.Row.Item("doctypeid")) Then
                myDocAttachment.Add(drv.Row.Item("doctypeid"), drv.Row.Item("doctype"))

            End If
        Next

        Dim mymodel As New ModificationTypeModel
        For Each mydrv As DataRowView In DetailBS.List
            Dim MyDoc As New Hashtable
            Dim sb As New StringBuilder
            MyDoc = (mymodel.GetDocTypeId(mydrv.Row.Item("fieldid")))
            For Each obj As System.Collections.DictionaryEntry In MyDoc
                If Not myDocAttachment.ContainsKey(obj.Key) Then
                    If sb.Length = 0 Then
                        sb.Append(String.Format("Missing Document : {0}", obj.Value))
                    Else
                        sb.Append(String.Format(",{0}", obj.Value))
                    End If
                    myret = False
                End If
            Next
            mydrv.Row.RowError = sb.ToString
        Next
        Return myret
    End Function

    Function GetRemarks() As Object
        Dim mymodel As New ModificationTypeModel
        Dim sb As New StringBuilder
        Dim ContainsBankInfo As Boolean = False
        For Each mydrv As DataRowView In DetailBS.List
            Dim result As String = mymodel.GetRemarks(mydrv.Row.Item("fieldid"))

            If result.Length > 0 Then
                If sb.Length = 0 Then
                    sb.Append(String.Format("Required attachments needed for supplier modification regarding to modify the different fields :{0}{0}", vbCrLf))
                End If
                Dim mydata = result.Split("|")
                If mydata(0) = 1 Then
                    If result.Length > 0 Then
                        sb.Append(String.Format("{0}{1}{1}", mydata(1), vbCrLf))
                    End If
                ElseIf mydata(0) = 2 Then
                    ContainsBankInfo = True
                End If
            End If
        Next
        If sb.Length > 0 Then
            If ContainsBankInfo Then
                sb.Append(String.Format("Bank Information{0}1. Formal letter from the supplier on the change{0}2)  Senior Purchasing Manager to call the usual contact (Sales Director/ Account Director or above) to check and confirm the name/ bank information change by phone.+ Email confirmation would be sent by Senior Purchasing Manager/ Purchasing Manager (when calling together) after the phone call. (Date of phone confirmation to be filled in above form){0}{0}", vbCrLf))
            End If
            sb.Append(String.Format("Signature for all formal letter requirements{0}1. Authorized signature (with clear name and title){0}2. Official company chop/stamp is required.{0}3. Proof of authorization signed (for position below Director) (e.g. letter/ meeting minutes from Board of Directors)", vbCrLf))

        End If

        Return sb.ToString
    End Function
    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged
        DisplayValidControl()
    End Sub

    Sub BindingControl(ByRef drv As DataRowView, ByRef detailbs As BindingSource, ByRef VBS As BindingSource, ByRef VCBS As BindingSource, ByRef CBS As BindingSource, ByRef ApplicantBS As BindingSource, ByRef ModificationTypeBS As BindingSource, ByRef ApprovalDeptBS As BindingSource, ByRef ApprovalDirectorBS As BindingSource, ByRef DocAttachmentBS As BindingSource)
        TextBox1.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        TextBox6.DataBindings.Clear()
        TextBox7.DataBindings.Clear()
        TextBox8.DataBindings.Clear()
        TextBox9.DataBindings.Clear()
        TextBox10.DataBindings.Clear()
        TextBox11.DataBindings.Clear()
        TextBox12.DataBindings.Clear()
        TextBox13.DataBindings.Clear()
        TextBox14.DataBindings.Clear()
        'TextBox15.DataBindings.Clear()
        'TextBox16.DataBindings.Clear()
        DateTimePicker1.DataBindings.Clear()

        ComboBox1.DataSource = ApplicantBS
        ComboBox1.DisplayMember = "username"
        ComboBox1.ValueMember = "username"

        ComboBox2.DataSource = ApprovalDeptBS
        ComboBox2.DisplayMember = "username"
        ComboBox2.ValueMember = "userid"


        ComboBox3.DataSource = ApprovalDirectorBS
        ComboBox3.DisplayMember = "username"
        ComboBox3.ValueMember = "userid"

        Me.ModificationTypeBS = ModificationTypeBS

        DataGridView1.AutoGenerateColumns = False
        With DirectCast(DataGridView1.Columns("FieldId"), DataGridViewComboBoxColumn)
            .DataSource = ModificationTypeBS
            .DisplayMember = "description"
            .ValueMember = "id"
        End With

        DataGridView1.DataSource = detailbs

        Me.DetailBS = detailbs
        Me.drv = drv
        Me.VBS = VBS
        Me.VCBS = VCBS
        Me.CBS = CBS
        Me.DocAttachmentBS = DocAttachmentBS

        TextBox1.DataBindings.Add(New Binding("Text", drv, "suppliermodificationid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", drv, "applicantname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("Text", drv, "vendorcodename", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox4.DataBindings.Add(New Binding("Text", drv, "familycode", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox5.DataBindings.Add(New Binding("Text", drv, "subfamilycode", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox6.DataBindings.Add(New Binding("Text", drv, "currency", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox7.DataBindings.Add(New Binding("Text", drv, "turnovervalue", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
        TextBox8.DataBindings.Add(New Binding("Text", drv, "yearreference", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox9.DataBindings.Add(New Binding("Text", drv, "ecoqualitycontactname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox10.DataBindings.Add(New Binding("Text", drv, "ecoqualitycontactemail", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox11.DataBindings.Add(New Binding("Text", drv, "sensitivitylevel", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox12.DataBindings.Add(New Binding("Text", drv, "dbusername", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox13.DataBindings.Add(New Binding("Text", drv, "vpusername", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox14.DataBindings.Add(New Binding("Text", drv, "fcusername", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", drv, "approvaldept", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox3.DataBindings.Add(New Binding("SelectedValue", drv, "approvaldept2", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        DateTimePicker1.DataBindings.Add(New Binding("Text", drv, "applicantdate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox1.ReadOnly = True
        GroupBox1.Visible = False

        DisplayValidControl()

    End Sub

    Sub SetEnabledControl(ByVal status As Boolean)
        'TextBox1.Enabled = False
        ComboBox1.Enabled = status
        TextBox3.Enabled = status
        TextBox4.Enabled = status
        TextBox5.Enabled = status
        TextBox6.Enabled = status
        TextBox7.Enabled = status
        TextBox8.Enabled = status
        TextBox9.Enabled = status
        TextBox10.Enabled = status
        TextBox11.Enabled = status
        TextBox12.Enabled = status
        TextBox13.Enabled = status
        TextBox14.Enabled = status
        ComboBox2.Enabled = status
        ComboBox3.Enabled = status
        'TextBox15.Enabled = status
        'TextBox16.Enabled = status
        DateTimePicker1.Enabled = status
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Add Contact  or Update Contact

        If Not IsNothing(VBS) Then
            Dim drvdt As DataRowView
            Dim mybutton As Button = DirectCast(sender, Button)
            If mybutton.Text = "New" Then
                drvdt = CBS.AddNew()
                drvdt.Item("contactname") = ""
                drvdt.Item("title") = ""
                drvdt.Item("email") = ""
                drvdt.Item("isecoqualitycontact") = True

            Else
                drvdt = CBS.Current
                Dim mydrv = VBS.Current
                mydrv.row.item("status") = False
            End If
            drvdt.EndEdit()
            'Add Contact       
            Dim myform = New FormDialogContact(CBS)
            myform.VCBS = VCBS
            myform.VBS = VBS
            If Not myform.ShowDialog = DialogResult.OK Then
                If mybutton.Text = "New" Then
                    CBS.RemoveCurrent()
                End If
            Else

                TextBox9.Text = drvdt.Item("contactname")
                TextBox10.Text = drvdt.Item("email")
                mybutton.Text = "Update"
                drv.EndEdit()
            End If
        End If

    End Sub

    Private Sub drv_PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Handles drv.PropertyChanged

        If TextBox10.Text = "" And TextBox9.Text = "" Then
            Button1.Text = "New"
        Else
            Button1.Text = "Update"
        End If
    End Sub

    Private Sub AddToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddToolStripMenuItem.Click
        drv.EndEdit()
        If Not IsNothing(dtdrv) Then
            DataGridView1.Invalidate()
            DataGridView1.EndEdit()
        End If


        dtdrv = DetailBS.AddNew
        'dtdrv = DetailBS.AddNew


    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = -1 Then
            DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
            DataGridView1.EndEdit()
        ElseIf DataGridView1.EditMode <> DataGridViewEditMode.EditOnEnter Then
            DataGridView1.EditMode = DataGridViewEditMode.EditOnEnter
            DataGridView1.BeginEdit(False)
        End If
    End Sub



    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim dtdrv As DataRowView = DetailBS.Current
        If Not IsNothing(dtdrv) Then


            dtdrv.Row.RowError = ""
            Try
                dtdrv.EndEdit()
                Select Case e.ColumnIndex
                    Case 0
                        Dim drv As DataRowView = ModificationTypeBS.Current

                        dtdrv.Row.Item("sensitivitylevel") = drv.Row.Item("sensitivitylevel")
                        TextBox11.Text = getSensitivityLevel().ToString

                    Case 1
                End Select
            Catch ex As Exception
                dtdrv.Row.RowError = ex.Message
            End Try
        End If

    End Sub

    Public Sub ShowSensitivityLevel()
        TextBox11.Text = getSensitivityLevel().ToString
    End Sub

    Public Function getSensitivityLevel() As Integer
        Dim sensitivityLevel As Integer
        'For i = 0 To DetailBS.List.Count - 1
        '    If i = 0 Then
        '        sensitivityLevel = drv.Row.Item("sensitivitylevel")
        '    Else
        '        If sensitivityLevel > drv.Row.Item("sensitivitylevel") Then
        '            sensitivityLevel = drv.Row.Item("sensitivitylevel")
        '        End If
        '    End If
        'Next
        For Each drv As DataRowView In DetailBS.List
            If sensitivityLevel = 0 Then
                sensitivityLevel = drv.Row.Item("sensitivitylevel")
            Else
                If sensitivityLevel > drv.Row.Item("sensitivitylevel") Then
                    If drv.Row.Item("sensitivitylevel") <> 0 Then sensitivityLevel = drv.Row.Item("sensitivitylevel")
                End If
            End If
        Next
        Return sensitivityLevel
    End Function




    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        'Dim drv As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
        Dim drv As DataRowView = DetailBS.Current
        drv.Row.RowError = e.Exception.Message
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If MessageBox.Show("Delete selected record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                DetailBS.RemoveAt(drv.Index)
            Next
        End If
        TextBox11.Text = getSensitivityLevel().ToString
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        'TextBox11_TextChanged(sender, e)
        DisplayValidControl()
    End Sub

    Private Sub DisplayValidControl()
        Panel1.Visible = ("1" = TextBox11.Text)
        Panel2.Visible = ("2" = TextBox11.Text)
        Panel3.Visible = ("3" = TextBox11.Text)
        GroupBox1.Visible = True

        For Each ctr As Control In GroupBox1.Controls
            ctr.Visible = True
        Next
        ToCompleteByDBDept = False

        Select Case TextBox11.Text
            Case 0
                'GroupBox1.Visible = False
                Label14.Visible = False
                Label16.Visible = False
                Label15.Visible = False
                'Label13.Visible = True
                ComboBox3.Visible = False
                'TextBox12.Visible = True
                TextBox13.Visible = False
                TextBox14.Visible = False
            Case 1
                If IsNumeric(TextBox7.Text) Then
                    If TextBox7.Text < 5000000 Then
                        Label14.Visible = False
                        TextBox13.Visible = False
                        ToCompleteByDBDept = True
                    End If
                Else
                    Label14.Visible = False
                    TextBox13.Visible = False
                    ToCompleteByDBDept = True
                End If
            Case 2
                Label14.Visible = False
                TextBox13.Visible = False
                ToCompleteByDBDept = True
            Case 3
                Label14.Visible = False
                Label16.Visible = False
                Label15.Visible = False
                Label13.Visible = False
                ComboBox3.Visible = False
                TextBox12.Visible = False
                TextBox13.Visible = False
                TextBox14.Visible = False

        End Select

        If TextBox10.Text = "" And TextBox9.Text = "" Then
            Button1.Text = "New"
        Else
            Button1.Text = "Update"
        End If
    End Sub




End Class
