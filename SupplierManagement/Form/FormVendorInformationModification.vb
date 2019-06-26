Imports System.Threading
'Imports SupplierManagement.PublicClass
Imports System.Text

Public Enum VendorInfoModiStatusEnum

    StatusNew = 1
    StatusResubmit = 2
    StatusValidatedbySPMCategoryLeader = 3
    StatusValidatedbyPurchasingDirector = 4
    StatusValidatedbyDatabaseManagementTeam = 5
    StatusValidatedbyFinancialController = 6
    StatusValidatedbyVPSourcingIndustry = 7
    StatusRejectedbySPMCategoryLeader = 8
    StatusRejectedbyPurchasingDirector = 9
    StatusRejectedbyDatabaseManagementTeam = 10
    StatusRejectedbyFinancialController = 11
    StatusRejectedbyVPSourcingIndustry = 12
    StatusCancelled = 13
    StatusCompleted = 14
End Enum

Public Class FormVendorInformationModification
    Dim HelperClass1 As HelperClass = HelperClass.getInstance

    Private _txEnum As TxEnum
    Private _p2 As Integer

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)

    Private _vendorcode As Long
    Private _vendorname As String
    Private myAdapter As New VendorInformationModificationAdapter
    Private myUserAdapter As New UserAdapter
    Private VendorFamilySubFamilyVCAdapter1 As New VendorFamilySubFamilyVCAdapter
    Private VendorCurrAdapter1 As New VendorCurrAdapter
    Private VendorTxAdapter1 As New VendorTxAdapter
    Private _referenceEnum As VendorInformationModificationReference
    Private ModificationTypeAdapter1 As New ModificationTypeAdapter
    Private UserAdapter1 As New UserAdapter

    Private TxTypeEnum As TxEnum
    Private HdId As Long
    Private drv As DataRowView
    Private DtlDrv As BindingSource
    Private yearref As Integer

    Private DocAttachmentBS As BindingSource

    Dim ApprovalDRV As DataRowView

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Sub New(ByVal TxTypeEnum As TxEnum, ByVal Id As Long)
        ' TODO: Complete member initialization 
        InitializeComponent()
        Me.TxTypeEnum = TxTypeEnum
        Me.HdId = Id
    End Sub

    Public Sub New(ByVal VendorDR As DataRow, ByVal yearref As Integer, ByVal ReferenceEnum As VendorInformationModificationReference, ByVal TxTypeEnum As TxEnum, ByVal ID As Long)
        InitializeComponent()
        _vendorcode = VendorDR.Item("vendorcode")
        _vendorname = VendorDR.Item("vendorname")
        _referenceEnum = ReferenceEnum
        Me.TxTypeEnum = TxTypeEnum
        Me.HdId = ID
        Me.yearref = yearref

        AddHandler UcVendorInformationModification1.RefreshInterface, AddressOf RefreshMYInterface

    End Sub

    Private Sub FormVendorInformationModification_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim DB2 As DataSet = myAdapter.GetChanges
        If Not IsNothing(DB2) Then
            If DB2.Tables(1).Rows.Count <> 0 Then
                Select Case MessageBox.Show("Save unsaved records?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                    Case Windows.Forms.DialogResult.Yes
                        If Me.validate Then
                            SaveRecord()
                        Else
                            e.Cancel = True
                        End If

                    Case Windows.Forms.DialogResult.Cancel
                        e.Cancel = True
                End Select
            End If

        End If
    End Sub


    Private Sub FormVendorInformationModification_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()

    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Try

            If myAdapter.loaddata(HdId, _vendorcode) Then
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
                        Dim ADBdrv As DataRowView = UserAdapter1.getApprovalDBVIMBS.Current
                        Dim AFCdrv As DataRowView = UserAdapter1.getApprovalFCBS.Current
                        Dim AVPdrv As DataRowView = UserAdapter1.getApprovalVPBS.Current


                        Select Case TxTypeEnum
                            Case TxEnum.NewRecord
                                drv = myAdapter.GetNewRecord
                                drv.Row.Item("creator") = HelperClass1.UserId
                                drv.EndEdit() 'To make sure the rowstate became added
                                'fill in selectec reference turnover
                                Dim budgetforecastAdapter1 = New BudgetForecastAdapter(_vendorcode, "shortname")

                                Dim latesperiod As Date = CDate(String.Format("{0}-1-1", yearref))
                                drv.Row.Item("yearreference") = yearref
                                Select Case _referenceEnum
                                    Case VendorInformationModificationReference.YBudget
                                        drv.Row.Item("turnovervalue") = budgetforecastAdapter1.GetBudgetForecast(latesperiod, _vendorcode, BudgetForecastAdapter.DataTypeEnum.BudgetEnum)
                                        drv.Row.Item("turnovertype") = VendorInformationModificationReference.YBudget
                                    Case VendorInformationModificationReference.YForecast48
                                        drv.Row.Item("turnovervalue") = budgetforecastAdapter1.GetBudgetForecast(latesperiod, _vendorcode, BudgetForecastAdapter.DataTypeEnum.Forecast4Plus8)
                                        drv.Row.Item("turnovertype") = VendorInformationModificationReference.YForecast48
                                    Case VendorInformationModificationReference.YForecast84
                                        drv.Row.Item("turnovervalue") = budgetforecastAdapter1.GetBudgetForecast(latesperiod, _vendorcode, BudgetForecastAdapter.DataTypeEnum.Forecast8Plus4)
                                        drv.Row.Item("turnovertype") = VendorInformationModificationReference.YForecast84
                                    Case VendorInformationModificationReference.Ymin1Actual
                                        drv.Row.Item("turnovervalue") = budgetforecastAdapter1.GetTurnover(yearref - 1, _vendorcode)
                                        drv.Row.Item("turnovertype") = VendorInformationModificationReference.Ymin1Actual
                                        drv.Row.Item("yearreference") = yearref - 1
                                End Select

                                drv.Row.Item("vendorcode") = _vendorcode
                                drv.Row.Item("vendorname") = _vendorname
                                drv.Row.Item("vendorcodename") = String.Format("{0} - {1}", _vendorcode, _vendorname)
                                Dim crcy As String = VendorCurrAdapter1.getLatesVendorCurr(_vendorcode)
                                drv.Row.Item("currency") = IIf(IsNothing(crcy), "USD", crcy)
                                drv.Row.Item("familycode") = VendorFamilySubFamilyVCAdapter1.GetFamilyCode(_vendorcode)
                                drv.Row.Item("subfamilycode") = VendorFamilySubFamilyVCAdapter1.GetSubFamilyCode(_vendorcode)
                                drv.Row.Item("subfamilycode") = VendorFamilySubFamilyVCAdapter1.GetSubFamilyCode(_vendorcode)

                                drv.Row.Item("ecoqualitycontactname") = VendorTxAdapter1.GetVendorEcoContactName(_vendorcode)
                                drv.Row.Item("ecoqualitycontactemail") = VendorTxAdapter1.GetVendorEcoContactEmail(_vendorcode)
                                Dim ismissing As Integer = 0
                                If drv.Row.Item("familycode") = "" Then
                                    ismissing = 1
                                    drv.Row.Item("familycode") = DBNull.Value
                                End If

                                If drv.Row.Item("subfamilycode") = "" Then
                                    ismissing = 2
                                    drv.Row.Item("subfamilycode") = DBNull.Value
                                End If

                                If IsDBNull(drv.Row.Item("familycode")) And IsDBNull(drv.Row.Item("subfamilycode")) Then ismissing = 3
                                drv.Row.Item("ismissing") = ismissing
                                drv.Row.Item("applicantdate") = Today.Date

                                drv.Row.Item("dbusername") = ADBdrv.Row.Item("username")
                                drv.Row.Item("approvaldb") = ADBdrv.Row.Item("userid")
                                drv.Row.Item("fcusername") = AFCdrv.Row.Item("username")
                                drv.Row.Item("approvalfc") = AFCdrv.Row.Item("userid")
                                drv.Row.Item("vpusername") = AVPdrv.Row.Item("username")
                                drv.Row.Item("approvalvp") = AVPdrv.Row.Item("userid")
                            Case TxEnum.UpdateRecord
                                drv = myAdapter.GetCurrentRecord

                                If IsDBNull(drv.Row.Item("dbusername")) Then
                                    drv.Row.Item("dbusername") = ADBdrv.Row.Item("username")
                                End If
                                If IsDBNull(drv.Row.Item("approvaldb")) Then
                                    drv.Row.Item("approvaldb") = ADBdrv.Row.Item("userid")
                                End If
                                If IsDBNull(drv.Row.Item("fcusername")) Then
                                    drv.Row.Item("fcusername") = AFCdrv.Row.Item("username")
                                End If
                                If IsDBNull(drv.Row.Item("approvalfc")) Then
                                    drv.Row.Item("approvalfc") = AFCdrv.Row.Item("userid")
                                End If
                                If IsDBNull(drv.Row.Item("vpusername")) Then
                                    drv.Row.Item("vpusername") = AVPdrv.Row.Item("username")
                                End If
                                If IsDBNull(drv.Row.Item("approvalvp")) Then
                                    drv.Row.Item("approvalvp") = AVPdrv.Row.Item("userid")
                                End If
                            Case TxEnum.ValidateRecord, TxEnum.HistoryRecord
                                drv = myAdapter.GetCurrentRecord
                            
                        End Select

                        'always set initial data, will be removed during saving.
                        'drv.Row.Item("dbusername") = ADBdrv.Row.Item("username")
                        'drv.Row.Item("approvaldb") = ADBdrv.Row.Item("userid")
                        'drv.Row.Item("fcusername") = AFCdrv.Row.Item("username")
                        'drv.Row.Item("approvalfc") = AFCdrv.Row.Item("userid")
                        'drv.Row.Item("vpusername") = AVPdrv.Row.Item("username")
                        'drv.Row.Item("approvalvp") = AVPdrv.Row.Item("userid")

                        'Binding Control

                        DtlDrv = myAdapter.getdtlBindingSource
                        DocAttachmentBS = myAdapter.GetDocumentBS
                        UcVendorInformationModification1.BindingControl(drv, DtlDrv, myAdapter.getVBS, myAdapter.getVCBS, myAdapter.getCBS, myUserAdapter.getApplicantBS, ModificationTypeAdapter1.GetModificationTypeBS, UserAdapter1.getApprovalDeptBS, UserAdapter1.getApprovalDirectorBS, DocAttachmentBS)
                        UcDocuments1.BindingControl(Me, DocAttachmentBS, myAdapter.getDocTypeBS, myAdapter.getFileSourceFullPath)
                        'Public Sub BindingControl(ByRef myform As Object, ByRef bs As BindingSource, ByVal DataTypeBS As BindingSource, ByVal FileSourceFullPath As String)
                        UcVendorInformationModification1.ShowSensitivityLevel()
                        'Handle Button Commit Submit Re-submit button enabled-disabled
                        ToolStripButton5.Visible = False 'Re-Submit
                        ToolStripButton6.Visible = False 'Validate
                        ToolStripButton3.Visible = False 'Cancel
                        ToolStripButton4.Visible = False 'submit
                        ToolStripButton2.Visible = False 'Commit
                        ToolStripButton7.Visible = False 'Reject
                        ToolStripButton8.Visible = False 'Complete

                        Select Case TxTypeEnum
                            Case TxEnum.NewRecord, TxEnum.UpdateRecord
                                If IsDBNull(drv.Row.Item("status")) Then
                                    ToolStripButton4.Visible = True
                                    ToolStripButton2.Visible = True
                                Else
                                    UcVendorInformationModification1.DisableDataGridViewMenu()
                                    UcDocuments1.DisableDataGridViewMenu()
                                End If
                            Case TxEnum.ValidateRecord
                                If drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyDatabaseManagementTeam Or
                                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyFinancialController Or
                                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyPurchasingDirector Or
                                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbySPMCategoryLeader Or
                                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyVPSourcingIndustry Then
                                    ToolStripButton5.Visible = True
                                ElseIf drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusNew Or
                                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusResubmit Or
                                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbyDatabaseManagementTeam Or
                                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbyPurchasingDirector Or
                                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbySPMCategoryLeader Or
                                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyVPSourcingIndustry Then

                                    If drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbySPMCategoryLeader And drv.Row.Item("sensitivitylevel") = 3 Then
                                        ToolStripButton8.Visible = True 'Complete
                                    Else
                                        ToolStripButton6.Visible = True
                                        ToolStripButton7.Visible = True
                                    End If
                                   
                                ElseIf drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbyFinancialController Then
                                    ToolStripButton8.Visible = UcVendorInformationModification1.ToCompleteByDBDept
                                    ToolStripButton6.Visible = Not (UcVendorInformationModification1.ToCompleteByDBDept)
                                    ToolStripButton7.Visible = Not (UcVendorInformationModification1.ToCompleteByDBDept)
                                ElseIf drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbyVPSourcingIndustry Then
                                    ToolStripButton8.Visible = True
                                End If
                                'Cancel button enabled-disable if admin
                                ToolStripButton3.Visible = HelperClass1.UserInfo.IsAdmin
                            Case TxEnum.HistoryRecord
                                UcVendorInformationModification1.DisableDataGridViewMenu()
                                UcDocuments1.DisableDataGridViewMenu()
                        End Select

                        'If IsDBNull(drv.Row.Item("status")) Then
                        '    ToolStripButton4.Visible = True
                        '    ToolStripButton2.Visible = True
                        'Else
                        '    If drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyDatabaseManagementTeam Or
                        '       drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyFinancialController Or
                        '       drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyPurchasingDirector Or
                        '       drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbySPMCategoryLeader Or
                        '       drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyVPSourcingIndustry Then
                        '        ToolStripButton5.Visible = True
                        '    ElseIf drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusNew Or
                        '        drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusResubmit Or
                        '        drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbyDatabaseManagementTeam Or
                        '        drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbyFinancialController Or
                        '        drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbyPurchasingDirector Or
                        '        drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbySPMCategoryLeader Or
                        '        drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyVPSourcingIndustry Then
                        '        ToolStripButton6.Visible = True
                        '    End If
                        '    'Commit button after submitted cannot be modified
                        '    'ToolStripButton2.Visible = False
                        '    'Cancel button enabled-disable if admin
                        '    ToolStripButton3.Visible = HelperClass1.UserInfo.IsAdmin
                        'End If

                        DataGridView5.AutoGenerateColumns = False
                        DataGridView5.DataSource = myAdapter.VendorInfoModiActionBS

                        'UcVendorInformationModification1.ShowSensitivityLevel()
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

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            'ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Function SaveRecord() As Boolean
        Dim myret As Boolean = False
        UcVendorInformationModification1.DataGridViewEndEdit()

        drv.EndEdit()
        DtlDrv.EndEdit()
        'If Not UcDocuments1.validate() Then
        '    ProgressReport(1, "Error:: Please check details.")
        '    Exit Sub
        'End If
        Dim mysb As New StringBuilder


        ProgressReport(1, "")

        'show message when 
        If Me.validate And UcDocuments1.validate Then
            If Not IsDBNull(drv.Row.Item("familycode")) Then
                If drv.Row.Item("familycode") = "" Then
                    mysb.Append(String.Format("Vendor Family Code is blank.{0}", vbCrLf))
                End If
            Else
                mysb.Append(String.Format("Vendor Family Code is blank.{0}", vbCrLf))
            End If

            If Not IsDBNull(drv.Row.Item("subfamilycode")) Then
                If drv.Row.Item("subfamilycode") = "" Then
                    mysb.Append(String.Format("Vendor Sub Family Code is blank.{0}", vbCrLf))
                End If
            Else
                mysb.Append(String.Format("Vendor Sub Family Code is blank.{0}", vbCrLf))
            End If

            If Not IsDBNull(drv.Row.Item("turnovervalue")) Then
                If drv.Row.Item("turnovervalue") = 0 Then
                    mysb.Append(String.Format("Reference Yearly turnover(USD) is 0.{0}", vbCrLf))
                End If
            Else
                mysb.Append(String.Format("Reference Yearly turnover(USD) is 0.{0}", vbCrLf))
            End If

            If Not IsDBNull(drv.Row.Item("ecoqualitycontactname")) Then
                If drv.Row.Item("ecoqualitycontactname") = "" Then
                    mysb.Append(String.Format("Eco Quality Contact Name is blank.{0}", vbCrLf))
                End If
            Else
                mysb.Append(String.Format("Eco Quality Contact Name is blank.{0}", vbCrLf))
            End If

            If Not IsDBNull(drv.Row.Item("ecoqualitycontactemail")) Then
                If drv.Row.Item("ecoqualitycontactemail") = "" Then
                    mysb.Append(String.Format("Eco Quality Contact Email is blank.{0}", vbCrLf))
                End If
            Else
                mysb.Append(String.Format("Eco Quality Contact Email is blank.{0}", vbCrLf))
            End If

            If mysb.Length > 0 Then
                mysb.Append(String.Format("Please inform Administrator."))
                MessageBox.Show(mysb.ToString, "Blank Value", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            myAdapter.save()
            If TxTypeEnum = TxEnum.NewRecord Then
                UcVendorInformationModification1.SupplierModificationID = myAdapter.getsuppliermodificationid
            End If
            myret = True
        Else
            myret = False
            ProgressReport(1, "Error:: Please check details.")
        End If
        Return myret
    End Function

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Me.validate()
        SaveRecord()
    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        myret = UcVendorInformationModification1.ValidateField()

        Return myret
    End Function

    Private Sub RefreshInterface()
        Throw New NotImplementedException
    End Sub

    Private Sub RefreshMYInterface()
        Me.Invalidate()
    End Sub


    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim myPrint = New ExportVendorInformationModification(Me, drv, DtlDrv)
        myPrint.GenerateExcel()
    End Sub


    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click

        If IsDBNull(drv.Row.Item("status")) Then
            If MessageBox.Show("Do you want to submit this record?", "Submit", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then

                drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusNew

                Dim approvaldrv As DataRowView = myAdapter.VendorInfoModiActionBS.AddNew               
                approvaldrv.Row.Item("status") = drv.Item("status")
                approvaldrv.Row.Item("statusname") = "New"
                approvaldrv.Row.Item("modifiedby") = HelperClass1.UserId
                approvaldrv.Row.Item("vendorinfomodiid") = drv.Item("id")
                approvaldrv.EndEdit()
                Logger.log(String.Format("** Submit {0}**", HelperClass1.UserId))
                If SaveRecord() Then SendEmail(drv)
            End If

        Else
            If drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusNew Then
                MessageBox.Show("Record already submitted.")
            End If
        End If

    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        'If drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusResubmit Then
        '    MessageBox.Show("Record already re-submitted")
        'Else
        If Not IsNothing(ApprovalDRV) Then
            MessageBox.Show("Nothing to do.")
        Else
            If MessageBox.Show("Do you want to re-submit this record?", "Re Submit", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusResubmit

                ApprovalDRV = myAdapter.VendorInfoModiActionBS.AddNew
                approvaldrv.Row.Item("status") = drv.Item("status")
                approvaldrv.Row.Item("modifiedby") = HelperClass1.UserId
                approvaldrv.Row.Item("vendorinfomodiid") = drv.Item("id")
                approvaldrv.EndEdit()
                If SaveRecord() Then SendEmail(drv)
                
            End If
        End If

    End Sub

    Private Sub TabControl1_Selected(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlEventArgs) Handles TabControl1.Selected
        If TabControl1.SelectedTab.Text = "Documents" Then
            Dim myRemarks = UcVendorInformationModification1.GetRemarks()
            UcDocuments1.RichTextBox1.Text = myRemarks
        End If
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If IsNothing(ApprovalDRV) Then
            If MessageBox.Show("Do you want to cancel this record?", "Cancel", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                Dim remarks As String = InputBox("Please input some comment.")

                drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusCancelled

                ApprovalDRV = myAdapter.VendorInfoModiActionBS.AddNew
                ApprovalDRV.Row.Item("status") = drv.Item("status")
                ApprovalDRV.Row.Item("modifiedby") = HelperClass1.UserId
                ApprovalDRV.Row.Item("vendorinfomodiid") = drv.Item("id")
                ApprovalDRV.Row.Item("remark") = remarks
                ApprovalDRV.EndEdit()
                If SaveRecord() Then SendEmail(drv)
            End If
        Else
            MessageBox.Show("Nothing to do.")
        End If
       
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        If IsNothing(ApprovalDRV) Then
            If MessageBox.Show("Do you want to validate this record?", "Validate", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                Dim remarks As String = InputBox("Please input some comment.")
                drv.Row.Item("status") = IIf(drv.Row.Item("status") = 1, 2, 1) + drv.Row.Item("status")
                ApprovalDRV = myAdapter.VendorInfoModiActionBS.AddNew
                ApprovalDRV.Row.Item("status") = drv.Item("status")
                ApprovalDRV.Row.Item("modifiedby") = HelperClass1.UserId
                ApprovalDRV.Row.Item("vendorinfomodiid") = drv.Item("id")
                ApprovalDRV.Row.Item("remark") = remarks
                ApprovalDRV.EndEdit()
                If SaveRecord() Then SendEmail(drv)

            End If
        Else
            MessageBox.Show("Nothing to do.")
        End If
        
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        If IsNothing(ApprovalDRV) Then
            If MessageBox.Show("Do you want to reject this record?", "Reject", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then

                Dim remarks As String = InputBox("Please input some comment.")
                drv.Row.Item("status") = IIf(drv.Row.Item("status") = 1, 7, 6) + drv.Row.Item("status")

                ApprovalDRV = myAdapter.VendorInfoModiActionBS.AddNew
                ApprovalDRV.Row.Item("status") = drv.Item("status")
                ApprovalDRV.Row.Item("modifiedby") = HelperClass1.UserId
                ApprovalDRV.Row.Item("vendorinfomodiid") = drv.Item("id")
                ApprovalDRV.Row.Item("remark") = remarks
                ApprovalDRV.EndEdit()

               If SaveRecord() Then SendEmail(drv)
            End If
        Else
            MessageBox.Show("Nothing to do.")
        End If
    End Sub

    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        If IsNothing(ApprovalDRV) Then
            If MessageBox.Show("Do you want to complete this record?", "Complete", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then

                Dim remarks As String = InputBox("Please input some comment.")
                drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusCompleted

                ApprovalDRV = myAdapter.VendorInfoModiActionBS.AddNew
                ApprovalDRV.Row.Item("status") = drv.Item("status")
                ApprovalDRV.Row.Item("modifiedby") = HelperClass1.UserId
                ApprovalDRV.Row.Item("vendorinfomodiid") = drv.Item("id")
                ApprovalDRV.Row.Item("remark") = remarks
                ApprovalDRV.EndEdit()

                If SaveRecord() Then SendEmail(drv)
            End If
        Else
            MessageBox.Show("Nothing to do.")
        End If
    End Sub

    Public Sub SendEmail(ByVal drv As DataRowView)
        Try
            Dim SPMUserName = drv.Row.Item("spmusername")
            Dim PDUserName = drv.Row.Item("pdusername")
            Dim DBUserName = drv.Row.Item("dbusername")
            Dim FCUserName = drv.Row.Item("fcusername")
            Dim StatusName As String = String.Empty
            Dim SendTo As String = String.Empty
            Dim SendToName As String = String.Empty


            Dim cc As String = String.Empty
            ' 1::New
            ' 2::Re-submit
            ' 3::Validated by SPM/ Category Leader
            ' 4::Validated by Purchasing Director
            ' 5::Validated by Database Management Team
            ' 6::Validated by Financial Controller
            ' 7::Validated by VP Sourcing/Industry 
            ' 8::Rejected by SPM/Category Leader
            ' 9::Rejected by Purchasing Director
            '10::Rejected by Database Management Team
            '11::Rejected by Financial Controller
            '12::Rejected by Financial Controller
            '13::Cancelled
            '14::Completed 


            'check validator and sensitivity level then assign the email to
            Dim SendTos As String() = Nothing
            Dim CCs As String() = Nothing
            Dim Creators As String() = Nothing
            Dim CCDBdrv As DataRowView = UserAdapter1.getCCDBBS.Current
            Dim CCDBListBS As BindingSource = UserAdapter1.getApprovalDBListBS
            Dim ccdblistsb As New StringBuilder
            For Each edrv In CCDBListBS.List
                If ccdblistsb.Length > 0 Then ccdblistsb.Append(";")
                ccdblistsb.Append(edrv.Row.Item("email"))
            Next
            Dim senddblist As Boolean = False
            Select Case drv.Item("status")
                Case 1
                    StatusName = "NEW"
                    Select Case drv.Row.Item("sensitivitylevel")
                        Case 0
                            'SendTos = drv.Row.Item("approvaldb").ToString.Split("\")
                            'SendToName = drv.Row.Item("dbusername")
                            senddblist = True
                        Case Else
                            SendTos = drv.Row.Item("approvaldept").ToString.Split("\")
                            SendToName = drv.Row.Item("spmusername")
                    End Select
                Case 2
                    StatusName = "Re-Submit"
                    Select Case drv.Row.Item("sensitivitylevel")
                        Case 0
                            SendTos = drv.Row.Item("approvaldb").ToString.Split("\")
                            SendToName = drv.Row.Item("dbusername")


                        Case Else
                            SendTos = drv.Row.Item("approvaldept").ToString.Split("\")
                            SendToName = drv.Row.Item("spmusername")
                    End Select
                Case 3
                    StatusName = "Validated by SPM/ Category Leader"
                    Select Case drv.Item("sensitivitylevel")
                        Case 3
                            'SendTos = drv.Row.Item("approvaldb").ToString.Split("\")
                            'SendToName = drv.Row.Item("dbusername")                       
                            senddblist = True
                        Case Else
                            SendTos = drv.Row.Item("approvaldept2").ToString.Split("\")
                            SendToName = drv.Row.Item("pdusername")
                    End Select
                Case 4
                    StatusName = "Validated by Purchasing Director"
                    SendTos = drv.Row.Item("approvaldb").ToString.Split("\")
                    SendToName = drv.Row.Item("dbusername")
                Case 5
                    StatusName = "Validated by Database Management Team"
                    SendTos = drv.Row.Item("approvalfc").ToString.Split("\")
                    SendToName = drv.Row.Item("fcusername")
                Case 6
                    StatusName = "Validated by Financial Controller"
                    'Check Sensitivity Level
                    If drv.Row.Item("turnovervalue") > 5000000 And drv.Item("sensitivitylevel") = 1 Then
                        SendTos = drv.Row.Item("approvalvp").ToString.Split("\")
                        SendToName = drv.Row.Item("vpusername")
                    Else
                        SendTos = drv.Row.Item("approvaldb").ToString.Split("\")
                        SendToName = drv.Row.Item("dbusername")
                        CCs = CCDBdrv.Row.Item("userid").ToString.Split("\")
                        cc = String.Format("{0}@groupeseb.com;", CCs(1))
                    End If
                Case 7
                    StatusName = "Validated by VP Sourcing/Industry"
                    SendTos = drv.Row.Item("approvaldb").ToString.Split("\")
                    SendToName = drv.Row.Item("dbusername")
                    CCs = CCDBdrv.Row.Item("username").ToString.Split("\")
                    cc = String.Format("{0}@groupeseb.com;", CCs(1))
            End Select

            If drv.Item("status") >= 1 And drv.Item("status") <= 7 Then
                If senddblist Then
                    SendTo = ccdblistsb.ToString
                    SendToName = "All"
                Else
                    SendTo = String.Format("{0}@groupeseb.com", SendTos(1))
                End If
            End If

            If drv.Item("status") >= 8 And drv.Item("status") <= 13 Then
                Select Case drv.Item("status")
                    Case 8
                        StatusName = "Rejected by SPM/Category Leader"
                    Case 9
                        StatusName = "Rejected by Purchasing Director"
                    Case 10
                        StatusName = "Rejected by Database Management Team"
                    Case 11
                        StatusName = "Rejected by Financial Controller"
                    Case 12
                        StatusName = "Rejected by Financial Controller"
                    Case 13
                        StatusName = "Cancelled"
                End Select
                Creators = drv.Row.Item("creator").ToString.Split("\")

                'SendTo = String.Format("{0}@groupeseb.com;{1};{2}", drv.Item("creator"), drv.Item("appemail"), "afok@groupeseb.com")
                SendTo = String.Format("{0}@groupeseb.com;{1};{2}", Creators(1), drv.Item("appemail"), "afok@groupeseb.com")
                SendToName = "All"
            End If


            If drv.Item("status") = VendorInfoModiStatusEnum.StatusCompleted Then
                StatusName = "Completed"
                Creators = drv.Row.Item("creator").ToString.Split("\")
                SendTo = String.Format("{0}@groupeseb.com;{1};{2}", Creators(1), drv.Item("appemail"), "afok@groupeseb.com;ttom@groupeseb.com")
                cc = String.Format("{0};", ccdblistsb.ToString)
                SendToName = "All"
            End If

            Dim myEmail As VIMEmail = New VIMEmail
            Logger.log(String.Format("SendTo: {0}, SendTo Name: {1}, StatusName: {2}", SendTo, SendToName, StatusName))
            If Not myEmail.Execute(SendTo, SendToName, StatusName, drv, cc) Then
                Logger.log(String.Format("Error Message: {0}", myEmail.errorMessage))
            Else
                Logger.log("Email Sent")
            End If
        Catch ex As Exception
            Logger.log(ex.Message)
            MessageBox.Show(ex.Message)
        End Try
        

    End Sub

End Class