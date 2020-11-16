Imports System.Threading
Imports System.Text

Public Class FormNewVendor
    Dim HelperClass1 As HelperClass = HelperClass.getInstance
    Private Property TxTypeEnum As TxEnum
    Private Property Id As Long
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Private myAdapter As New NewVendorAdapter
    Dim FamilyVCController As New FamilyVCAdapter
    Dim SubFamilyVCController As New SubFamilyVCAdapter
    Dim PaymentTermController As New PaymentTermAdapter
    Dim VendorStatusController As New FormVendorStatus001
    Dim VendorIndirectFamilyController As New FormVendorIndirectFamily
    Dim SupplierGSMController As New SupplierGSMAdapter
    Dim MasterGroupController As New FormMasterGroup
    Dim FamilyController As New FamilyAdapter
    Dim PMController As New PMAdapter
    Dim TechnologyController As New FormMasterTechnology
    Dim ParamDTLController As New ParamDTLAdapter
    Private myUserAdapter As New UserAdapter
    Private drv As DataRowView
    Private DocAttachmentBS As BindingSource
    Private SITxAttachmentBS As BindingSource
    Private ZetolAttachmentBS As BindingSource
    Private ZetolMasterBS As BindingSource

    'Public Shared Event Submit()
    'Public Event Submit()
    'Public Shared Event RefreshData()

    Public Shared Event RefreshDataGridView(ByVal obj As Object, ByVal e As EventArgs)

    Dim ApprovalDRV As DataRowView

    Sub New(ByVal TxTypeEnum As TxEnum, ByVal Id As Long)
        ' TODO: Complete member initialization 
        InitializeComponent()
        Me.TxTypeEnum = TxTypeEnum
        Me.Id = Id
    End Sub

    Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub


    Private Sub FormNewVendor_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Try

            If myAdapter.loaddata(Id) Then


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

                        'Refresh UCNewVendor
                        Dim ApplicantNameBS = myUserAdapter.getApplicantBS
                        Dim ApprovalDept1BS = myUserAdapter.getApprovalDeptBLBS
                        Dim ApprovalDept2BS = myUserAdapter.getApprovalDirectorBLBS
                        Dim FinanceBS = myUserAdapter.getApprovalFCBS


                        Dim FamilyCodeBS = FamilyVCController.GetFamilyVCBS
                        Dim SubFamilyCodeBS = SubFamilyVCController.getSubFamilyVCBS
                        Dim PaymentTermBS = PaymentTermController.GetPaymentTermBS

                        Dim VendorStatusBS = VendorStatusController.getStatusBS
                        Dim VendorTypeBS = VendorStatusController.getProductTypeBS
                        Dim VendorIndFamBS = VendorIndirectFamilyController.getIndirectFamilyBS
                        Dim GSMBS = SupplierGSMController.getGSMBS
                        Dim GroupSBUBS = MasterGroupController.getGroupSBUBS
                        Dim FamilyBS = FamilyController.getFamilyBS
                        Dim PMBS = PMController.getPMBS
                        Dim TechnologyBS1 = TechnologyController.getTechnologyBS
                        Dim TechnologyBS2 = TechnologyController.getTechnologyBS
                        Dim TechnologyBS3 = TechnologyController.getTechnologyBS
                        Dim TechnologyBS4 = TechnologyController.getTechnologyBS
                        Dim CurrencyBS = ParamDTLController.GetParamDTLBS("currency")

                        If TxTypeEnum = TxEnum.NewRecord Then
                            drv = myAdapter.GetNewRecord
                            drv.Item("vendorstatus") = 1 'Active Supplier
                            drv.Item("approvalfc") = DirectCast(FinanceBS.Current, DataRowView).Row.Item("userid")
                            drv.Item("applicantdate") = Date.Today
                            drv.Row.Item("creator") = HelperClass1.UserId
                            drv.Row.Item("currency") = "USD"
                            drv.EndEdit() 'Need this one for relation other table
                        Else
                            drv = myAdapter.GetCurrentRecord
                        End If

                        TabControl1.TabPages.Remove(TabPage2)
                        UcNewVendor1.BindingControls(TxTypeEnum, drv, ApplicantNameBS, FamilyCodeBS, SubFamilyCodeBS, PaymentTermBS, VendorStatusBS, VendorTypeBS,
                                                     VendorIndFamBS, GSMBS, GroupSBUBS, FamilyBS, PMBS, TechnologyBS1, TechnologyBS2, TechnologyBS3, TechnologyBS4,
                                                     ApprovalDept1BS, ApprovalDept2BS, FinanceBS, CurrencyBS)

                        DocAttachmentBS = myAdapter.GetDocumentBS
                        SITxAttachmentBS = myAdapter.GetSITxAttachmentBS
                        ZetolAttachmentBS = myAdapter.GetZetolAttachmentBS
                        ZetolMasterBS = myAdapter.GetZetolMasterBS

                        UcDocuments1.BindingControl(Me, DocAttachmentBS, myAdapter.getDocTypeBS, myAdapter.getFileSourceFullPath)

                        UcnsDocumentHeader1.BindingControl(Me, DocAttachmentBS, SITxAttachmentBS, ZetolAttachmentBS, ZetolMasterBS, myAdapter.getDocTypeBS, myAdapter.getFileSourceFullPath, myAdapter.GetPaymentTermBS, myAdapter.GetLevelBS)
                        'UcnsDocumentHeader1.BindingControl(Me, DocAttachmentBS, SITxAttachmentBS, myAdapter.getDocTypeBS, myAdapter.getFileSourceFullPath)

                        showValidToolbar(TxTypeEnum, drv)
                        DataGridView5.AutoGenerateColumns = False
                        DataGridView5.DataSource = myAdapter.VendorInfoModiActionBS

                        TabControl1.TabPages.Remove(TabPage2)
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


    Private Sub showValidToolbar(ByVal txEnum As TxEnum, ByVal drv As DataRowView)
        ToolStripButtonCommit.Visible = False
        ToolStripButtonSubmit.Visible = False
        ToolStripButtonReSubmit.Visible = False
        ToolStripButtonValidate.Visible = False
        ToolStripButtonReject.Visible = False
        ToolStripButtonStatusCancelled.Visible = False
        ToolStripButtonComplete.Visible = False
        ToolStripButtonConfirmation.Visible = False

        'ToolStripButton5.Visible = False 'Re-Submit
        'ToolStripButton6.Visible = False 'Validate
        'ToolStripButton3.Visible = False 'Cancel
        'ToolStripButton4.Visible = False 'submit
        'ToolStripButton2.Visible = False 'Commit
        'ToolStripButton7.Visible = False 'Reject
        'ToolStripButton8.Visible = False 'Complete
        Select Case txEnum
            Case SupplierManagement.TxEnum.NewRecord, SupplierManagement.TxEnum.UpdateRecord
                ToolStripButtonCommit.Visible = True
                If IsDBNull(drv.Item("status")) Then
                    ToolStripButtonSubmit.Visible = True
                Else
                    ToolStripButtonSubmit.Visible = True
                End If
            Case SupplierManagement.TxEnum.ValidateRecord
                If drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyDatabaseManagementTeam Or
                                     drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyFinancialController Or
                                     drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyPurchasingDirector Or
                                     drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbySPMCategoryLeader Or
                                     drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyVPSourcingIndustry Then
                    ToolStripButtonReSubmit.Visible = True
                ElseIf drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusNew Or
                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusResubmit Or
                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbyDatabaseManagementTeam Or
                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbyPurchasingDirector Or
                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbySPMCategoryLeader Or
                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusRejectedbyVPSourcingIndustry Then

                    'If drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbySPMCategoryLeader Then
                    '    'ToolStripButtonComplete.Visible = True
                    '    ToolStripButtonValidate.Visible = True
                    '    ToolStripButtonReject.Visible = True
                    'Else
                    ToolStripButtonValidate.Visible = True
                    ToolStripButtonReject.Visible = True
                    'End If
                ElseIf drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusValidatedbyFinancialController Then
                    ToolStripButtonConfirmation.Visible = True
                ElseIf drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusCreatedInSAP Then
                    ToolStripButtonComplete.Visible = True
                End If
                'Cancel button enabled-disable if admin
                ToolStripButtonStatusCancelled.Visible = HelperClass1.UserInfo.IsAdmin
            Case SupplierManagement.TxEnum.HistoryRecord
                UcNewVendor1.HistoryMode()
                UcDocuments1.HistoryMode()
        End Select
        ToolStripButtonCommit.Visible = HelperClass1.UserInfo.IsAdmin
    End Sub

    Private Sub ToolStripButtonCommit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonCommit.Click
        'If Me.Validate Then
        '    myAdapter.save()
        '    If TxTypeEnum = TxEnum.NewRecord Then
        '        UcNewVendor1.SupplierModificationID = myAdapter.getsuppliermodificationid
        '    End If
        'End If
        'Me.Validate()
        Try
            SaveRecord()
            RaiseEvent RefreshDataGridView(Me, New EventArgs)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Overloads Function Validate() As Boolean
        Dim myret As Boolean

        myret = UcNewVendor1.Validate

        Return myret
    End Function

   

    Private Function SaveRecord() As Boolean
        Dim myret As Boolean = False
        drv.EndEdit()
        Dim mysb As New StringBuilder
        ProgressReport(1, "")

        'show message when 
        If Me.Validate And UcDocuments1.validate Then
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

            'If Not IsDBNull(drv.Row.Item("turnovervalue")) Then
            '    If drv.Row.Item("turnovervalue") = 0 Then
            '        mysb.Append(String.Format("Reference Yearly turnover(USD) is 0.{0}", vbCrLf))
            '    End If
            'Else
            '    mysb.Append(String.Format("Reference Yearly turnover(USD) is 0.{0}", vbCrLf))
            'End If

            'If Not IsDBNull(drv.Row.Item("ecoqualitycontactname")) Then
            '    If drv.Row.Item("ecoqualitycontactname") = "" Then
            '        mysb.Append(String.Format("Eco Quality Contact Name is blank.{0}", vbCrLf))
            '    End If
            'Else
            '    mysb.Append(String.Format("Eco Quality Contact Name is blank.{0}", vbCrLf))
            'End If

            'If Not IsDBNull(drv.Row.Item("ecoqualitycontactemail")) Then
            '    If drv.Row.Item("ecoqualitycontactemail") = "" Then
            '        mysb.Append(String.Format("Eco Quality Contact Email is blank.{0}", vbCrLf))
            '    End If
            'Else
            '    mysb.Append(String.Format("Eco Quality Contact Email is blank.{0}", vbCrLf))
            'End If

            If mysb.Length > 0 Then
                mysb.Append(String.Format("Please inform Administrator."))
                MessageBox.Show(mysb.ToString, "Blank Value", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
           
            If myAdapter.save() Then
                UcNewVendor1.SupplierModificationID = myAdapter.getsuppliermodificationid
                myret = True
            End If

            
        Else
            myret = False
            ProgressReport(1, "Error:: Please check details.")
        End If
        Return myret
    End Function

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
            '15::VendorCode Created In SAP

            'VendorInfoModiStatusEnum
            'StatusNew = 1
            'StatusResubmit = 2
            'StatusValidatedbySPMCategoryLeader = 3
            'StatusValidatedbyPurchasingDirector = 4
            'StatusValidatedbyDatabaseManagementTeam = 5
            'StatusValidatedbyFinancialController = 6
            'StatusValidatedbyVPSourcingIndustry = 7
            'StatusRejectedbySPMCategoryLeader = 8
            'StatusRejectedbyPurchasingDirector = 9
            'StatusRejectedbyDatabaseManagementTeam = 10
            'StatusRejectedbyFinancialController = 11
            'StatusRejectedbyVPSourcingIndustry = 12
            'StatusCancelled = 13
            'StatusCompleted = 14
            'StatusVendorCodeCreatedInSAP = 15


            'check validator and sensitivity level then assign the email to
            Dim SendTos As String() = Nothing
            Dim CCs As String() = Nothing
            Dim Creators As String() = Nothing
            Dim CCDBdrv As DataRowView = myUserAdapter.getCCDBBS.Current
            Dim CCDBListBS As BindingSource = myUserAdapter.getApprovalDBListBS
            Dim ccdblistsb As New StringBuilder
            For Each edrv In CCDBListBS.List
                If ccdblistsb.Length > 0 Then ccdblistsb.Append(";")
                ccdblistsb.Append(edrv.Row.Item("email"))
            Next
            Dim senddblist As Boolean = False
            Select Case drv.Item("status")
                Case 1
                    StatusName = "NEW"
                    SendTos = drv.Row.Item("approvaldept").ToString.Split("\")
                    SendToName = drv.Row.Item("spmusername")
                Case 2
                    StatusName = "Re-Submit"
                    SendTos = drv.Row.Item("approvaldept").ToString.Split("\")
                    SendToName = drv.Row.Item("spmusername")                    
                Case 3
                    StatusName = "Validated by SPM/ Category Leader"
                    SendTos = drv.Row.Item("approvaldept2").ToString.Split("\")
                    SendToName = drv.Row.Item("pdusername")
                Case 4
                    StatusName = "Validated by Purchasing Director"
                    SendTos = drv.Row.Item("approvalfc").ToString.Split("\")
                    SendToName = drv.Row.Item("fcusername")
                Case 5
                    StatusName = "Validated by Database Management Team"                   
                    Creators = drv.Row.Item("creator").ToString.Split("\")
                    SendTo = String.Format("{0}@groupeseb.com;{1}", Creators(1), "afok@groupeseb.com")
                    SendToName = "All"
                Case 6
                    StatusName = "Validated by Financial Controller"          
                    senddblist = True

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
                        StatusName = "Rejected by VP SourcingIndustry"
                    Case 13
                        StatusName = "Cancelled"
                End Select
                Creators = drv.Row.Item("creator").ToString.Split("\")

                SendTo = String.Format("{0}@groupeseb.com;{1};{2}", Creators(1), drv.Item("appemail"), "afok@groupeseb.com")
                SendToName = "All"
            End If


            If drv.Item("status") = VendorInfoModiStatusEnum.StatusCompleted Then
                StatusName = "Completed"
                Creators = drv.Row.Item("creator").ToString.Split("\")
                SendTo = String.Format("{0}@groupeseb.com;{1};{2}", Creators(1), drv.Item("appemail"), "afok@groupeseb.com;ttom@groupeseb.com")
                'cc = String.Format("{0};", ccdblistsb.ToString)
                SendToName = "All"
            End If

            If drv.Item("status") = VendorInfoModiStatusEnum.StatusCreatedInSAP Then
                StatusName = "Vendor Code Created in SAP"
                Creators = drv.Row.Item("creator").ToString.Split("\")
                SendTo = String.Format("{0}@groupeseb.com;{1}", Creators(1), "afok@groupeseb.com;ttom@groupeseb.com")
                cc = String.Format("{0};", ccdblistsb.ToString)
                SendToName = "All"
            End If

            Dim myEmail As NSEmail = New NSEmail
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
    Private Sub ToolStripButtonSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonSubmit.Click
        Try
            If IsDBNull(drv.Row.Item("status")) Then
                If Me.Validate Then
                    If MessageBox.Show("Do you want to submit this record?", "Submit", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                        drv.EndEdit()
                        drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusNew

                        ApprovalDRV = myAdapter.VendorInfoModiActionBS.AddNew
                        ApprovalDRV.Row.Item("status") = drv.Item("status")
                        ApprovalDRV.Row.Item("statusname") = "New"
                        ApprovalDRV.Row.Item("modifiedby") = HelperClass1.UserId
                        ApprovalDRV.Row.Item("vendorinfomodiid") = drv.Item("id")
                        ApprovalDRV.EndEdit()
                        Logger.log(String.Format("** Submit {0}**", HelperClass1.UserId))
                        If SaveRecord() Then
                            Debug.Print("send email")
                            SendEmail(drv)
                            'disable Commit button
                            ToolStripButtonCommit.Visible = False
                            RaiseEvent RefreshDataGridView(Me, New EventArgs)

                        End If

                    End If
                End If
            Else
                If drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusNew Then
                    MessageBox.Show("Record already submitted.")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        
    End Sub

    Private Sub ToolStripButtonValidate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonValidate.Click
        If IsNothing(ApprovalDRV) Then
            If MessageBox.Show("Do you want to validate this record?", "Validate", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                Dim remarks As String = InputBox("Please input some comment.")
                Dim myAddition As Integer
                Select Case drv.Row.Item("status")
                    Case 1, 4
                        myAddition = 2
                    Case 6
                        myAddition = -1                        
                    Case Else
                        myAddition = 1
                End Select
                'drv.Row.Item("status") = IIf(drv.Row.Item("status") = 1, 2, 1) + drv.Row.Item("status")
                drv.Row.Item("status") = myAddition + drv.Row.Item("status")
                ApprovalDRV = myAdapter.VendorInfoModiActionBS.AddNew
                ApprovalDRV.Row.Item("status") = drv.Item("status")
                ApprovalDRV.Row.Item("modifiedby") = HelperClass1.UserId
                ApprovalDRV.Row.Item("vendorinfomodiid") = drv.Item("id")
                ApprovalDRV.Row.Item("remark") = remarks
                ApprovalDRV.EndEdit()
                If SaveRecord() Then
                    SendEmail(drv)
                    'refresh Event
                    RaiseEvent RefreshDataGridView(Me, New EventArgs)
                End If


            End If
        Else
            MessageBox.Show("Nothing to do.")
        End If
    End Sub

    Private Sub ToolStripButtonComplete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonComplete.Click
        'This function will create related tables 
        If IsNothing(ApprovalDRV) Then
            If MessageBox.Show("Do you want to complete this record?", "Complete", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then

                Dim remarks As String = InputBox("Please input Vendor Code.")
                If remarks.Length > 0 Then
                    drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusCompleted
                    drv.Row.Item("vendorcode") = remarks
                    ApprovalDRV = myAdapter.VendorInfoModiActionBS.AddNew
                    ApprovalDRV.Row.Item("status") = drv.Item("status")
                    ApprovalDRV.Row.Item("modifiedby") = HelperClass1.UserId
                    ApprovalDRV.Row.Item("vendorinfomodiid") = drv.Item("id")
                    ApprovalDRV.Row.Item("remark") = remarks
                    ApprovalDRV.EndEdit()

                    If SaveRecord() Then
                        SendEmail(drv)
                    End If
                End If
                


            End If
        Else
            MessageBox.Show("Nothing to do.")
        End If
    End Sub

    Private Sub ToolStripButtonConfirmation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonConfirmation.Click
        If IsNothing(ApprovalDRV) Then
            If MessageBox.Show("Do you want to confirm this record?", "Confirmation", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                Dim remarks As String = InputBox("Please input some comment.")
                drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusCreatedInSAP
                ApprovalDRV = myAdapter.VendorInfoModiActionBS.AddNew
                ApprovalDRV.Row.Item("status") = drv.Item("status")
                ApprovalDRV.Row.Item("modifiedby") = HelperClass1.UserId
                ApprovalDRV.Row.Item("vendorinfomodiid") = drv.Item("id")
                ApprovalDRV.Row.Item("remark") = remarks
                ApprovalDRV.EndEdit()
                If SaveRecord() Then
                    SendEmail(drv)
                End If


            End If
        Else
            MessageBox.Show("Nothing to do.")
        End If
    End Sub

    Private Sub ToolStripButtonReject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonReject.Click
        If IsNothing(ApprovalDRV) Then
            If MessageBox.Show("Do you want to reject this record?", "Reject", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then

                Dim remarks As String = InputBox("Please input some comment.")

                Dim myAddition As Integer
                Select Case drv.Row.Item("status")
                    Case 1, 4
                        myAddition = 7
                    Case 6
                        myAddition = -1
                    Case Else
                        myAddition = 6
                End Select
                'drv.Row.Item("status") = IIf(drv.Row.Item("status") = 1, 7, 6) + drv.Row.Item("status")
                drv.Row.Item("status") = myAddition + drv.Row.Item("status")
                ApprovalDRV = myAdapter.VendorInfoModiActionBS.AddNew
                ApprovalDRV.Row.Item("status") = drv.Item("status")
                ApprovalDRV.Row.Item("modifiedby") = HelperClass1.UserId
                ApprovalDRV.Row.Item("vendorinfomodiid") = drv.Item("id")
                ApprovalDRV.Row.Item("remark") = remarks
                ApprovalDRV.EndEdit()

                If SaveRecord() Then
                    SendEmail(drv)
                End If

            End If
        Else
            MessageBox.Show("Nothing to do.")
        End If
    End Sub

    Private Sub ToolStripButtonReSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonReSubmit.Click

        If Not IsNothing(ApprovalDRV) Then
            MessageBox.Show("Nothing to do.")
        Else
            If MessageBox.Show("Do you want to re-submit this record?", "Re Submit", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                drv.Row.Item("status") = VendorInfoModiStatusEnum.StatusResubmit

                ApprovalDRV = myAdapter.VendorInfoModiActionBS.AddNew
                ApprovalDRV.Row.Item("status") = drv.Item("status")
                ApprovalDRV.Row.Item("modifiedby") = HelperClass1.UserId
                ApprovalDRV.Row.Item("vendorinfomodiid") = drv.Item("id")                
                ApprovalDRV.EndEdit()
                If SaveRecord() Then SendEmail(drv)

            End If
        End If
    End Sub

    Private Sub ToolStripButtonStatusCancelled_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonStatusCancelled.Click
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
                If SaveRecord() Then
                    SendEmail(drv)
                End If

            End If
        Else
            MessageBox.Show("Nothing to do.")
        End If
    End Sub
End Class