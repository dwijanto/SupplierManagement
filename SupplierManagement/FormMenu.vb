Imports System.Reflection
Imports SupplierManagement.PublicClass

'Remarks: Below are the AutoTask (3)
'
'1. FormAutoSendExpired
'2. FormSupplierManagementAutoReport
'3. FormAutoSendActonPlan


Public Enum TxRecord
    TxAdd = 1
    TxUpdate = 2
    TxDelete = 3
End Enum

Public Enum VendorInformationModificationReference
    YBudget = 0
    YForecast48 = 1
    YForecast84 = 2
    Ymin1Actual = 3
End Enum
Public Class FormMenu
    ' Delegate Function DBAdapterTable(ByVal obj As Object, ByVal arguments As Object)
    Private Sub FormMenu_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            'HelperClass1 = New HelperClass

            HelperClass1 = HelperClass.getInstance

            DbAdapter1 = New DbAdapter

            HelperClass1.UserInfo.IsAdmin = DbAdapter1.IsAdmin(HelperClass1.UserId)
            'HelperClass1.UserInfo.isOfficer = DbAdapter1.isOfficer(HelperClass1.UserId)
            HelperClass1.UserInfo.AllowUpdateDocument = DbAdapter1.AllowUpdateDocument(HelperClass1.UserId)
            HelperClass1.UserInfo.IsFinance = DbAdapter1.isFinance(HelperClass1.UserId)
            If HelperClass1.UserInfo.IsFinance Then
                Dim errMsg As String = String.Empty
                Dim myGroupList As New List(Of String)
                myGroupList = DbAdapter1.GetGroupName(HelperClass1.UserId, errMsg)
                If myGroupList.Count = 0 Then
                    MessageBox.Show(errMsg)
                Else
                    HelperClass1.UserInfo.GroupName = myGroupList(0)
                    HelperClass1.UserInfo.Remarks = myGroupList(1)
                End If
                'HelperClass1.UserInfo.GroupName = myhashtable(1)
            End If
           
            'HelperClass1.template = "\\172.22.10.44\SharedFolder\PriceCMMF\New\template"
            'HelperClass1.document = "\\172.22.10.44\SharedFolder\PriceCMMF\New\documents"
            'HelperClass1.attachment = "\\172.22.10.44\SharedFolder\PriceCMMF\New\attachment"
            HelperClass1.template = "\\172.22.10.77\SharedFolder\PriceCMMF\New\template"
            HelperClass1.document = "\\172.22.10.77\SharedFolder\PriceCMMF\New\documents"
            HelperClass1.attachment = "\\172.22.10.77\SharedFolder\PriceCMMF\New\attachment"
            HelperClass1.proformapo = "\\172.22.10.77\SharedFolder\PriceCMMF\New\proformapo"
            Try
                loglogin(DbAdapter1.userid)
            Catch ex As Exception

            End Try

            Me.Text = GetMenuDesc()
            Me.Location = New Point(300, 10)

            'hide overview/analysis
            'GroupBox1.Visible = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
        End Try


    End Sub
    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "Supplier Management"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        DbAdapter1.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub
    Public Function GetMenuDesc() As String
        'Label1.Text = "Welcome, " & HelperClass1.UserInfo.DisplayName
        Return "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & DbAdapter1.ConnectionStringDict.Item("HOST") & ", Database: " & DbAdapter1.ConnectionStringDict.Item("DATABASE") & ", Userid: " & HelperClass1.UserId

    End Function
    Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ctrl As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim assembly1 As Assembly = Assembly.GetAssembly(GetType(FormMenu))
        Dim frm As Form = CType(assembly1.CreateInstance(assembly1.GetName.Name.ToString & "." & ctrl.Tag.ToString, True), Form)
        Dim inMemory As Boolean = False
        For i = 0 To My.Application.OpenForms.Count - 1
            If My.Application.OpenForms.Item(i).Name = frm.Name Then
                ExecuteForm(My.Application.OpenForms.Item(i))
                inMemory = True
            End If
        Next
        If Not inMemory Then
            ExecuteForm(frm)
        End If
    End Sub

    Private Sub ExecuteForm(ByVal obj As Windows.Forms.Form)
        With obj
            .WindowState = FormWindowState.Normal
            .StartPosition = FormStartPosition.CenterScreen
            .Show()
            .Focus()
        End With
    End Sub

   

    Private Sub FormMenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not e.CloseReason = CloseReason.ApplicationExitCall Then
            If MessageBox.Show("Are you sure?", "Exit", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Me.CloseOpenForm()
                HelperClass1.fadeout(Me)
                DbAdapter1.Dispose()
                HelperClass1.Dispose()
            Else
                e.Cancel = True
            End If
        End If
    End Sub
    Private Sub CloseOpenForm()
        For i = 1 To (My.Application.OpenForms.Count - 1)
            My.Application.OpenForms.Item(1).Close()
        Next
    End Sub



    Private Sub NewAssetPurchaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewAssetPurchaseToolStripMenuItem.Click
        Dim myform As New FormAssetsPurchase(0)
        myform.Show()
    End Sub

    Private Sub FindAssetPurchaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindAssetPurchaseToolStripMenuItem.Click
        'Dim myform As New FormFindAssetsPurchase
        'myform.showdialog()
        Dim myreport As New FormReportAssetsPurchase("sqlstr", ReportType.AssetPurchase, "FileName")
        myreport.Show()



    End Sub


    Private Sub AssetsPurchaseReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AssetsPurchaseReportToolStripMenuItem.Click
        Dim myreport As New FormReportAssetsPurchase("sqlstr", ReportType.AssetPurchase, "FileName")
        myreport.Show()

    End Sub

    Private Sub UpdateToolingIdToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateToolingIdToolStripMenuItem.Click
        Dim myform As New FormUpdateToolingStatus
        myform.ShowDialog()
    End Sub

    Private Sub UserGuideAdminToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserGuideAdminToolStripMenuItem.Click
        Dim p As New System.Diagnostics.Process
        'p.StartInfo.FileName = "\\172.22.10.44\SharedFolder\PriceCMMF\New\template\Supplier Management Task User Guide-Admin.pdf"
        p.StartInfo.FileName = HelperClass1.template + "\Supplier Management Task User Guide-Admin.pdf"
        p.Start()
    End Sub


    Private Sub UserGuideAssetsPurchaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserGuideAssetsPurchaseToolStripMenuItem.Click
        Dim p As New System.Diagnostics.Process
        'p.StartInfo.FileName = "\\172.22.10.44\SharedFolder\PriceCMMF\New\template\Supplier Management Task User Guide-Assets Purchase.pdf"
        p.StartInfo.FileName = HelperClass1.template + "\Supplier Management Task User Guide-Assets Purchase.pdf"
        p.Start()
    End Sub

    Private Sub UserGuideUploadDocumentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserGuideUploadDocumentToolStripMenuItem.Click
        Dim p As New System.Diagnostics.Process
        'p.StartInfo.FileName = "\\172.22.10.44\SharedFolder\PriceCMMF\New\template\Supplier Management Task User Guide-Upload document.pdf"
        p.StartInfo.FileName = HelperClass1.template + "\Supplier Management Task User Guide-Upload document.pdf"
        p.Start()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim myform As New FormDocumentHeader
        Dim myform As New FormMyTaskDocument2
        myform.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        NewAssetPurchaseToolStripMenuItem_Click(Me, e)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim myform As New FormFindAssetsPurchase
        'myform.Show()
        Dim myreport As New FormReportAssetsPurchase("sqlstr", ReportType.AssetPurchase, "FileName")
        myreport.Show()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim myform As New FormSearchDocument
        myform.Show()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim myForm As New FormUser
        myForm.Show()
    End Sub



    Private Sub FormMyTaskDocumentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FormMyTaskDocumentToolStripMenuItem.Click

    End Sub

    Private Sub FormSupplierDashBoardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FormSupplierDashBoardToolStripMenuItem.Click
        Dim myform As New FormSupplierDashboard
        myform.Show()
    End Sub


    Private Sub Button8_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        'Dim myform As New FormFindAssetsPurchase
        'myform.Show()
        FindAssetPurchaseToolStripMenuItem.PerformClick()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim myform As New FormUpdateToolingStatus
        myform.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, SupplierDashboardToolStripMenuItem.Click
        Dim myform As New FormSupplierDashboard
        myform.Show()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim myform As New FormFullContract
        myform.Show()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim myform As New FormDocType
        myform.Show()
    End Sub


    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim myform As New FormMasterGroup
        myform.Show()
    End Sub


    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim myform As New FormMasterVendor
        myform.Show()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim myform As New FormVendorStatus001
        myform.Show()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim myform As New FormSupplierPanel
        myform.Show()
    End Sub



    Private Sub CountryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CountryToolStripMenuItem.Click, A1CountryToolStripMenuItem.Click
        Dim myfunction As New DBAdapterTable(AddressOf DbAdapter1.CountryTx)
        Dim myform As New FormMasterCountry(myfunction)

        'myform.TableName = "Country"
        'myform.Sqlstr = "select p.paramname as countryname, p.paramdtid as id from doc.paramdt p" &
        '              " left join doc.paramhd ph on ph.paramhdid = p.paramhdid" &
        '              " where ph.paramname = 'country'" &
        '              "order by p.paramname"

        'myform.ColumnId = "id"
        'myform.ColumnName = "countryname"
        'myform.labelColumnName = "Country Name"
        'myform.Text = "FormCountry"

        myform.Show()
    End Sub

    Private Sub ProvinceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProvinceToolStripMenuItem.Click, A2ProvinceToolStripMenuItem.Click
        Dim myfunction As New DBAdapterTable(AddressOf DbAdapter1.ProvinceTx)
        Dim myform As New FormMasterCountry(myfunction)
        'Dim abc As New FormSupplierDashboard
        myform.TableName = "Province"
        myform.Sqlstr = "select p.paramname as provincename, p.paramdtid as id from doc.paramdt p" &
                      " left join doc.paramhd ph on ph.paramhdid = p.paramhdid" &
                      " where ph.paramname = 'province'" &
                      "order by p.paramname"

        myform.ColumnId = "id"
        myform.ColumnName = "provincename"
        myform.LabelColumnName = "Province Name"
        myform.Text = "FormProvince"
        myform.Show()
    End Sub

    Private Sub DocumentDownloadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DocumentDownloadToolStripMenuItem.Click

    End Sub

    Private Sub OtherDocumentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OtherDocumentsToolStripMenuItem.Click
        Dim p As New System.Diagnostics.Process
        p.StartInfo.FileName = "explorer.exe"
        'Dim MyFolder = "\\172.22.10.44\sharedfolder\pricecmmf\new\otherdocuments"
        Dim MyFolder = "\\172.22.10.77\sharedfolder\pricecmmf\new\otherdocuments"
        p.StartInfo.Arguments = String.Format("{0}", MyFolder)       
        p.Start()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub




    Private Sub FormMenu_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.FormMenu_Load(Me, New EventArgs)

        Try
            AddHandler UploadDocumentToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierDocumentRawDataToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierCategoryToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler PanelStatusToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler VendorStatusToolStripMenuItem1.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierPanelToolStripMenuItem1.Click, AddressOf ToolStripMenuItem_Click
            AddHandler MasterStatusToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler DocumentCountToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler ContractToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler DocumentToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler UserToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler GroupToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SearchDocumentToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler MasterSupplierToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler MasterVendorToolStripMenuItem1.Click, AddressOf ToolStripMenuItem_Click
            AddHandler NewAssetPurchaseToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler FamilyGroupSBUToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler MasterDocTypeToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler DocumentDownloadToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler FormMyTaskDocumentToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler ImportTurnoverToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler ImportNQSUToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler ImportLogisticsToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler ImportProjectPDToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler ImportVendorPaymentTermToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler ImportSEBPlatformToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler ImportCitiProgramToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler FactoryContactToolStripMenuItem1.Click, AddressOf ToolStripMenuItem_Click
            AddHandler TechnologyToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler FactoryContactToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierDocumentSIFIDToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierDocumentSocialAuditToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierDocumentContractToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierPhotosToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler ImportBudgetForecastToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler DataTypeToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierGSMToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler FamilyPMToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierFamilyToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierPMToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierFamilyAssignmenttToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler MasterDataToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A1UserToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A2GroupToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A1MasterSupplierToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A2SupplierGSMToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A3SupplierFamilyToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A4SupplierFamilyToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A5ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A6SupplierPanelToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A1MasterSupplierToolStripMenuItem1.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A2SupplierPMToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A3SupplierStatusToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A4SupplierPanelToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A3FactoryContactHistoryToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler A1DocTypeToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler A1FamilyGroupSBUToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A2PanelStatusToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler A3SupplierCategoryToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A4MasterStatusToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler A5TechnologyToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler A6DataTypeToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler A1ImportTurnoverToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A2ImportNQSUToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler A3ImportLogisticsToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A4ImportProjectStatusToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A5ImportVendorPaymentTermToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A6ImportSEBAsiaPlatformToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A7ImportCitiProgramToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A8ImportBudgetForecastToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            'AddHandler A9ImportToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler ImportZF0041forSAPToolingContractToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler A1SupplierCurToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierCurrencyToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SSSActionPlanFollowUpToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler ImportPlantZZA013ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler FamilySupplierCreationToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SubFamilySupplierCreationToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler AssetPurchaseApprovalToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler SupplierFamilyExceptionToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierPaymentTermEffectiveDateToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler A7PaymentTermToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler IndirectFamilyToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click


            AddHandler A8VendorModificationTypeMasterToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler A9ModificationDocumentTypeMasterToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler CMMF3750ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            AddHandler SupplierListWithTONQSUSSLToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            AddHandler AssetPurchaseApprovalToolStripMenuItem1.Click, AddressOf ToolStripMenuItem_Click

            AddHandler ToolingSupplierToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

            'AddHandler FindAssetPurchaseToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
            'Admin
            MasterToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin 'And Not HelperClass1.UserInfo.isOfficer
            SupplierDocumentRawDataToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin
            SupplierDocumentSIFIDToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin
            SupplierDocumentSocialAuditToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin
            SupplierDocumentContractToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin
            DocumentCountToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin
            GroupBox4.Visible = HelperClass1.UserInfo.IsAdmin
            DocumentCountToolStripMenuItem.Visible = False
            AdminActionToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin



            DataTypeToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin
            'AssetManagementToolStripMenuItem.Visible = HelperClass1.UserInfo.IsAdmin
            VendorInformationApprovalToolStripMenuItem1.Visible = False
            If HelperClass1.UserInfo.IsFinance Then
                SupplierDocumentsToolStripMenuItem.Visible = False
                MasterStatusToolStripMenuItem.Visible = False
                ReportToolStripMenuItem.Visible = False
                UserGuideToolStripMenuItem.Visible = False
                NewAssetPurchaseToolStripMenuItem.Visible = False
                UpdateToolingIdToolStripMenuItem.Visible = False
                ImportZF0041forSAPToolingContractToolStripMenuItem.Visible = False
                ImportPlantZZA013ToolStripMenuItem.Visible = False
                VendorInformationApprovalToolStripMenuItem1.Visible = True
                'GroupBox1.Visible = False
                GroupBox2.Visible = False
                GroupBox3.Visible = False
                GroupBox4.Visible = False
                Label1.Visible = False
                Dim mysize = New Point(900, 200)
                Me.MinimumSize = mysize
                Me.Size = mysize
            End If

        Catch ex As Exception
            MessageBox.Show(String.Format("Load: {0}", ex.Message))
        End Try


    End Sub

   
  

  

    Private Sub CreateVendorInformationModificationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateVendorInformationModificationToolStripMenuItem.Click, Button15.Click
        Dim myvendor As New DialogVendorSelection
        If myvendor.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim myFormVendor = New FormVendorInformationModification(myvendor.VendorDR, myvendor.YearRef, myvendor.ComboBox1.SelectedIndex, TxEnum.NewRecord, 0)
            'Dim myFormVendor = New FormVendorInformationModification(myvendor.VendorDR, myvendor.YearRef, myvendor.ComboBox1.SelectedIndex, TxEnum.UpdateRecord, 19)
            myFormVendor.ShowDialog()
        End If
    End Sub

    Private Sub FindVendorInformationModificationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindVendorInformationModificationToolStripMenuItem.Click, Button17.Click
        Dim myVendor As New FormFindVendorInformationModification
        myVendor.ShowDialog()
    End Sub

  
    Private Sub VendorInformationApprovalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VendorInformationApprovalToolStripMenuItem.Click, VendorInformationApprovalToolStripMenuItem1.Click, SupplierInformationApprovalToolStripMenuItem.Click
        Dim myform As New FormVendorInformationTask()

        myform.Show()
    End Sub

    Private Sub CreateNewVendorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateNewVendorToolStripMenuItem.Click, Button14.Click
        Dim myform As New FormNewVendor(SupplierManagement.TxEnum.NewRecord, 0)
        'Dim myform As New FormNewVendor(SupplierManagement.TxEnum.UpdateRecord, 251)
        myform.Show()
    End Sub

    Private Sub FindNewVendorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindNewVendorToolStripMenuItem.Click, Button16.Click
        Dim myform As New FormFindNewVendor
        myform.Show()
    End Sub

    Private Sub CurrencyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CurrencyToolStripMenuItem.Click

        Dim myform As New FormParamDTL("currency")
        With myform.DataGridView1.Columns(0)
            .HeaderText = "Currency"
            .Width = 150
        End With
        With myform.DataGridView1.Columns(1)
            .HeaderText = "Sort Order"
            .Visible = True
        End With
        myform.DataGridView1.Columns(1).Visible = True
        myform.Show()
    End Sub

    Private Sub CMMF3750ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMMF3750ToolStripMenuItem.Click

    End Sub

    Private Sub MasterSupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MasterSupplierToolStripMenuItem.Click

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click

    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click

    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click

    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click

    End Sub

    Private Sub AssetPurchaseApprovalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AssetPurchaseApprovalToolStripMenuItem.Click

    End Sub

    Private Sub SupplierInformationApprovalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupplierInformationApprovalToolStripMenuItem.Click

    End Sub

    Private Sub AssetPurchaseApprovalToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AssetPurchaseApprovalToolStripMenuItem1.Click

    End Sub

    Private Sub SupplierListWithTONQSUSSLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupplierListWithTONQSUSSLToolStripMenuItem.Click

    End Sub

    Private Sub ToolingSupplierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolingSupplierToolStripMenuItem.Click

    End Sub
End Class
