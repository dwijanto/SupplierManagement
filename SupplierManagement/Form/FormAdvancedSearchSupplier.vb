Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass

Public Class FormAdvancedSearchSupplier
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myThreadSearch As New System.Threading.Thread(AddressOf DoSearch)

    Private Const SocialAudit As Integer = 22
    Private Const SIF As Integer = 21
    Private Const IDENTITY_SHEET As Integer = 54



    Dim DS As DataSet
    Dim DSSearch As DataSet
    Dim VendorBS As BindingSource
    Dim ShortnameBS As BindingSource
    Dim PMBS As BindingSource
    Dim SPMBS As BindingSource
    Dim ProductTypeBS As BindingSource
    Dim StatusNameBS As BindingSource
    Dim SBUBS As BindingSource
    Dim DocTypeBS As BindingSource
    Dim ProjectNameBS As BindingSource

    Dim docTypeId As Integer
    Dim sb As New StringBuilder
    Public WithEvents DocumentBS As BindingSource
    Dim SelectedFolder As String = String.Empty

    Dim TextHelper As UCTextWithHelper
    Dim SqlStrHelper As String

    Dim WithSocialAuditSB As StringBuilder
    Dim WithSB As StringBuilder
    Dim withTechnologySB As StringBuilder
    Dim WhereSB As StringBuilder

    Dim LeftJoinMMSB As StringBuilder
    Dim LeftJoinSTR As String = String.Empty

    Dim LeftJoinProjectSB As StringBuilder
    Dim LeftJoinProjectSTR As String = String.Empty

    Dim LeftJoinSocialAuditSB As StringBuilder
    Dim LeftJoinSocialAuditSTR As String = String.Empty

    Dim leftJoinVFPSB As StringBuilder
    Dim leftJoinVFPSTR As String = String.Empty

    Dim LeftJoinSupplierPanelSB As StringBuilder
    Dim LeftJoinSupplierPanelSTR As String

    Dim LeftJoinTechnologySB As StringBuilder
    Dim LeftJoinTechnologySTR As String

    Dim LeftJoinFactorySB As StringBuilder
    Dim LeftJoinFactorySTR As String

    Dim LeftJoinSIFSB As StringBuilder
    Dim LeftJoinSIFSTR As String

    Dim LeftJoinIDSB As StringBuilder
    Dim LeftJoinIDSTR As String

    Dim WithSIFSB As StringBuilder
    Dim WithIDSB As StringBuilder

    Dim AdditionalDisplayFieldSB As StringBuilder
    Dim AdditionalDisplayFieldSTR As String
    

    Dim WithSIFIDSTRLatest As String
    Dim WithSIFIDSTRNOTLatest As String



    Private Sub loaddata()

        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub



    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Dim mymessage As String = String.Empty
        Dim myquery As String = GetMainQuery(HelperClass1.UserInfo.IsAdmin)
        DS = New DataSet
        If DbAdapter1.TbgetDataSet(myquery, DS, mymessage) Then
            Try

            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(4, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        If Not IsNothing(DocumentBS) Then
            ProgressReport(1, String.Format("{0} {1} record(s) found.", "Loading Data.Done!", DocumentBS.Count))
        Else
            ProgressReport(1, String.Format("{0}", "Loading Data.Done!"))
        End If

        ProgressReport(5, "Continuous")
    End Sub

    Private Function getWITHSIFID(ByVal isLatest As Boolean, ByVal doctypeid As Integer, ByVal withquery As String)
        Dim myret As String
        If isLatest Then
            myret = String.Format("{1} as (select distinct vendorcode, id from (select vd.vendorcode,first_value(d.id) over (partition by vd.vendorcode order by docdate desc) as id from doc.vendordoc vd" &
                        " left join doc.document d on d.id = vd.documentid" &
                        " where(d.doctypeid = {0} or d.doctypeid = 54)" &
                        " group by vd.vendorcode,d.id) as foo" &
                        " order by vendorcode)", doctypeid, withquery)
        Else
            myret = String.Format("{1} as (select vd.vendorcode,d.id,d.docdate from doc.vendordoc vd" &
                        " left join doc.document d on d.id = vd.documentid" &
                        " where(d.doctypeid = {0} or d.doctypeid = 54)" &
                        " order by vd.vendorcode)", doctypeid, withquery)
        End If
        Return myret
    End Function

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        AddHandler UcTextWithHelperProductType.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperSebBrand.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperRangeCode.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperCMMF.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperCommercialCode.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperProject.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperSocialAudit.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperVFP.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperSupplierStatus.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperSBU.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperFamily.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperSupplierCategory.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperFPPanel.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperCPPanel.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperGSM.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperSPM.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperPM.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperProvince.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperCountry.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperMainProductSold.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperMainCustomer.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperSupplierBrand.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperIDSBU.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperIDFamily.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperIDOEMODM.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperIDMaterial.ButtonClick, AddressOf Me.GetData
        AddHandler UcTextWithHelperTechnology.ButtonClick, AddressOf Me.GetData

        WhereSB = New StringBuilder
        LeftJoinMMSB = New StringBuilder
        LeftJoinProjectSB = New StringBuilder
        WithSB = New StringBuilder
        WithSocialAuditSB = New StringBuilder
        LeftJoinSocialAuditSB = New StringBuilder
        leftJoinVFPSB = New StringBuilder
        LeftJoinSupplierPanelSB = New StringBuilder
        LeftJoinTechnologySB = New StringBuilder
        LeftJoinFactorySB = New StringBuilder
        LeftJoinSIFSB = New StringBuilder
        LeftJoinIDSB = New StringBuilder
        WithSIFSB = New StringBuilder
        WithIDSB = New StringBuilder
        withTechnologySB = New StringBuilder
        AdditionalDisplayFieldSB = New StringBuilder

        With UcTextWithHelperProductType
            .WhereSB = WhereSB
            .SearchField = "lower(tu.fpcp)"
            .TextBox1.Text = My.Settings.FASS_ProductType
        End With
        With UcTextWithHelperSebBrand
            .WhereSB = WhereSB
            .SearchField = "lower(tu.brand)"
            .TextBox1.Text = My.Settings.FASS_SEBBrand
        End With
        With UcTextWithHelperRangeCode
            .WhereSB = WhereSB
            .SearchField = "lower(tu.range)"
            .TextBox1.Text = My.Settings.FASS_RangeCode
        End With
        With UcTextWithHelperCMMF
            .WhereSB = WhereSB
            .SearchField = "lower(tu.cmmf::text)"
            .TextBox1.Text = My.Settings.FASS_CMMF
        End With
        With UcTextWithHelperCommercialCode
            .WhereSB = WhereSB           
            .SearchField = "lower(mm.commref)" '"(replace(replace(mm.commref,'(','\('),')','\)') || ',')"
            .LeftJoinMMSB = LeftJoinMMSB
            .LeftJoinStr = "left join materialmaster mm on mm.cmmf = tu.cmmf "
            .TextBox1.Text = My.Settings.FASS_CommercialCode
        End With
        With UcTextWithHelperProject 'Tooling list
            .WhereSB = WhereSB
            .SearchField = "lower(tp.projectname)"
            .LeftJoinProjectSB = LeftJoinProjectSB
            .LeftJoinProjectSTR = " left join doc.assetpurchase ap on ap.vendorcode = v.vendorcode" &
                                    " left join doc.toolingproject tp on tp.id = ap.projectid "
            .TextBox1.Text = My.Settings.FASS_Project
        End With
        With UcTextWithHelperSocialAudit 'Document
            .WhereSB = WhereSB
            .SearchField = "lower(sau.auditgrade)"
            .WithSB = WithSB
            .WithSocialAuditSB = WithSocialAuditSB
            '.WithSocialAuditSTR = String.Format("sa as(" &
            '                      " select vd.vendorcode,first_value(d.id) over (partition by vd.vendorcode order by docdate desc) as id  " &
            '                      " from doc.vendordoc vd " &
            '                      " left join doc.document d on d.id = vd.documentid where(d.doctypeid = {0})" &
            '                      "group by vd.vendorcode,d.id)", SocialAudit)
            .WithSocialAuditSTR = String.Format("sa as(" &
                                  " select vendorcode,id from (select v.shortname,first_value(d.id) over (partition by v.shortname order by docdate desc) as id  " &
                                  " from doc.vendordoc vd " &
                                  " left join vendor v on v.vendorcode = vd.vendorcode" &
                                  " left join doc.document d on d.id = vd.documentid where(d.doctypeid = {0} )" &
                                  "group by v.shortname,d.id) as foo left join vendor v on v.shortname = foo.shortname)", SocialAudit)
            .LeftJoinSocialAuditSB = LeftJoinSocialAuditSB
            .LeftJoinSocialAuditSTR = "left join sa on sa.vendorcode = v.vendorcode " &
                                      "left join doc.socialaudit sau on sau.documentid = sa.id"
            .TextBox1.Text = My.Settings.FASS_LatestSocialAudit
        End With
        With UcTextWithHelperVFP 'Finance Program
            .WhereSB = WhereSB
            .SearchField = "lower(cp.fpcp)"
            .LeftJoinVFPSB = leftJoinVFPSB
            .LeftJoinVFPSTR = " left join doc.citiprogram cp on cp.vendorcode = v.vendorcode"
            .TextBox1.Text = My.Settings.FASS_VendorFinanceProgram
        
        End With
        With UcTextWithHelperSupplierStatus
            .WhereSB = WhereSB
            .SearchField = "lower(pr.paramname)"
            .TextBox1.Text = My.Settings.FASS_SupplierStatus
        End With

        With UcTextWithHelperSBU
            .WhereSB = WhereSB
            .SearchField = "lower(tu.sbu)"
            .TextBox1.Text = My.Settings.FASS_SBU
        End With

        With UcTextWithHelperTechnology
            .WhereSB = WhereSB

            .WithSB = WithSB
            .WithTechnologySB = withTechnologySB

            .WithTechnologySTR = String.Format("st as(" &
                                  " select * from crosstab('select vtec.vendorcode,row_number() over(partition by vtec.vendorcode order by vtec.vendorcode,vtec.lineno),tec.technologyname from doc.vendortechnology vtec" &
                                  " left join doc.technology tec on tec.id = vtec.technologyid" &
                                  " order by vendorcode,lineno','select m from generate_series(1,5) m') as (vendorcode bigint,smt character varying, st1 character varying, st2 character varying, st3 character varying, st4 character varying))")

            .LeftJoinTechnologySB = LeftJoinTechnologySB
            .LeftJoinTechnologySTR = " left join doc.vendortechnology  vtec on vtec.vendorcode = v.vendorcode" &
                                     " left join doc.technology tec on tec.id = vtec.technologyid" &
                                     " left join st on st.vendorcode = v.vendorcode"

            .SearchField = "lower(tec.technologyname)"


            .AdditionalDisplayFieldSB = AdditionalDisplayFieldSB
            .AdditionalDisplayFieldSTR = ",st.smt,st.st1,st.st2,st.st3,st.st4 "
            .TextBox1.Text = My.Settings.FASS_Technology
        End With

        With UcTextWithHelperFamily
            .WhereSB = WhereSB
            .SearchField = "lower(f.familyname)"
            .LeftJoinVFPSB = leftJoinVFPSB
            .LeftJoinVFPSTR = " left join family f on f.familyid = tu.comfam"
            .TextBox1.Text = My.Settings.FASS_FamilyCode
        End With
        With UcTextWithHelperSupplierCategory
            .WhereSB = WhereSB
            .SearchField = "lower(sc.category)"
            .TextBox1.Text = My.Settings.FASS_SupplierCategory
        End With

        LeftJoinSupplierPanelSTR = " left join doc.panelstatus psfp on psfp.id = sp.fp" &
                                        " left join doc.panelstatus cpfp on cpfp.id = sp.cp"
        With UcTextWithHelperFPPanel
            .WhereSB = WhereSB
            .SearchField = "lower(psfp.paneldescription)"
            .LeftJoinSupplierPanelSB = LeftJoinSupplierPanelSB
            .LeftJoinSupplierPanelSTR = LeftJoinSupplierPanelSTR
            .TextBox1.Text = My.Settings.FASS_FPPanel
        End With
        With UcTextWithHelperCPPanel
            .WhereSB = WhereSB
            .SearchField = "lower(cpfp.paneldescription)"
            .LeftJoinSupplierPanelSB = LeftJoinSupplierPanelSB
            .LeftJoinSupplierPanelSTR = LeftJoinSupplierPanelSTR
            .TextBox1.Text = My.Settings.FASS_CPPanel
        End With
        With UcTextWithHelperGSM
            .WhereSB = WhereSB
            .SearchField = "lower(gsm.username)"
            .TextBox1.Text = My.Settings.FASS_SPM
        End With
        With UcTextWithHelperSPM
            .WhereSB = WhereSB
            '.SearchField = "lower(spm.officersebname)"
            .SearchField = "lower(spm.username)"
            .TextBox1.Text = My.Settings.FASS_SPM
        End With
        With UcTextWithHelperPM
            .WhereSB = WhereSB
            '.SearchField = "lower(pm.officersebname)"
            .SearchField = "lower(pm.username)"
            .TextBox1.Text = My.Settings.FASS_PM
        End With
        LeftJoinFactorySTR = " left join doc.vendorfactory vf on vf.vendorcode = v.vendorcode" &
                                  " left join doc.factorydtl fd on fd.id = vf.factoryid" &
                                  " left join doc.paramdt prov on prov.paramdtid = fd.provinceid" &
                                  " left join doc.paramdt c on c.paramdtid = fd.countryid"

        With UcTextWithHelperProvince
            .WhereSB = WhereSB
            .SearchField = "lower(prov.paramname)"
            .LeftJoinFactorySB = LeftJoinFactorySB
            .LeftJoinFactorySTR = LeftJoinFactorySTR
            .TextBox1.Text = My.Settings.FASS_Province
        End With
        With UcTextWithHelperCountry
            .WhereSB = WhereSB
            .SearchField = "lower(c.paramname)"
            .LeftJoinFactorySB = LeftJoinFactorySB
            .LeftJoinFactorySTR = LeftJoinFactorySTR
            .TextBox1.Text = My.Settings.FASS_Country
        End With

        
        LeftJoinSIFSTR = "left join vdsif on vdsif.vendorcode = v.vendorcode " &
                         "left join doc.sitx ssif on ssif.documentid = vdsif.id" &
                         " left join doc.silabel slsif on slsif.id = ssif.labelid"
        LeftJoinIDSTR = "left join vdid on vdid.vendorcode = v.vendorcode " &
                        "left join doc.sitx sid on sid.documentid = vdid.id" &
                        " left join doc.silabel slid on slid.id = sid.labelid"

        With UcTextWithHelperMainProductSold
            .WhereSB = WhereSB
            .SearchField = "slsif.name = 'Main Product Sold' and  lower(ssif.value) " '"Main Product Sold"
            .WithSB = WithSB
            .WithSIFSB = WithSIFSB
            .WithSIFIDSSTRLatest = getWITHSIFID(True, SIF, "vdsif")
            .WithSIFIDSSTRNOTLatest = getWITHSIFID(False, SIF, "vdsif")
            .LeftJoinSIFSB = LeftJoinSIFSB
            .LeftJoinSIFSTR = LeftJoinSIFSTR
            .TextBox1.Text = My.Settings.FASS_MainProductSold
        End With
        With UcTextWithHelperMainCustomer
            .WhereSB = WhereSB
            .SearchField = "slsif.name = 'Main Customer' and  lower(ssif.value) " ' "Main Customer"
            .WithSB = WithSB
            .WithSIFSB = WithSIFSB
            .WithSIFIDSSTRLatest = getWITHSIFID(True, SIF, "vdsif")
            .WithSIFIDSSTRNOTLatest = getWITHSIFID(False, SIF, "vdsif")
            .LeftJoinSIFSB = LeftJoinSIFSB
            .LeftJoinSIFSTR = LeftJoinSIFSTR
            .TextBox1.Text = My.Settings.FASS_MainCustomers
        End With
        With UcTextWithHelperSupplierBrand
            .WhereSB = WhereSB
            .SearchField = "slsif.name = 'Supplier Product Brand' and  lower(ssif.value) " '"Supplier Product"
            .WithSB = WithSB
            .WithSIFSB = WithSIFSB
            .WithSIFIDSSTRLatest = getWITHSIFID(True, SIF, "vdsif")
            .WithSIFIDSSTRNOTLatest = getWITHSIFID(False, SIF, "vdsif")
            .LeftJoinSIFSB = LeftJoinSIFSB
            .LeftJoinSIFSTR = LeftJoinSIFSTR
            .TextBox1.Text = My.Settings.FASS_SuppliersProductBrand
        End With
        With UcTextWithHelperIDSBU
            .WhereSB = WhereSB
            .SearchField = "slid.name = 'SBU' and  lower(sid.value) " '"SBU"
            .WithSB = WithSB
            .WithIDSB = WithIDSB
            .WithSIFIDSSTRLatest = getWITHSIFID(True, IDENTITY_SHEET, "vdid")
            .WithSIFIDSSTRNOTLatest = getWITHSIFID(False, IDENTITY_SHEET, "vdid")
            .LeftJoinIDSB = LeftJoinIDSB
            .LeftJoinIDSTR = LeftJoinIDSTR
            .TextBox1.Text = My.Settings.FASS_SBUIdentitySheet
        End With
        With UcTextWithHelperIDFamily
            .WhereSB = WhereSB
            .SearchField = "slid.name = 'Family Name' and  lower(sid.value) " '"Family Name"
            .WithSB = WithSB
            .WithIDSB = WithIDSB
            .WithSIFIDSSTRLatest = getWITHSIFID(True, IDENTITY_SHEET, "vdid")
            .WithSIFIDSSTRNOTLatest = getWITHSIFID(False, IDENTITY_SHEET, "vdid")
            .LeftJoinIDSB = LeftJoinIDSB
            .LeftJoinIDSTR = LeftJoinIDSTR
            .TextBox1.Text = My.Settings.FASS_Family
        End With
        With UcTextWithHelperIDOEMODM
            .WhereSB = WhereSB
            .SearchField = "slid.name = 'OEM ODM' and  lower(sid.value) " '"OEM ODM"
            .WithSB = WithSB
            .WithIDSB = WithIDSB
            .WithSIFIDSSTRLatest = getWITHSIFID(True, IDENTITY_SHEET, "vdid")
            .WithSIFIDSSTRNOTLatest = getWITHSIFID(False, IDENTITY_SHEET, "vdid")
            .LeftJoinIDSB = LeftJoinIDSB
            .LeftJoinIDSTR = LeftJoinIDSTR
            .TextBox1.Text = My.Settings.FASS_OEMODM
        End With
        With UcTextWithHelperIDMaterial
            .WhereSB = WhereSB
            .SearchField = "slid.name = 'Material' and  lower(sid.value) " '"Material"
            .WithSB = WithSB
            .WithIDSB = WithIDSB
            .WithSIFIDSSTRLatest = getWITHSIFID(True, IDENTITY_SHEET, "vdid")
            .WithSIFIDSSTRNOTLatest = getWITHSIFID(False, IDENTITY_SHEET, "vdid")
            .LeftJoinIDSB = LeftJoinIDSB
            .LeftJoinIDSTR = LeftJoinIDSTR
            .TextBox1.Text = My.Settings.FASS_Material
        End With
        ' Add any initialization after the InitializeComponent() call.
        CheckBox1.Checked = My.Settings.FASS_LatestOnly
        If RememberFilter Then
            CheckBox2.Checked = True
        End If
    End Sub

    Private Function RememberFilter() As Boolean
        Dim myret As Boolean = False
        If My.Settings.FASS_ProductType <> "" Then
            myret = True
        ElseIf My.Settings.FASS_SEBBrand <> "" Then
            myret = True
        ElseIf My.Settings.FASS_RangeCode <> "" Then
            myret = True
        ElseIf My.Settings.FASS_CMMF <> "" Then
            myret = True
        ElseIf My.Settings.FASS_CommercialCode <> "" Then
            myret = True
        ElseIf My.Settings.FASS_Project <> "" Then
            myret = True
        ElseIf My.Settings.FASS_LatestSocialAudit <> "" Then
            myret = True
        ElseIf My.Settings.FASS_VendorFinanceProgram <> "" Then
            myret = True
        ElseIf My.Settings.FASS_SupplierStatus <> "" Then
            myret = True
        ElseIf My.Settings.FASS_SBU <> "" Then
            myret = True
        ElseIf My.Settings.FASS_FamilyCode <> "" Then
            myret = True
        ElseIf My.Settings.FASS_SupplierCategory <> "" Then
            myret = True
        ElseIf My.Settings.FASS_FPPanel <> "" Then
            myret = True
        ElseIf My.Settings.FASS_CPPanel <> "" Then
            myret = True
        ElseIf My.Settings.FASS_SPM <> "" Then
            myret = True
        ElseIf My.Settings.FASS_PM <> "" Then
            myret = True
        ElseIf My.Settings.FASS_Province <> "" Then
            myret = True
        ElseIf My.Settings.FASS_Country <> "" Then
            myret = True
        ElseIf My.Settings.FASS_MainCustomers <> "" Then
            myret = True
        ElseIf My.Settings.FASS_SuppliersProductBrand <> "" Then
            myret = True
        ElseIf My.Settings.FASS_MainProductSold <> "" Then
            myret = True
        ElseIf My.Settings.FASS_SBUIdentitySheet <> "" Then
            myret = True
        ElseIf My.Settings.FASS_Family <> "" Then
            myret = True
        ElseIf My.Settings.FASS_OEMODM <> "" Then
            myret = True
        ElseIf My.Settings.FASS_Material <> "" Then
            myret = True
        ElseIf My.Settings.FASS_LatestOnly <> True Then
            myret = True
        End If
        Return myret
    End Function
    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            My.Settings.FASS_ProductType = UcTextWithHelperProductType.TextBox1.Text
            My.Settings.FASS_SEBBrand = UcTextWithHelperSebBrand.TextBox1.Text
            My.Settings.FASS_RangeCode = UcTextWithHelperRangeCode.TextBox1.Text
            My.Settings.FASS_CMMF = UcTextWithHelperCMMF.TextBox1.Text
            My.Settings.FASS_CommercialCode = UcTextWithHelperCommercialCode.TextBox1.Text
            My.Settings.FASS_Project = UcTextWithHelperProject.TextBox1.Text
            My.Settings.FASS_LatestSocialAudit = UcTextWithHelperSocialAudit.TextBox1.Text
            My.Settings.FASS_VendorFinanceProgram = UcTextWithHelperVFP.TextBox1.Text
            My.Settings.FASS_SupplierStatus = UcTextWithHelperSupplierStatus.TextBox1.Text
            My.Settings.FASS_SBU = UcTextWithHelperSBU.TextBox1.Text
            My.Settings.FASS_FamilyCode = UcTextWithHelperFamily.TextBox1.Text
            My.Settings.FASS_SupplierCategory = UcTextWithHelperSupplierCategory.TextBox1.Text
            My.Settings.FASS_FPPanel = UcTextWithHelperFPPanel.TextBox1.Text
            My.Settings.FASS_CPPanel = UcTextWithHelperCPPanel.TextBox1.Text
            My.Settings.FASS_SPM = UcTextWithHelperSPM.TextBox1.Text
            My.Settings.FASS_PM = UcTextWithHelperPM.TextBox1.Text
            My.Settings.FASS_Province = UcTextWithHelperProvince.TextBox1.Text
            My.Settings.FASS_Country = UcTextWithHelperCountry.TextBox1.Text
            My.Settings.FASS_MainCustomers = UcTextWithHelperMainCustomer.TextBox1.Text
            My.Settings.FASS_SuppliersProductBrand = UcTextWithHelperSupplierBrand.TextBox1.Text
            My.Settings.FASS_MainProductSold = UcTextWithHelperMainProductSold.TextBox1.Text
            My.Settings.FASS_SBUIdentitySheet = UcTextWithHelperIDSBU.TextBox1.Text
            My.Settings.FASS_Family = UcTextWithHelperIDFamily.TextBox1.Text
            My.Settings.FASS_OEMODM = UcTextWithHelperIDOEMODM.TextBox1.Text
            My.Settings.FASS_Material = UcTextWithHelperIDMaterial.TextBox1.Text
            My.Settings.FASS_LatestOnly = CheckBox1.Checked
        Else
            My.Settings.FASS_ProductType = ""            
            My.Settings.FASS_SEBBrand = ""
            My.Settings.FASS_RangeCode = ""
            My.Settings.FASS_CMMF = ""
            My.Settings.FASS_CommercialCode = ""
            My.Settings.FASS_Project = ""
            My.Settings.FASS_LatestSocialAudit = ""
            My.Settings.FASS_VendorFinanceProgram = ""
            My.Settings.FASS_SupplierStatus = ""
            My.Settings.FASS_SBU = ""
            My.Settings.FASS_FamilyCode = ""
            My.Settings.FASS_SupplierCategory = ""
            My.Settings.FASS_FPPanel = ""
            My.Settings.FASS_CPPanel = ""
            My.Settings.FASS_SPM = ""
            My.Settings.FASS_PM = ""
            My.Settings.FASS_Province = ""
            My.Settings.FASS_Country = ""
            My.Settings.FASS_MainCustomers = ""
            My.Settings.FASS_SuppliersProductBrand = ""
            My.Settings.FASS_MainProductSold = ""
            My.Settings.FASS_SBUIdentitySheet = ""
            My.Settings.FASS_Family = ""
            My.Settings.FASS_OEMODM = ""
            My.Settings.FASS_Material = ""
            My.Settings.FASS_LatestOnly = True
        End If
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
                        Try
                            DocumentBS = New BindingSource
                            DocumentBS.DataSource = DS.Tables(0)
                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = DocumentBS
                            DataGridView1.Invalidate()
                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                    Case 7
                        Try
                            DocumentBS = New BindingSource
                            DocumentBS.DataSource = DSSearch.Tables(0)
                            TextHelper.HelperBindingSource = DocumentBS
                            TextHelper.LoadHelper()
                        Catch ex As Exception
                            message = ex.Message
                        End Try
                   
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub


    Private Sub DoSearch()
        DSSearch = New DataSet
        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sb.ToString, DSSearch, mymessage) Then
            Try
                DSSearch.Tables(0).TableName = "Document"
            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(7, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If

        ProgressReport(1, String.Format("{0} {1:#,##0} record(s) found.", "Loading Data.Done!", DocumentBS.Count))

        ProgressReport(5, "Continuous")
    End Sub

    Private Function getUserVendor()
        Dim myquery As String

        myquery = String.Format("with uv as (select distinct v.vendorcode  from doc.user u  left join doc.groupuser gu on gu.userid = u.id" &
                                " left join doc.groupvendor gv on gv.groupid = gu.groupid left join vendor v on v.vendorcode = gv.vendorcode " &
                                " where lower(u.userid)  ~ '{0}$'  ) ", HelperClass1.UserInfo.userid)
        '" where lower(u.userid) = '{0}' ) ", HelperClass1.UserInfo.userid)
        Return myquery
    End Function

    Private Function getProject(ByVal IsAdmin As Boolean)
        Dim myquery As String
        If IsAdmin Then
            myquery = " select distinct projectname as  name from doc.assetpurchase ap " &
                                    " left join doc.toolingproject tp on tp.id = ap.projectid " &
                                    " where not projectname isnull" &
                                    " order by projectname"
        Else
            myquery = String.Format("{0} select distinct projectname as name from uv " &
                                    " left join doc.assetpurchase ap on ap.vendorcode = uv.vendorcode" &
                                    " left join doc.toolingproject tp on tp.id = ap.projectid " &
                                    " where not projectname isnull" &
                                    " order by projectname", getUserVendor)
        End If


        Return myquery
    End Function

    Private Function getSocialAudit(ByVal isAdmin As Boolean)
        Dim myquery As String
        If isAdmin Then
            myquery = String.Format(" select distinct sa.auditgrade as name from (select vd.vendorcode,first_value(d.id) over (partition by vd.vendorcode order by docdate desc) as id  from doc.vendordoc vd" &
                      " left join doc.document d on d.id = vd.documentid" &
                      " where(d.doctypeid = {0})" &
                      " group by vd.vendorcode,d.id) as foo " &
                      "left join doc.socialaudit sa on sa.documentid = id" &
                      " order by sa.auditgrade", SocialAudit)
        Else
            myquery = String.Format("{0} select distinct sa.auditgrade as name from (select uv.vendorcode,first_value(d.id) over (partition by uv.vendorcode order by docdate desc) as id  from uv" &
                                    " left join doc.vendordoc vd on vd.vendorcode = uv.vendorcode" &
                                    " left join doc.document d on d.id = vd.documentid" &
                                    " where d.doctypeid = {1} " &
                                    " group by uv.vendorcode,d.id) as foo " &
                                    " left join doc.socialaudit sa on sa.documentid = id" &
                                    " order by sa.auditgrade", getUserVendor, SocialAudit)
            
        End If

        Return myquery
    End Function
    Private Function getCitiProgram(ByVal isAdmin As Boolean)
        Dim myquery As String
        If isAdmin Then
            myquery = String.Format(" select distinct fpcp as name from doc.citiprogram order by fpcp;")
                     
        Else
            myquery = String.Format("{0} select distinct fpcp as name from uv" &
                                    " left join doc.citiprogram cp on cp.vendorcode = uv.vendorcode" &
                                    " order by fpcp", getUserVendor)
        End If

        Return myquery
    End Function

    Private Function getTurnoverQuery(ByVal fieldname As String, ByVal isAdmin As Boolean, Optional ByVal DisplayField As String = "", Optional ByVal join As String = "", Optional ByVal otherFieldname As String = "") As String
        Dim myQuery As String
        If DisplayField = "" Then
            DisplayField = fieldname
        End If

        If isAdmin Then
            myQuery = String.Format("select distinct {0}::text, {1} as name from doc.turnover tu" &
                                    " {2} " &
                                   " where not {0} isnull and not {0}::text = '' and not {0}::text = '.'" &
                                   " order by {0}::text ", fieldname, DisplayField, join)
        Else
            myQuery = String.Format("with uv as (select distinct v.vendorcode " &
                   " from doc.user u" &
                   " left join doc.groupuser gu on gu.userid = u.id" &
                   " left join doc.groupvendor gv on gv.groupid = gu.groupid" &
                   " left join vendor v on v.vendorcode = gv.vendorcode " &
                   " where lower(u.userid)   ~ '{0}$' " &
                   " order by v.vendorcode), " &
                   " tu as (select distinct tu.vendorcode,{1} {4}" &
                    " from uv " &
                    " left join doc.turnover tu on tu.vendorcode = uv.vendorcode" &
                    " {3}" &
                    " )" &
                    " select distinct {1}::text, {2} as name from tu" &
                    " {3}" &
                    " where not {1} isnull or not {1}::text = '' or not {1}::text = '.'  order by {1}::text", HelperClass1.UserInfo.userid, fieldname, DisplayField, join, otherFieldname)
            '" where not {1} isnull or not {1}::text = '' or not {1}::text = '.'  order by {1}::text", HelperClass1.UserInfo.userid, fieldname, DisplayField, join, otherFieldname)
        End If

        Return myQuery
    End Function

    Private Function getSupplierPanel(ByVal fieldname As String, ByVal IsAdmin As Boolean)
        Dim myquery As String
        If isAdmin Then
            myquery = String.Format(" select distinct {0}::text as name from supplierspanel sp" &
                                    " left join supplierscategory sc on sc.supplierscategoryid = sp.supplierscategoryid" &
                                    " left join doc.panelstatus psfp on psfp.id = sp.fp" &
                                    " left join doc.panelstatus cpfp on cpfp.id = sp.cp" &
                                    " where not {0} isnull order by {0}::text", fieldname)

        Else
            myquery = String.Format("{0} select distinct {1}::text as name from uv" &
                                    " left join supplierspanel sp on sp.vendorcode = uv.vendorcode" &
                                    " left join supplierscategory sc on sc.supplierscategoryid = sp.supplierscategoryid" &
                                    " left join doc.panelstatus psfp on psfp.id = sp.fp" &
                                    " left join doc.panelstatus cpfp on cpfp.id = sp.cp" &
                                    " where not {1} isnull order by {1}::text", getUserVendor, fieldname)

        End If

        Return myquery
    End Function
    Private Function getVendorInfo(ByVal fieldname As String, ByVal isadmin As Boolean)
        Dim myquery As String
        'If isadmin Then
        '    myquery = String.Format(" select distinct {0}::text as name from vendor v" &
        '                            " left join officerseb pm on pm.ofsebid = v.pmid" &
        '                            " left join officerseb spm on spm.ofsebid = v.ssmidpl" &
        '                            " where not {0} isnull order by {0}::text", fieldname)

        'Else
        '    myquery = String.Format("{0} select distinct {1}::text as name from uv" &
        '                            " left join vendor v on v.vendorcode = uv.vendorcode" &
        '                            " left join officerseb pm on pm.ofsebid = v.pmid" &
        '                            " left join officerseb spm on spm.ofsebid = v.ssmidpl" &
        '                            " where not {1} isnull order by {1}::text", getUserVendor, fieldname)
        'End If
        If isadmin Then
            'viewvendorpm replace viewvendorfamilypm
            myquery = String.Format(" select distinct {0} as name from vendor v" &
                                    " left join  doc.viewvendorpm vfp on vfp.vendorcode = v.vendorcode" &
                                    " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid" &
                                    " LEFT JOIN masteruser pm ON pm.id = os.muid" &
                                    " LEFT JOIN officerseb o ON o.ofsebid = os.parent" &
                                    " LEFT JOIN masteruser spm ON spm.id = o.muid" &
                                    " LEFT JOIN doc.vendorgsm vgsm ON vgsm.vendorcode = v.vendorcode " &
                                    " LEFT JOIN officerseb o1 ON o1.ofsebid = vgsm.gsmid" &
                                    " LEFT JOIN masteruser gsm ON gsm.id = o1.muid", fieldname)

        Else
            myquery = String.Format("{0} select distinct {0} as name from vendor uv" &
                                    " left join  doc.viewvendorpm vfp on vfp.vendorcode = v.vendorcode" &
                                    " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid" &
                                    " LEFT JOIN masteruser pm ON pm.id = os.muid" &
                                    " LEFT JOIN officerseb o ON o.ofsebid = os.parent" &
                                    " LEFT JOIN masteruser spm ON spm.id = o.muid" &
                                    " LEFT JOIN doc.vendorgsm vgsm ON vgsm.vendorcode = v.vendorcode " &
                                    " LEFT JOIN officerseb o1 ON o1.ofsebid = vgsm.gsmid" &
                                    " LEFT JOIN masteruser gsm ON gsm.id = o1.muid" &
                                    " where not {1} isnull order by {1}::text", getUserVendor, fieldname)
        End If
        Return myquery
    End Function


    Private Function getVendorTechnology(ByVal fieldname As String, ByVal isadmin As Boolean)
        Dim myquery As String
        If isadmin Then
            myquery = String.Format(" select distinct {0}::text as name from doc.vendortechnology vtec" &                                    
                                    " left join doc.technology tec on tec.id = vtec.technologyid" &
                                    " where not {0} isnull order by {0}::text", fieldname)

        Else
            myquery = String.Format("{0} select distinct {1}::text as name from doc.vendortechnology vtec" &
                                    " left join doc.technology tec on tec.id = vtec.technologyid" &
                                    " where not {1} isnull order by {1}::text", getUserVendor, fieldname)
        End If
        Return myquery
    End Function


    Private Function getFactoryInfo(ByVal fieldname As String, ByVal isadmin As Boolean)
        Dim myquery As String
        If isadmin Then
            myquery = String.Format(" select distinct {0}::text as name from doc.vendorfactory vf" &
                                    " left join doc.factorydtl fd on fd.id = vf.factoryid" &
                                    " left join doc.paramdt prov on prov.paramdtid = fd.provinceid" &
                                    " left join doc.paramdt c on c.paramdtid = fd.countryid" &
                                    " where not {0} isnull order by {0}::text", fieldname)

        Else
            myquery = String.Format("{0} select distinct {1}::text as name from uv" &
                                    " left join doc.vendorfactory vf on vf.vendorcode = uv.vendorcode" &
                                    " left join doc.factorydtl fd on fd.id = vf.factoryid" &
                                    " left join doc.paramdt prov on prov.paramdtid = fd.provinceid" &
                                    " left join doc.paramdt c on c.paramdtid = fd.countryid" &
                                    " where not {1} isnull order by {1}::text", getUserVendor, fieldname)
        End If

        Return myquery
    End Function

    Private Function getSIFIDSheet(ByVal fieldname As String, ByVal doctypeid As Integer, ByVal latest As Boolean, ByVal isadmin As Boolean)
        Dim myquery As String
        Dim VendorDoc As String

        If latest Then
            VendorDoc = String.Format("vd as (select distinct vendorcode, id from (select vd.vendorcode,first_value(d.id) over (partition by vd.vendorcode order by docdate desc) as id from doc.vendordoc vd" &
                        " left join doc.document d on d.id = vd.documentid" &
                        " where(d.doctypeid = {0} or d.doctypeid = 54)" &
                        " group by vd.vendorcode,d.id) as foo" &
                        " order by vendorcode)", doctypeid)

        Else
            VendorDoc = String.Format("vd as (select vd.vendorcode,d.id,d.docdate from doc.vendordoc vd" &
                        " left join doc.document d on d.id = vd.documentid" &
                        " where(d.doctypeid = {0} or d.doctypeid = 54)" &
                        " order by vd.vendorcode)", doctypeid)
        End If


        If isadmin Then
            myquery = String.Format(" with {0} select distinct s.value as name from vd" &
                                    " left join doc.sitx s on s.documentid = vd.id" &
                                    " left join doc.silabel sl on sl.id = s.labelid" &
                                    " where sl.name = '{1}'" &
                                    " and not s.value isnull order by s.value", VendorDoc, fieldname)

        Else
            myquery = String.Format("{0} , {1} select distinct s.value as name from uv" &
                                    " left join vd on vd.vendorcode = uv.vendorcode" &
                                   " left join doc.sitx s on s.documentid = vd.id" &
                                    " left join doc.silabel sl on sl.id = s.labelid" &
                                    " where sl.name = '{2}'" &
                                    " and not s.value isnull order by s.value", getUserVendor, VendorDoc, fieldname)
        End If

        Return myquery
    End Function

    Private Sub GetData(ByVal sender As Object, ByVal e As EventArgs)
       
        Select Case sender.name
            Case "UcTextWithHelperProductType"
                SqlStrHelper = getTurnoverQuery("fpcp", HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperSebBrand"
                SqlStrHelper = getTurnoverQuery("brand", HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperRangeCode"
                SqlStrHelper = getTurnoverQuery("tu.range", HelperClass1.UserInfo.IsAdmin, "tu.range::text || ' - ' || r.rangedesc", "left join range r on r.range = tu.range ")
            Case "UcTextWithHelperCMMF"
                SqlStrHelper = getTurnoverQuery("tu.cmmf", HelperClass1.UserInfo.IsAdmin, "tu.cmmf || ' - ' || mm.materialdesc", " left join materialmaster mm on mm.cmmf = tu.cmmf ")
            Case "UcTextWithHelperCommercialCode"
                SqlStrHelper = getTurnoverQuery("mm.commref", HelperClass1.UserInfo.IsAdmin, "mm.commref || ' - ' || mm.materialdesc ", " left join materialmaster mm on mm.cmmf = tu.cmmf ", ",tu.cmmf")
            Case "UcTextWithHelperProject"
                SqlStrHelper = getProject(HelperClass1.UserInfo.IsAdmin)                
            Case "UcTextWithHelperSocialAudit"
                SqlStrHelper = getSocialAudit(HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperVFP"
                SqlStrHelper = getCitiProgram(HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperSupplierStatus"
                'SqlStrHelper = getTurnoverQuery("pr.paramname", HelperClass1.UserInfo.IsAdmin, "pr.paramname", "left join doc.vendorstatus vs on vs.vendorcode = tu.vendorcode left join doc.paramdt pr on pr.paramdtid = vs.status ")
                SqlStrHelper = getTurnoverQuery("pr.paramname", HelperClass1.UserInfo.IsAdmin, "pr.paramname", "left join doc.vendorstatus vs on vs.vendorcode = tu.vendorcode left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2 ")
            Case "UcTextWithHelperSBU"
                SqlStrHelper = getTurnoverQuery("sbu", HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperFamily"
                SqlStrHelper = getTurnoverQuery("familyname", HelperClass1.UserInfo.IsAdmin, "familyid::text || ' - ' || f.familyname", join:="left join family f on f.familyid = tu.comfam")
            Case "UcTextWithHelperSupplierCategory"
                SqlStrHelper = getSupplierPanel("sc.category", HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperFPPanel"
                SqlStrHelper = getSupplierPanel("psfp.paneldescription", HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperCPPanel"
                SqlStrHelper = getSupplierPanel("cpfp.paneldescription", HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperGSM"
                'SqlStrHelper = getVendorInfo("spm.officersebname", HelperClass1.UserInfo.IsAdmin)
                SqlStrHelper = getVendorInfo("gsm.username", HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperSPM"
                'SqlStrHelper = getVendorInfo("spm.officersebname", HelperClass1.UserInfo.IsAdmin)
                SqlStrHelper = getVendorInfo("spm.username", HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperPM"
                'SqlStrHelper = getVendorInfo("pm.officersebname", HelperClass1.UserInfo.IsAdmin)
                SqlStrHelper = getVendorInfo("pm.username", HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperProvince"
                SqlStrHelper = getFactoryInfo("prov.paramname", HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperCountry"
                SqlStrHelper = getFactoryInfo("c.paramname", HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperMainProductSold"
                SqlStrHelper = getSIFIDSheet("Main Product Sold", SIF, CheckBox1.Checked, HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperMainCustomer"
                SqlStrHelper = getSIFIDSheet("Main Customer", SIF, CheckBox1.Checked, HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperSupplierBrand"
                SqlStrHelper = getSIFIDSheet("Supplier Product Brand", SIF, CheckBox1.Checked, HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperIDSBU"
                SqlStrHelper = getSIFIDSheet("SBU", IDENTITY_SHEET, CheckBox1.Checked, HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperIDFamily"
                SqlStrHelper = getSIFIDSheet("Family Name", IDENTITY_SHEET, CheckBox1.Checked, HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperIDOEMODM"
                SqlStrHelper = getSIFIDSheet("OEM ODM", IDENTITY_SHEET, CheckBox1.Checked, HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperIDMaterial"
                SqlStrHelper = getSIFIDSheet("Material", IDENTITY_SHEET, CheckBox1.Checked, HelperClass1.UserInfo.IsAdmin)
            Case "UcTextWithHelperTechnology"
                SqlStrHelper = getVendorTechnology("tec.technologyname", HelperClass1.UserInfo.IsAdmin)
        End Select




        TextHelper = sender
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoLoadData)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoLoadData()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Helper.")
        Dim mymessage As String = String.Empty

        DSSearch = New DataSet
        If DbAdapter1.TbgetDataSet(SqlStrHelper, DSSearch, mymessage) Then
            Try

            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(7, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If

        ProgressReport(1, String.Format("{0}", "Loading Helper.Done!"))
        ProgressReport(5, "Continuous")
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        For Each ctr As Object In Me.Controls
            If (TypeOf ctr Is UCTextWithHelper) Then
                ctr.cleartext()
            End If
        Next
    End Sub
    Private Function GetMainQuery(ByVal IsAdmin As Boolean)
        Dim myquery As String = String.Empty

        WhereSB.Clear()
        LeftJoinMMSB.Clear()
        LeftJoinProjectSB.Clear()
        WithSB.Clear()
        WithSocialAuditSB.Clear()
        LeftJoinSocialAuditSB.Clear()
        leftJoinVFPSB.Clear()
        LeftJoinSupplierPanelSB.Clear()
        LeftJoinFactorySB.Clear()

        WithSIFSB.Clear()
        WithIDSB.Clear()
        LeftJoinSIFSB.Clear()
        LeftJoinIDSB.Clear()
        withTechnologySB.Clear()
        For Each ctr As Object In Me.Controls
            If (TypeOf ctr Is UCTextWithHelper) Then
                ctr.latestvalue = CheckBox1.Checked
                ctr.doprocess()
            End If
        Next
        'and not tu.fpcp isnull is removed 
        If IsAdmin Then
            'myquery = String.Format("{3} select distinct v.vendorcode, v.shortname::text,v.vendorname::text, sc.category::text,pr.paramname as status ,pm.officersebname::text as pm,spm.officersebname::text as spm {11} from vendor v" &
            '          " left join doc.turnover tu on tu.vendorcode = v.vendorcode" &
            '          " left join supplierspanel sp on sp.vendorcode = v.vendorcode" &
            '          " left join supplierscategory sc on sc.supplierscategoryid = sp.supplierscategoryid" &
            '          " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
            '          " left join doc.paramdt pr on pr.paramdtid = vs.status " &
            '          " left join officerseb pm on pm.ofsebid = v.pmid" &
            '          " {1} {2} {4} {5} {6} {7} {8} {9} {10} " &
            '          " left join officerseb spm on spm.ofsebid = v.ssmidpl {0} ", WhereSB, LeftJoinMMSB.ToString, LeftJoinProjectSB.ToString, WithSB.ToString, LeftJoinSocialAuditSB.ToString, leftJoinVFPSB.ToString, LeftJoinSupplierPanelSB.ToString, LeftJoinFactorySB.ToString, LeftJoinSIFSB.ToString, LeftJoinIDSB.ToString, LeftJoinTechnologySB.ToString, AdditionalDisplayFieldSB.ToString)
            'myquery = String.Format("{3} select distinct v.vendorcode, v.shortname::text,v.vendorname::text, sc.category::text,pr.paramname as status ,pm.officersebname::text as pm,spm.officersebname::text as spm {11} from vendor v" &
            '          " left join doc.turnover tu on tu.vendorcode = v.vendorcode" &
            '          " left join supplierspanel sp on sp.vendorcode = v.vendorcode" &
            '          " left join supplierscategory sc on sc.supplierscategoryid = sp.supplierscategoryid" &
            '          " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
            '          " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2 " &
            '          " left join officerseb pm on pm.ofsebid = v.pmid" &
            '          " {1} {2} {4} {5} {6} {7} {8} {9} {10} " &
            '          " left join officerseb spm on spm.ofsebid = v.ssmidpl {0} ", WhereSB, LeftJoinMMSB.ToString, LeftJoinProjectSB.ToString, WithSB.ToString, LeftJoinSocialAuditSB.ToString, leftJoinVFPSB.ToString, LeftJoinSupplierPanelSB.ToString, LeftJoinFactorySB.ToString, LeftJoinSIFSB.ToString, LeftJoinIDSB.ToString, LeftJoinTechnologySB.ToString, AdditionalDisplayFieldSB.ToString)
            'viewvendorpmeffectivedate replace viewvendorfamilypmeffectivedate
            myquery = String.Format("{3} select distinct v.vendorcode, v.shortname::text,v.vendorname::text, sc.category::text,pr.paramname as status ,gsm.username::text as gsm,vgsm.effectivedate,pm.username::text as pm,vfp.pmeffectivedate,spm.username::text as spm , vfp.spmeffectivedate{11} from vendor v" &
                      " left join doc.turnover tu on tu.vendorcode = v.vendorcode" &
                      " left join supplierspanel sp on sp.vendorcode = v.vendorcode" &
                      " left join supplierscategory sc on sc.supplierscategoryid = sp.supplierscategoryid" &
                      " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
                      " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2 " &
                      " left join doc.viewvendorpmeffectivedate vfp on vfp.vendorcode = v.vendorcode " &
                      " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid " &
                      " LEFT JOIN masteruser pm ON pm.id = os.muid " &
                      " LEFT JOIN officerseb o ON o.ofsebid = os.parent " &
                      " LEFT JOIN masteruser spm ON spm.id = o.muid " &
                      " LEFT JOIN doc.vendorgsm vgsm ON vgsm.vendorcode = v.vendorcode  " &
                      " LEFT JOIN officerseb o1 ON o1.ofsebid = vgsm.gsmid " &
                      " LEFT JOIN masteruser gsm ON gsm.id = o1.muid  " &
                      " {1} {2} {4} {5} {6} {7} {8} {9} {10} {0}", WhereSB, LeftJoinMMSB.ToString, LeftJoinProjectSB.ToString, WithSB.ToString, LeftJoinSocialAuditSB.ToString, leftJoinVFPSB.ToString, LeftJoinSupplierPanelSB.ToString, LeftJoinFactorySB.ToString, LeftJoinSIFSB.ToString, LeftJoinIDSB.ToString, LeftJoinTechnologySB.ToString, AdditionalDisplayFieldSB.ToString)

        Else
            'myquery = String.Format("{0} {4} select distinct v.vendorcode, v.shortname::text,v.vendorname::text,sc.category::text,pr.paramname as status ,pm.officersebname::text as pm,spm.officersebname::text as spm {12} from uv" &
            '                       " left join vendor v on v.vendorcode = uv.vendorcode" &
            '                       " left join doc.turnover tu on tu.vendorcode = v.vendorcode" &
            '                       " left join supplierspanel sp on sp.vendorcode = v.vendorcode" &
            '                       " left join supplierscategory sc on sc.supplierscategoryid = sp.supplierscategoryid" &
            '                       " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
            '                       " left join doc.paramdt pr on pr.paramdtid = vs.status " &
            '                       " left join officerseb pm on pm.ofsebid = v.pmid" &
            '                       " left join officerseb spm on spm.ofsebid = v.ssmidpl {2} {3} {5} {6} {7} {8} {9} {10} {11} {1} ", getUserVendor, WhereSB.ToString, LeftJoinMMSB.ToString, LeftJoinProjectSB.ToString, WithSB.ToString.Replace("with", ","), LeftJoinSocialAuditSB.ToString, leftJoinVFPSB.ToString, LeftJoinSupplierPanelSB.ToString, LeftJoinFactorySB.ToString, LeftJoinSIFSB.ToString, LeftJoinIDSB.ToString, LeftJoinTechnologySB.ToString, AdditionalDisplayFieldSB.ToString)
            'myquery = String.Format("{0} {4} select distinct v.vendorcode, v.shortname::text,v.vendorname::text,sc.category::text,pr.paramname as status ,pm.officersebname::text as pm,spm.officersebname::text as spm {12} from uv" &
            '                      " left join vendor v on v.vendorcode = uv.vendorcode" &
            '                      " left join doc.turnover tu on tu.vendorcode = v.vendorcode" &
            '                      " left join supplierspanel sp on sp.vendorcode = v.vendorcode" &
            '                      " left join supplierscategory sc on sc.supplierscategoryid = sp.supplierscategoryid" &
            '                      " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
            '                      " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
            '                      " left join officerseb pm on pm.ofsebid = v.pmid" &
            '                      " left join officerseb spm on spm.ofsebid = v.ssmidpl {2} {3} {5} {6} {7} {8} {9} {10} {11} {1} ", getUserVendor, WhereSB.ToString, LeftJoinMMSB.ToString, LeftJoinProjectSB.ToString, WithSB.ToString.Replace("with", ","), LeftJoinSocialAuditSB.ToString, leftJoinVFPSB.ToString, LeftJoinSupplierPanelSB.ToString, LeftJoinFactorySB.ToString, LeftJoinSIFSB.ToString, LeftJoinIDSB.ToString, LeftJoinTechnologySB.ToString, AdditionalDisplayFieldSB.ToString)
            'viewvendorpmeffectivedate replace viewvendorfamilypmeffectivedate
            myquery = String.Format("{0} {4} select distinct v.vendorcode, v.shortname::text,v.vendorname::text,sc.category::text,pr.paramname as status ,gsm.username::text as gsm,vgsm.effectivedate,pm.username::text as pm,vfp.pmeffectivedate,spm.username::text as spm,vfp.spmeffectivedate {12} from uv" &
                                  " left join vendor v on v.vendorcode = uv.vendorcode" &
                                  " left join doc.turnover tu on tu.vendorcode = v.vendorcode" &
                                  " left join supplierspanel sp on sp.vendorcode = v.vendorcode" &
                                  " left join supplierscategory sc on sc.supplierscategoryid = sp.supplierscategoryid" &
                                  " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
                                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                                  " left join doc.viewvendorpmeffectivedate vfp on vfp.vendorcode = v.vendorcode " &
                                  " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid " &
                                  " LEFT JOIN masteruser pm ON pm.id = os.muid " &
                                  " LEFT JOIN officerseb o ON o.ofsebid = os.parent " &
                                  " LEFT JOIN masteruser spm ON spm.id = o.muid " &
                                  " LEFT JOIN doc.vendorgsm vgsm ON vgsm.vendorcode = v.vendorcode  " &
                                  " LEFT JOIN officerseb o1 ON o1.ofsebid = vgsm.gsmid " &
                                  " LEFT JOIN masteruser gsm ON gsm.id = o1.muid   {2} {3} {5} {6} {7} {8} {9} {10} {11} {1} ", getUserVendor, WhereSB.ToString, LeftJoinMMSB.ToString, LeftJoinProjectSB.ToString, WithSB.ToString.Replace("with", ","), LeftJoinSocialAuditSB.ToString, leftJoinVFPSB.ToString, LeftJoinSupplierPanelSB.ToString, LeftJoinFactorySB.ToString, LeftJoinSIFSB.ToString, LeftJoinIDSB.ToString, LeftJoinTechnologySB.ToString, AdditionalDisplayFieldSB.ToString)


        End If
        Return myquery
    End Function

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        loaddata()       
    End Sub


    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Button29.PerformClick()
    End Sub




    Private Sub FormAdvancedSearchSupplier_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        CheckBox2_CheckedChanged(sender, e)
    End Sub



End Class