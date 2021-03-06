﻿Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Imports Microsoft.Office.Interop
Public Enum TypeOfInvestmentEnum
    NewAsset = 1
    ToolingModification = 2
    BackUp = 3
    IncreaseCapacity = 4
End Enum
Public Class FormAssetsPurchase
    Dim myProformaPO As proformaPO
    Public Property myId As Long
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myThreadDownload As New System.Threading.Thread(AddressOf DoDownload)
    Dim SelectedFolder As String = String.Empty
    Dim myUser As String = String.Empty

    Dim DS As DataSet
    Dim sb As StringBuilder

    Dim WithEvents APBS As BindingSource
    Dim WithEvents AttachmentBS As BindingSource

    Dim ToolingInvoiceBS As BindingSource
    Dim ToolingPaymentbs As BindingSource
    Dim TypeOfInvestmentList As List(Of TypeOfInvestment)
    Dim DocumentTypeList As List(Of DocumentType)
    Dim PaymentMethodList As List(Of PaymentMethod)



    Dim TypeOfInvestmentBS As BindingSource
    Dim DocumentTypeBS As BindingSource
    Dim PaymentMethodBS As BindingSource


    Dim VendorBS As BindingSource
    Dim VendorHelperBS As BindingSource
    Dim ToolingProjectHelperBS As BindingSource
    Dim ApplicantBS As BindingSource
    Dim ApprovalBS As BindingSource
    Dim ApprovalBS2 As BindingSource
    Dim FamilyGroupHelperBS As BindingSource
    Dim ToolingProjectDict As Dictionary(Of String, Long)
    Dim ToolingListDTBS As BindingSource
    Dim TrackingBS As BindingSource
    Dim AgreementTXBS As BindingSource
    Dim AssetPurchaseActionBS As BindingSource
    Dim ParamBS As BindingSource
    Dim CM1 As CurrencyManager

    Dim ProformaInvoiceBS As BindingSource
    Dim OriginalCost As Decimal
    Dim OriginalCurrency As String
    Dim FromApproval As Boolean = False
    Dim FromApprovalHistory As Boolean = False


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByVal myid As Long, ByVal status As Integer, ByVal FromApprovalHistory As Boolean)
        InitializeComponent()
        Me.FromApprovalHistory = FromApprovalHistory
        Me.FromApproval = True
        initData()
        Me.myId = myid
        If HelperClass1.UserInfo.IsFinance Then
            ToolStripButton1.Visible = False
            ToolStripButton2.Visible = False
            ToolStripButton5.Visible = False
            DataGridView2.ContextMenuStrip = Nothing
            DataGridView3.ContextMenuStrip = Nothing
        End If
    End Sub
    Public Sub New(ByVal myid As Long, ByVal status As Integer)
        InitializeComponent()
        'If status = AssetPurchaseStatusEnum.StatusNew Or status = AssetPurchaseStatusEnum.StatusReSubmit Then
        '    'FromApproval = False
        '    FromApproval = True
        'Else
        '    FromApproval = True
        'End If

        FromApproval = True

        initData()
        Me.myId = myid
        If HelperClass1.UserInfo.IsFinance Then
            ToolStripButton1.Visible = False
            ToolStripButton2.Visible = False
            ToolStripButton5.Visible = False
            'ToolStripButton2.Visible = True
            'Button10.Enabled = True
            'Button11.Enabled = True
            DataGridView2.ContextMenuStrip = Nothing
            DataGridView3.ContextMenuStrip = Nothing
        End If        
    End Sub

    Public Sub New(ByVal myId As Long)
        InitializeComponent()
        initData()
        Me.myId = myId

        If HelperClass1.UserInfo.IsFinance Then
            ToolStripButton1.Visible = False
            ToolStripButton2.Visible = False            
            DataGridView2.ContextMenuStrip = Nothing
            DataGridView3.ContextMenuStrip = Nothing
        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.WParam.ToInt32 = &HF060 Then Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        MyBase.WndProc(m)
    End Sub

    Private Sub FormAssetsPurchase_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If (Not (HelperClass1.UserInfo.IsFinance)) And (Not FromApproval) Then
            Dim abc = DS.GetChanges()
            If Not IsNothing(abc) Then

                Select Case MessageBox.Show("Save unsaved records?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                    Case Windows.Forms.DialogResult.Yes
                        If Me.validate Then
                            ToolStripButton2.PerformClick()
                        Else
                            e.Cancel = True
                        End If

                    Case Windows.Forms.DialogResult.Cancel
                        e.Cancel = True
                End Select
            End If
        End If
        
    End Sub

    Private Sub FormAssetsPurchase_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        
        'initData()
        Loadata()

        'Me.Show()
        'Application.DoEvents()
        'TextBox1.Focus()
        If (Not HelperClass1.UserInfo.IsAdmin) Or (FromApproval) Then
            'disabled menu item
            ToolStripButton3.Visible = False
            UpdateToolingIdToolStripMenuItem.Visible = False
        End If
    End Sub
    Sub Loadata()
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
        sb = New StringBuilder
        DS = New DataSet
        Dim mymessage As String = String.Empty

        '0
        sb.Append(String.Format("with v as" &
                                " (select vendorcode,(vendorcode::text || ' - ' || vendorname::text) as description,vendorname::text from vendor order by vendorname)," &
                                " f as(select (s.sbuname2 || ' - ' || f.familyid || ' - ' || upper(familyname) || '(' || fg.groupingcode || ')') as description ,s.sbuname2,f.familyid from doc.familygroupsbu fg" &
                                " left join family f on f.familyid = fg.familyid" &
                                " left join sbusap s on s.sbuid = fg.sbusapid" &
                                " where not familyname isnull and  not fg.groupingcode isnull" &
                                " order by sbuname,familyid,familyname)," &
                                " c as (select sum(originalcost * ap.otcexrate) as toolingcost,t.assetpurchaseid from doc.toolinglist t left join doc.assetpurchase ap on ap.id = t.assetpurchaseid where t.assetpurchaseid = {0} group by t.assetpurchaseid)" &
                                " select ap.*,case ap.paymentmethodid when 1 then 'Amortization' when 2 then 'Invoice Investment' end as paymentmethod,c.toolingcost,tp.projectcode,tp.projectname,tp.dept,tp.ppps,tp.sbuid,tp.familyid,v.vendorname,v.description as vendordescription,f.description as familydescription,f.sbuname2,u.username," &
                                " ts.toolingsuppliername,ts.address,ts.deliveryaddress,ts.fax,ts.tel,ts.toolingsupplierid || ' - ' || ts.toolingsuppliername as toolingsupplierdescription" &
                                " ,u1.email as approvalemail, u2.email as approvalemail2,va.shortname,coalesce(va.street,'') || coalesce(va.city,'') as vaaddress,va.telephone,va.faxnumber " &
                                " from doc.assetpurchase ap" &
                                " left join doc.toolingproject tp on tp.id = ap.projectid" &
                                " left join  v on v.vendorcode = ap.vendorcode" &
                                " left join f on f.familyid = tp.familyid" &
                                " left join c on c.assetpurchaseid = ap.id" &
                                " left join doc.user u on u.userid = ap.creator" &
                                " left join doc.user u1 on u1.username = ap.approvalname" &
                                " left join doc.user u2 on u1.username = ap.approvalname2" &
                                " left join doc.toolingsupplier ts on ts.toolingsupplierid = ap.toolingsupplier" &
                                " left join doc.vendoraddress va on va.vendorcode = ap.vendorcode" &
                                " where ap.id = {0};", myId))
        'sb.Append(String.Format("select tldt.*,tlhd.assetpurchaseid,tlhd.sebmodelno,tlhd.suppliermodelreference,tlhd.purchasedate,tlhd.location from doc.toolinglistdt tldt left join doc.toolinglisthd tlhd on tldt.toolinglisthdid = tlhd.id where assetpurchaseid = {0} order by lineno;", myId))
        'sb.Append(String.Format("select tl.*,ap.assetpurchaseid || to_char(lineno,'_0000FM') as toolingid from doc.toolinglist tl left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid  where ap.id = {0} order by lineno;", myId))
        'sb.Append(String.Format("select tl.*,0 as typeofinvestment from doc.toolinglist tl left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid  where ap.id = {0} order by lineno;", myId))

        '1
        sb.Append(String.Format("with cost as (select sum(tp.invoiceamount * ap.otcexrate) as total,tp.toolinglistid  from doc.toolingpayment tp" &
                        " left join doc.toolinglist tl on tl.id = tp.toolinglistid" &
                        " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid" &
                        " where(ap.id = {0})" &
                        " group by tp.toolinglistid" &
                        " ) " &
                        "  ,costcny as  (select sum(tp.invoiceamount) as totalcny,tp.toolinglistid from doc.toolingpayment tp " &
                        " left join doc.toolinglist tl on tl.id = tp.toolinglistid " &
                        " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid where(ap.id = {0}) and currency = 'CNY'" &
                        " group by tp.toolinglistid )" &
                        " select tl.id,tl.assetpurchaseid,tm.sebmodelno,tm.suppliermodelreference,tm.suppliermoldno,tm.toolsdescription,tm.material,tm.cavities,tm.numberoftools,tm.dailycapacity,tl.originalcurrency,tl.originalcost,tl.originalcost * ap.otcexrate as cost,tm.purchasedate ,tm.location,tm.comments ,tl.lineno,tl.vendorcode,tl.toolinglistid,tm.dailycaps,tm.commontool ,0 as typeofinvestment ," &
                        " case" &
                        " when c.total isnull then tl.originalcost * ap.otcexrate " &
                        " else (tl.originalcost * ap.otcexrate ) - c.total " &
                        " end as balance ," &
                        " case when cn.totalcny isnull then tl.originalcost  else tl.originalcost - cn.totalcny  end as balancecny ," &
                        " tl.toolinglistid || ' - ' || tl.suppliermoldno || ' - ' || tl.toolsdescription as displaymember from doc.toolinglist tl " &
                        " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid  " &
                        " left join cost c on c.toolinglistid = tl.id" &
                        " left join costcny cn on cn.toolinglistid = tl.id " &
                        " left join doc.toolinglistmst tm on tm.toolinglistid = tl.toolinglistid" &
                        " where ap.id = {0} order by lineno;", myId))

        'sb.Append(String.Format("select tp.*,tl.toolinglistid as displaymember,tl.suppliermoldno,null::numeric as pct,tl.toolsdescription from doc.toolingpayment tp" &
        '                        " left join doc.toolinglist tl on tl.id = tp.toolinglistid" &
        '                        " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid" &

        '                        " where(ap.id = {0});", myId))

        '2
        sb.Append(String.Format("with cost as (" &
                                " select sum(tp.invoiceamount * tp.exrate) as total,tp.toolinglistid  from doc.toolingpayment tp " &
                                " left join doc.toolinglist tl on tl.id = tp.toolinglistid " &
                                " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid " &
                                " where(ap.id = {0}) group by tp.toolinglistid )  " &
                                " select tp.*,tl.toolinglistid as displaymember,tl.suppliermoldno,tl.toolsdescription," &
                                "  tl.cost * ap.otcexrate as cost,case when c.total isnull then tl.originalcost * ap.otcexrate   else (tl.originalcost * ap.otcexrate ) - c.total  end as balance,(tp.invoiceamount * tp.exrate / (case when tl.cost = 0 then 1 else tl.cost end)) * 100 as pct ,tp.invoiceamount * tp.exrate as invoiceamountusd from doc.toolingpayment tp " &
                                " left join doc.toolinglist tl on tl.id = tp.toolinglistid " &
                                " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid " &
                                " left join cost c on c.toolinglistid = tl.id" &
                                " where(ap.id = {0});", myId))
        '3
        sb.Append(String.Format(" with tot as (select ti.id,sum(tp.exrate * tp.invoiceamount) as totalamount from  " &
                                " doc.toolinginvoice ti" &
                                " left join doc.toolingpayment tp on tp.invoiceid = ti.id" &
                                " left join doc.toolinglist tl on tl.id = tp.toolinglistid" &
                                " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid" &
                                " where (ap.id = {0}) " &
                                " group by ti.id )," &
                                " totcny as (select ti.id,sum(tp.invoiceamount) as totalamountcny from  " &
                                " doc.toolinginvoice ti" &
                                " left join doc.toolingpayment tp on tp.invoiceid = ti.id" &
                                " left join doc.toolinglist tl on tl.id = tp.toolinglistid" &
                                " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid" &
                                " where (ap.id = {0}) and tp.currency='CNY' " &
                                " group by ti.id )" &
                                " select distinct ti.*,tot.totalamount,totcny.totalamountcny,null::numeric as pct,ap.id as apid from doc.toolinginvoice ti" &
                                " left join doc.toolingpayment tp on tp.invoiceid = ti.id" &
                                " left join doc.toolinglist tl on tl.id = tp.toolinglistid" &
                                " left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid" &
                                " left join tot on tot.id = ti.id" &
                                " left join totcny on totcny.id = ti.id" &
                                " where(ap.id = {0}) order by invoicedate desc,invoiceno;", myId))
        '4
        sb.Append("with f as (select (s.sbuname2 || ' - ' || f.familyid || ' - ' || upper(familyname::text) || '(' || fg.groupingcode || ')') as familydescription ,s.sbuname2,f.familyid,f.familyname::Text,fg.groupingcode from doc.familygroupsbu fg" &
                  " left join family f on f.familyid = fg.familyid" &
                  " left join sbusap s on s.sbuid = fg.sbusapid" &
                  " where not familyname isnull and  not fg.groupingcode isnull " &
                  " order by sbuname,familyid,familyname)" &
                  " select tp.*,f.*,(tp.projectcode || ' - ' || tp.projectname) as projectdescription from doc.toolingproject tp left join f on f.familyid = tp.familyid order by tp.projectcode;")

        '5
        sb.Append("with f as (select (s.sbuname2 || ' - ' || f.familyid || ' - ' || upper(familyname::text) || '(' || fg.groupingcode || ')') as description ,s.sbuname2,f.familyid,f.familyname::text,fg.groupingcode from doc.familygroupsbu fg" &
                  " left join family f on f.familyid = fg.familyid" &
                  " left join sbusap s on s.sbuid = fg.sbusapid" &
                  " where not familyname isnull and  not fg.groupingcode isnull and isdeleted isnull " &
                  " order by sbuname,familyid,familyname)" &
                  " select * from f;")
        '6
        If HelperClass1.UserInfo.IsAdmin Or HelperClass1.UserInfo.IsFinance Then
            sb.Append(" select null::bigint as vendorcode,''::text as description,''::text as vendorname union all (select vendorcode,(vendorcode::text || ' - ' || vendorname::text) as description,vendorname::text from vendor order by vendorname);")
        Else

            sb.Append(String.Format("select null as vendorcode,''::text as description,''::text as vendorname,null::text as shortname union all (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text  from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid" &
                    " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{0}$' and not shortname isnull  order by vendorname);", HelperClass1.UserInfo.userid))
        End If

       


        '7
        sb.Append(" select null::integer as typeofinvestmentid,''::text as typeofinvestmentname union all(select d.ivalue,d.paramname from doc.paramdt d left join doc.paramhd h on h.paramhdid = d.paramhdid where h.paramname = 'typeofinvestment');")
        '8
        sb.Append(String.Format(" select * from doc.toolinglist where assetpurchaseid = {0} ;", myId))
        '9
        'sb.Append("select ''::text as username union all (select username from doc.user where userid  <> 'as\dlie' and isactive and not isfinance order by username); ")
        sb.Append("select ''::text as username union all (select u.username from doc.user u left join officerseb o on o.userid = u.userid left join masteruser mu on mu.id = o.muid" &
                  " left join teamtitle tt on tt.teamtitleid = o.teamtitleid where teamtitleshortname in ('PD','SPM','PM','PO','WMF','PCL') and u.isactive  order by u.username);")
        '10
        sb.Append(String.Format("with d as (select d.ivalue as id,d.paramname as doctypename from doc.paramdt d left join doc.paramhd h on h.paramhdid = d.paramhdid where h.paramname = 'attachmenttype') select false as download,at.*, d.doctypename,'' as docfullname from doc.assetattachment at left join d on d.id = at.doctypeid where assetpurchaseid = {0} ;", myId))
        '11
        sb.Append("select h.cvalue,h.paramname from doc.paramhd h where h.paramname in ('attachmentfolder','assettemplate') order by ivalue;")

        '12
        sb.Append(String.Format("select * from doc.assetpurchasetracking where assetpurchaseid = {0};", myId))

        '13
        'sb.Append(String.Format("select tr.trackingno, ag.agreement,dt.material,shorttext,startdate,enddate,ag.totalqty,doc.getqty(dt.material,startdate,doc.addyear(startdate,1)) as c1" &
        '          " ,doc.getqty(dt.material,doc.addyear(startdate,1),doc.addyear(startdate,2)) as c2, doc.getqty(dt.material, doc.addyear(startdate, 2), doc.addyear(startdate, 3)) as c3" &
        '          " from doc.assetpurchase ap left join doc.assetpurchasetracking tr on tr.assetpurchaseid = ap.id left join agvalue ag on ag.trackingno = tr.trackingno" &
        '          " left join agreementdt dt on dt.agreement = ag.agreement where  ag.totalqty <> 0 and ap.id = {0}", myId))

        sb.Append(String.Format("select tr.trackingno, ag.agreement,dt.material,shorttext,startdate,enddate,ag.totalqty,doc.getqty(dt.material,startdate,doc.addyear(startdate,1)) as c1" &
                  " ,doc.getqty(dt.material,doc.addyear(startdate,1),doc.addyear(startdate,2)) as c2, doc.getqty(dt.material, doc.addyear(startdate, 2), doc.addyear(startdate, 3)) as c3" &
                  " from doc.assetpurchase ap left join doc.assetpurchasetracking tr on tr.assetpurchaseid = ap.id left join agvalue ag on ag.trackingno = tr.trackingno" &
                  " left join agreementdt dt on dt.agreement = ag.agreement where ap.id = {0};", myId))
        '14 Table AssetPurchaseAction get blank table for validation purpose
        sb.Append(String.Format("select *,doc.getassetpurchasestatusname(status) as statusname from doc.assetpurchaseaction where assetpurchaseid = {0} order by id desc;", myId))
        myUser = HelperClass1.UserId.ToLower
        '15 FormInvoiceCurrencyList
        sb.Append(String.Format("select cvalue from doc.paramhd where paramname = 'forminputinvoicecurrencylist';"))
        '16 ProformaInvoice
        sb.Append(String.Format("select false as download,p.* from doc.proformainvoice p where assetpurchaseid = {0};", myId))

        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                Dim pk(0) As DataColumn
                pk(0) = DS.Tables(0).Columns("id")
                DS.Tables(0).PrimaryKey = pk
                DS.Tables(0).Columns("id").AutoIncrement = True
                DS.Tables(0).Columns("id").AutoIncrementSeed = -1
                DS.Tables(0).Columns("id").AutoIncrementStep = -1
                DS.Tables(0).TableName = "AssetPurchase"

                Dim pk1(0) As DataColumn
                pk1(0) = DS.Tables(1).Columns("id")
                DS.Tables(1).PrimaryKey = pk1
                DS.Tables(1).Columns("id").AutoIncrement = True
                DS.Tables(1).Columns("id").AutoIncrementSeed = -1
                DS.Tables(1).Columns("id").AutoIncrementStep = -1
                DS.Tables(1).TableName = "ToolingListDT"

                Dim pk2(0) As DataColumn
                pk2(0) = DS.Tables(2).Columns("id")
                DS.Tables(2).PrimaryKey = pk2
                DS.Tables(2).Columns("id").AutoIncrement = True
                DS.Tables(2).Columns("id").AutoIncrementSeed = -1
                DS.Tables(2).Columns("id").AutoIncrementStep = -1
                DS.Tables(2).TableName = "ToolingPayment"
                Dim uidx2(1) As DataColumn
                uidx2(0) = DS.Tables(2).Columns("invoiceid")
                uidx2(1) = DS.Tables(2).Columns("displaymember")
                DS.Tables(2).Constraints.Add(New UniqueConstraint(uidx2))

                Dim pk3(0) As DataColumn
                pk3(0) = DS.Tables(3).Columns("id")
                DS.Tables(3).PrimaryKey = pk3
                DS.Tables(3).Columns("id").AutoIncrement = True
                DS.Tables(3).Columns("id").AutoIncrementSeed = -1
                DS.Tables(3).Columns("id").AutoIncrementStep = -1
                DS.Tables(3).TableName = "ToolingInvoice"
                Dim uidx3(1) As DataColumn
                uidx3(0) = DS.Tables(3).Columns("apid")
                uidx3(1) = DS.Tables(3).Columns("invoiceno")
                DS.Tables(3).Constraints.Add(New UniqueConstraint(uidx3))

                Dim pk4(0) As DataColumn
                pk4(0) = DS.Tables(4).Columns("id")
                DS.Tables(4).PrimaryKey = pk4
                DS.Tables(4).Columns("id").AutoIncrement = True
                DS.Tables(4).Columns("id").AutoIncrementSeed = -1
                DS.Tables(4).Columns("id").AutoIncrementStep = -1
                DS.Tables(4).TableName = "ToolingProject"


                DS.Tables(5).TableName = "FamilyGroup"
                DS.Tables(6).TableName = "Vendor"
                DS.Tables(7).TableName = "TypeOfInvestment"
                DS.Tables(8).TableName = "ToolingListHD"
                DS.Tables(9).TableName = "Applicant"

                Dim pk10(0) As DataColumn
                pk10(0) = DS.Tables(10).Columns("id")
                DS.Tables(10).PrimaryKey = pk10
                DS.Tables(10).Columns("id").AutoIncrement = True
                DS.Tables(10).Columns("id").AutoIncrementSeed = -1
                DS.Tables(10).Columns("id").AutoIncrementStep = -1
                DS.Tables(10).TableName = "Attachment"

                DS.Tables(11).TableName = "ParamName"

                Dim pk12(0) As DataColumn
                pk12(0) = DS.Tables(12).Columns("id")
                DS.Tables(12).PrimaryKey = pk12
                DS.Tables(12).Columns("id").AutoIncrement = True
                DS.Tables(12).Columns("id").AutoIncrementSeed = -1
                DS.Tables(12).Columns("id").AutoIncrementStep = -1
                DS.Tables(12).TableName = "TrackingNo"
                DS.Tables(13).TableName = "AgreementTX"
                DS.Tables(14).TableName = "AssetPurchaseAction"

                DS.Tables(16).TableName = "ProformaInvoice"
                Dim pk16(0) As DataColumn
                pk16(0) = DS.Tables(16).Columns("id")
                DS.Tables(16).PrimaryKey = pk16
                DS.Tables(16).Columns("id").AutoIncrement = True
                DS.Tables(16).Columns("id").AutoIncrementSeed = -1
                DS.Tables(16).Columns("id").AutoIncrementStep = -1

                'Create Constraint -- create Constraints before you create DataRelation
                'DS.Tables(2).Constraints.Add("ToolingPaymentFK", DS.Tables(3).Columns("id"), DS.Tables(2).Columns("invoiceid"))
                'Do not create Constraint!! Use relations instead!

                'RELATIONS
                Dim rel As DataRelation
                Dim TPCol As DataColumn
                Dim APCol As DataColumn
                Dim TLCol As DataColumn
                Dim TICol As DataColumn
                Dim ATCol As DataColumn
                Dim TRCol As DataColumn
                Dim PICol As DataColumn 'ProformaInvoice


                'Create Relation ToolingProject -> Asset Purchase
                TPCol = DS.Tables(4).Columns("id")
                APCol = DS.Tables(0).Columns("projectid")
                rel = New DataRelation("TPAPRel", TPCol, APCol)
                DS.Relations.Add(rel)

                'create relation AssetPurchase and ToolingList
                APCol = DS.Tables(0).Columns("id")
                TLCol = DS.Tables(1).Columns("assetpurchaseid")
                rel = New DataRelation("APTLRel", APCol, TLCol)
                DS.Relations.Add(rel)

                'create relation ToolingList -> toolingpayment
                TLCol = DS.Tables(1).Columns("id")
                TPCol = DS.Tables(2).Columns("toolinglistid")
                rel = New DataRelation("TLTPRel", TLCol, TPCol)
                DS.Relations.Add(rel)

                'create relation ToolingInvoice -> toolingpayment
                TICol = DS.Tables(3).Columns("id")
                TPCol = DS.Tables(2).Columns("invoiceid")
                rel = New DataRelation("TITPRel", TICol, TPCol)
                DS.Relations.Add(rel)

                'create relation AssetPurchase and Attachment
                APCol = DS.Tables(0).Columns("id")
                ATCol = DS.Tables(10).Columns("assetpurchaseid")
                rel = New DataRelation("APATRel", APCol, ATCol)
                DS.Relations.Add(rel)

                'create relation AssetPurchase and ToolingList
                APCol = DS.Tables(0).Columns("id")
                TRCol = DS.Tables(12).Columns("assetpurchaseid")
                rel = New DataRelation("APTRRel", APCol, TRCol)
                DS.Relations.Add(rel)

                'create relation AssetPurchase and ProformaInvoice
                APCol = DS.Tables(0).Columns("id")
                PICol = DS.Tables(16).Columns("assetpurchaseid")
                rel = New DataRelation("APPIRel", APCol, PICol)
                DS.Relations.Add(rel)

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
        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Try
                Me.Invoke(d, New Object() {id, message})
            Catch ex As Exception

            End Try
        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel1.Text = message
                    Case 4
                        Try
                            APBS = New BindingSource
                            TypeOfInvestmentBS = New BindingSource

                            ToolingInvoiceBS = New BindingSource
                            ToolingPaymentbs = New BindingSource
                            PaymentMethodBS = New BindingSource

                            VendorBS = New BindingSource
                            VendorHelperBS = New BindingSource
                            ToolingProjectHelperBS = New BindingSource
                            FamilyGroupHelperBS = New BindingSource
                            ToolingListDTBS = New BindingSource
                            AttachmentBS = New BindingSource
                            ApplicantBS = New BindingSource
                            ApprovalBS = New BindingSource
                            ApprovalBS2 = New BindingSource
                            DocumentTypeBS = New BindingSource
                            ParamBS = New BindingSource
                            TrackingBS = New BindingSource
                            AgreementTXBS = New BindingSource
                            AssetPurchaseActionBS = New BindingSource

                            ProformaInvoiceBS = New BindingSource

                            APBS.DataSource = DS.Tables("AssetPurchase")



                            TypeOfInvestmentBS.DataSource = TypeOfInvestmentList 'DS.Tables("TypeOfInvestment") 'TypeOfInvestmentList
                            DocumentTypeBS.DataSource = DocumentTypeList
                            PaymentMethodBS.DataSource = PaymentMethodList

                            VendorBS.DataSource = DS.Tables("Vendor")
                            VendorHelperBS.DataSource = New DataView(DS.Tables("Vendor"))
                            ToolingProjectHelperBS.DataSource = DS.Tables("ToolingProject")
                            FamilyGroupHelperBS.DataSource = DS.Tables("FamilyGroup")
                            ToolingListDTBS.DataSource = DS.Tables("ToolingListDT")
                            ApplicantBS.DataSource = DS.Tables("Applicant")
                            ApprovalBS.DataSource = DS.Tables("Applicant")
                            ApprovalBS2.DataSource = DS.Tables("Applicant")

                            ToolingInvoiceBS.DataSource = DS.Tables("ToolingInvoice")
                            ToolingPaymentbs.DataSource = DS.Tables("ToolingPayment")
                            AttachmentBS.DataSource = DS.Tables("Attachment")

                            ParamBS.DataSource = DS.Tables("ParamName")
                            TrackingBS.DataSource = DS.Tables("TrackingNo")
                            AgreementTXBS.DataSource = DS.Tables("AgreementTx")
                            AssetPurchaseActionBS.DataSource = DS.Tables("AssetPurchaseAction")

                            ProformaInvoiceBS.DataSource = APBS
                            ProformaInvoiceBS.DataMember = "APPIRel"

                            If IsNothing(APBS.Current) Then
                                Dim drv As DataRowView = APBS.AddNew()
                                'drv.BeginEdit()
                                drv.Row.Item("dept") = "Purchasing"
                                drv.Row.Item("categoryofasset") = 1
                                drv.Row.Item("applicantdate") = Today.Date
                                drv.Row.Item("budgetcurr") = "USD"
                                drv.Row.Item("exchangerate") = 1
                                drv.Row.Item("creator") = myUser
                                drv.Row.Item("username") = HelperClass1.UserInfo.DisplayName
                                drv.Row.Item("status") = AssetPurchaseStatusEnum.StatusDraft
                                'APBS.EndEdit()
                            Else
                                'sync APBS.Current ProjectId with ToolingPorjectHelper ID
                                Dim myapbs As DataRowView = APBS.Current
                                Dim myposition = ToolingProjectHelperBS.Find("id", myapbs.Item("projectid"))
                                ToolingProjectHelperBS.Position = myposition
                            End If

                            ClearDataBinding()
                            FillDictionary()
                            AddDataBinding()
                            If IsNothing(AttachmentBS.Current) Then
                                ComboBox3.SelectedIndex = -1
                            End If
                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = ToolingListDTBS

                            DataGridView3.AutoGenerateColumns = False
                            DataGridView3.DataSource = ToolingInvoiceBS

                            DataGridView2.AutoGenerateColumns = False
                            DataGridView2.DataSource = AttachmentBS

                            DataGridView4.AutoGenerateColumns = False
                            DataGridView4.DataSource = AgreementTXBS

                            DataGridView5.AutoGenerateColumns = False
                            DataGridView5.DataSource = AssetPurchaseActionBS

                            DataGridView6.AutoGenerateColumns = False
                            DataGridView6.DataSource = ProformaInvoiceBS

                            Me.Show()
                            Application.DoEvents()
                            TextBox1.Focus()
                            CM1 = CType(BindingContext(ToolingInvoiceBS), CurrencyManager)


                            'If HelperClass1.UserInfo.IsFinance Or FromApproval Then
                            If HelperClass1.UserInfo.IsFinance Then 'Or FromApprovalHistory 

                                ToolStripButton1.Visible = False
                                'ToolStripButton2.Visible = True 
                                'GroupBox1.Enabled = False
                                For Each myobj In TabControl1.Controls
                                    If myobj.GetType() Is GetType(TabPage) Then
                                        For Each obj As Control In myobj.Controls
                                            If obj.GetType() Is GetType(GroupBox) Then
                                                For Each c As Control In obj.Controls
                                                    If c.GetType() Is GetType(TextBox) Then
                                                        Dim Txt As TextBox = CType(c, TextBox)
                                                        Txt.BackColor = Color.WhiteSmoke
                                                        Txt.ReadOnly = True
                                                    End If
                                                    If c.GetType() Is GetType(DateTimePicker) Then
                                                        Dim dt As DateTimePicker = CType(c, DateTimePicker)
                                                        dt.Enabled = False
                                                        dt.CalendarMonthBackground = Color.WhiteSmoke
                                                        dt.CalendarForeColor = Color.Black
                                                    End If
                                                    If c.GetType() Is GetType(Button) Then
                                                        Dim b As Button = CType(c, Button)
                                                        b.Enabled = False
                                                    End If
                                                    If c.GetType() Is GetType(ComboBox) Then
                                                        Dim cb As ComboBox = CType(c, ComboBox)
                                                        cb.Enabled = False
                                                        cb.ForeColor = Color.Black
                                                        cb.BackColor = Color.WhiteSmoke

                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                Next
                                DataGridView2.ContextMenuStrip = Nothing
                                DataGridView3.ContextMenuStrip = Nothing
                                Button10.Enabled = True
                                Button11.Enabled = True
                                If FromApprovalHistory Then
                                    Button10.Enabled = False
                                    Button11.Enabled = False
                                End If
                            End If
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        'Save ToolingListTabPage
        'Dim mydrv As DataRowView = APBS.Current

        If TextBox1.Text <> "" And TextBox12.Text <> "" And ComboBox1.Text <> "" Then 'And mydrv.Row.Item("projectcode", DataRowVersion.Original) = TextBox1.Text Then
            TabControl1.SelectedTab = ToolingListTabPage
            Dim myform As New FormImportToolingList(DS)

            If myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
                'Application.DoEvents()
                ShowTotalInfo("Tooling List")
            Else

                ProgressReport(2, String.Format("Error: {0}", myform.mymessage))
            End If

            'If Not IsNothing(ToolingListDTBS) Then

            '    ToolingListDTBS.Position = 0
            '    DataGridView1.DataSource = ToolingListDTBS
            'End If
            'ToolingListDTBS = New BindingSource
            'ToolingListDTBS.DataSource = DS.Tables(1)
            'DataGridView1.DataSource = ToolingListDTBS
            'CM1 = CType(BindingContext(ToolingInvoiceBS), CurrencyManager)
            'DataGridView1.Refresh()

            DataGridView1.Invalidate()
        ElseIf TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Project Code cannot be blank. ")
        ElseIf TextBox12.Text = "" Then
            TabControl1.SelectedTab = AssetsPurchaseTabPage
            ErrorProvider1.SetError(ComboBox5, "Supplier cannot be blank. ")
        ElseIf ComboBox1.Text = "" Then
            TabControl1.SelectedTab = AssetsPurchaseTabPage
            ErrorProvider1.SetError(ComboBox1, "Type Of Investment cannot be blank. ")
        End If



    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TabControl1.SelectedTab = AttachmentTabPage
    End Sub

    'Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    TabControl1.SelectedTab = PaymentHistoryTabPage
    'End Sub


    Private Sub DataGridView3_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellDoubleClick
        If Not IsNothing(ToolingInvoiceBS.Current) Then
            Dim drv As DataRowView = ToolingInvoiceBS.Current
            Dim myform As New FormInputInvoice(DS, ToolingInvoiceBS, Me)
            If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
                'add new record
                'DataGridView3.Invalidate()
                Dim apdrv As DataRowView = APBS.Current
                If drv.Row.RowState = DataRowState.Modified Or drv.Row.RowState = DataRowState.Added Then
                    If apdrv.Row.Item("otcexrate") <> drv.Row.Item("exrate") Then
                        apdrv.Row.Item("otcexrate") = drv.Row.Item("exrate")
                    End If
                End If
            Else
                'do nothing
                'ToolingInvoiceBS.RemoveCurrent()
            End If
            DataGridView3.Invalidate()
        End If
    End Sub

    Private Sub DeleteAttachmentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteAttachmentToolStripMenuItem.Click
        If Not IsNothing(AttachmentBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView2.SelectedRows
                    AttachmentBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim openfiledialog1 As New OpenFileDialog
        If openfiledialog1.ShowDialog = DialogResult.OK Then
            Dim mydrv As DataRowView = AttachmentBS.Current
            mydrv.Row.Item("docfullname") = IO.Path.GetFullPath(openfiledialog1.FileName)
            mydrv.Row.Item("docname") = IO.Path.GetFileName(openfiledialog1.FileName)
            mydrv.Row.Item("docext") = IO.Path.GetExtension(openfiledialog1.FileName)
            TextBox32.Text = IO.Path.GetFileName(openfiledialog1.FileName)
            DataGridView2.Invalidate()
        End If
    End Sub
    Private Sub NewRecordAttachmentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewRecordAttachmentToolStripMenuItem.Click

        Dim drv As DataRowView = AttachmentBS.AddNew
        drv.Item("docdate") = Today.Date
        drv.Item("assetpurchaseid") = DirectCast(APBS.Current, DataRowView).Row.Item("id")        
        drv.EndEdit() 'Change RowState status from Detached to Added, to allow add ToolingPayment        
        DataGridView2.Invalidate()
        TextBox32.Enabled = False
    End Sub
    Private Sub NewRecordInvoiceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewRecordInvoiceToolStripMenuItem.Click

        If ToolingListDTBS.Count > 0 Then
            Dim drv As DataRowView = ToolingInvoiceBS.AddNew()
            drv.Item("invoicedate") = Today.Date
            drv.Item("apid") = DirectCast(APBS.Current, DataRowView).Row.Item("id")
            drv.Item("vendorcode") = DirectCast(APBS.Current, DataRowView).Row.Item("vendorcode")
            drv.EndEdit() 'Change RowState status from Detached to Added, to allow add ToolingPayment

            Dim myform As New FormInputInvoice(DS, ToolingInvoiceBS, Me)
            If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
                'add new record
                'DataGridView3.Invalidate()
                Dim apdrv As DataRowView = APBS.Current
                If drv.Row.RowState = DataRowState.Modified Or drv.Row.RowState = DataRowState.Added Then
                    If apdrv.Row.Item("otcexrate") <> drv.Row.Item("exrate") Then
                        apdrv.Row.Item("otcexrate") = drv.Row.Item("exrate")
                    End If
                End If
            Else
                'do nothing
                'ToolingInvoiceBS.RemoveCurrent()
            End If
            'myform.Close()
            'myform.Dispose()
            'myform = Nothing
            DataGridView3.Invalidate()
            ShowTotalInfo("Payment & History")
        Else
            MessageBox.Show("Please upload tooling list first.")
        End If
        
    End Sub

    Private Sub DeleteInvoiceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteInvoiceToolStripMenuItem.Click
        If Not IsNothing(ToolingInvoiceBS.Current) Then
            'ToolingInvoiceBS.EndEdit()
            'ToolingPaymentbs.EndEdit()
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView3.SelectedRows
                    ToolingInvoiceBS.RemoveAt(drv.Index)
                Next
                ShowTotalInfo("Payment & History")
            End If
        End If
    End Sub


    Private Sub ClearDataBinding()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        TextBox6.DataBindings.Clear()

        'General Information
        DateTimePicker1.DataBindings.Clear()
        DateTimePicker2.DataBindings.Clear()
        TextBox10.DataBindings.Clear()
        'TextBox11.DataBindings.Clear()

        TextBox12.DataBindings.Clear()
        TextBox9.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()


        'TextBox8.DataBindings.Clear()

        ComboBox3.DataBindings.Clear()

        ComboBox5.DataBindings.Clear()
        ComboBox7.DataBindings.Clear()
        ComboBox6.DataBindings.Clear()
        ListBox1.DataBindings.Clear()
        Button5.DataBindings.Clear()

        ComboBox4.DataBindings.Clear()
        TextBox39.DataBindings.Clear()
        TextBox42.DataBindings.Clear()
        DateTimePicker3.DataBindings.Clear()

        'Related Information
        TextBox7.DataBindings.Clear()
        TextBox16.DataBindings.Clear()
        TextBox17.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        ComboBox9.DataBindings.Clear()
        TextBox19.DataBindings.Clear()
        TextBox14.DataBindings.Clear()
        TextBox13.DataBindings.Clear()
        TextBox43.DataBindings.Clear()

        'Amortization
        TextBox24.DataBindings.Clear()
        TextBox15.DataBindings.Clear()
        TextBox21.DataBindings.Clear()
        TextBox20.DataBindings.Clear()
        TextBox23.DataBindings.Clear()
        TextBox22.DataBindings.Clear()
        TextBox26.DataBindings.Clear()
        TextBox25.DataBindings.Clear()
        TextBox27.DataBindings.Clear()
        TextBox28.DataBindings.Clear()

        TextBox29.DataBindings.Clear()
        TextBox30.DataBindings.Clear()
        TextBox31.DataBindings.Clear()
        TextBox34.DataBindings.Clear()
        TextBox32.DataBindings.Clear()
        TextBox40.DataBindings.Clear()
        ComboBox8.DataBindings.Clear()
        TextBox35.DataBindings.Clear()
        TextBox36.DataBindings.Clear()
        TextBox37.DataBindings.Clear()
    End Sub

    Private Sub AddDataBinding()
        TextBox1.DataBindings.Add(New Binding("Text", APBS, "projectcode", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("Text", APBS, "projectname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("Text", APBS, "dept", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox4.DataBindings.Add(New Binding("Text", APBS, "ppps", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox5.DataBindings.Add(New Binding("Text", APBS, "familydescription", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox6.DataBindings.Add(New Binding("Text", APBS, "sbuname2", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        'General Information

        TextBox12.DataBindings.Add(New Binding("Text", APBS, "assetpurchaseid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        'TextBox11.DataBindings.Add(New Binding("Text", APBS, "applicantname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox7.DataSource = ApplicantBS
        ComboBox7.DisplayMember = "username"
        ComboBox7.ValueMember = "username"
        ComboBox7.DataBindings.Add("SelectedValue", APBS, "applicantname", True, DataSourceUpdateMode.OnPropertyChanged)

        DateTimePicker1.DataBindings.Add(New Binding("Text", APBS, "applicantdate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker2.DataBindings.Add(New Binding("text", APBS, "sapcapdate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox10.DataBindings.Add(New Binding("Text", APBS, "assetdescription", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox9.DataBindings.Add(New Binding("Text", APBS, "categoryofasset", True, DataSourceUpdateMode.OnPropertyChanged, ""))


        ComboBox5.DataSource = VendorBS
        ComboBox5.DisplayMember = "description"
        ComboBox5.ValueMember = "vendorcode"
        ComboBox5.DataBindings.Add("SelectedValue", APBS, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged)


        ComboBox1.DataSource = TypeOfInvestmentBS
        ComboBox1.DisplayMember = "typeofinvestmentname"
        ComboBox1.ValueMember = "typeofinvestmentid"
        ComboBox1.DataBindings.Add("SelectedValue", APBS, "typeofinvestment", True, DataSourceUpdateMode.OnPropertyChanged)


        ComboBox3.DataSource = DocumentTypeBS
        ComboBox3.DisplayMember = "documenttypedesc"
        ComboBox3.ValueMember = "documenttypeid"
        ComboBox3.DataBindings.Add("SelectedValue", AttachmentBS, "doctypeid", True, DataSourceUpdateMode.OnPropertyChanged)
        'ComboBox3.SelectedIndex = -1

        'TextBox8.DataBindings.Add(New Binding("Text", APBS, "approvalname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        'ComboBox4.DataSource = ApprovalBS
        'ComboBox4.DisplayMember = "username"
        'ComboBox4.ValueMember = "username"
        'ComboBox4.DataBindings.Add("SelectedValue", APBS, "approvalname", True, DataSourceUpdateMode.OnPropertyChanged)

        ComboBox6.DataSource = PaymentMethodBS
        ComboBox6.DisplayMember = "paymentmethoddesc"
        ComboBox6.ValueMember = "paymentmethodid"
        ComboBox6.DataBindings.Add("SelectedValue", APBS, "paymentmethodid", True, DataSourceUpdateMode.OnPropertyChanged)


        'listbox

        ListBox1.DisplayMember = "trackingno"
        ListBox1.ValueMember = "trackingno"
        ListBox1.DataSource = TrackingBS
        'ListBox1.DataBindings.Add("SelectedValue", TrackingBS, "trackingno", True, DataSourceUpdateMode.OnPropertyChanged)


        'Related Information
        TextBox17.DataBindings.Add(New Binding("Text", APBS, "toolingcost", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))
        TextBox40.DataBindings.Add(New Binding("Text", APBS, "aeb", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox16.DataBindings.Add(New Binding("Text", APBS, "budgetamount", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00"))


        ' ComboBox2.DisplayMember = "budgetcurr"
        ' ComboBox2.ValueMember = "budgetcurr"
        ComboBox2.DataBindings.Add(New Binding("SelectedItem", APBS, "budgetcurr", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox9.DataBindings.Add(New Binding("SelectedItem", APBS, "originaltoolingcostcurrency", True, DataSourceUpdateMode.OnPropertyChanged))

        TextBox19.DataBindings.Add(New Binding("Text", APBS, "exchangerate", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00##"))
        TextBox43.DataBindings.Add(New Binding("Text", APBS, "otcexrate", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00##"))
        TextBox14.DataBindings.Add(New Binding("Text", APBS, "investmentorderno", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox13.DataBindings.Add(New Binding("Text", APBS, "toolingpono", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox7.DataBindings.Add(New Binding("Text", APBS, "financeassetno", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        'Amortization
        ComboBox8.DataBindings.Add(New Binding("SelectedItem", APBS, "amortcurr", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox34.DataBindings.Add(New Binding("Text", APBS, "amortexrate", True, DataSourceUpdateMode.OnPropertyChanged, "", "#,##0.00##"))
        TextBox24.DataBindings.Add(New Binding("Text", APBS, "amortperiod_1", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox15.DataBindings.Add(New Binding("Text", APBS, "amortperiod_2", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox21.DataBindings.Add(New Binding("Text", APBS, "amortqty_1", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox20.DataBindings.Add(New Binding("Text", APBS, "amortqty_2", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox23.DataBindings.Add(New Binding("Text", APBS, "totalamortqty_1", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox22.DataBindings.Add(New Binding("Text", APBS, "totalamortqty_2", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox26.DataBindings.Add(New Binding("Text", APBS, "totalamortamount_1", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox25.DataBindings.Add(New Binding("Text", APBS, "totalamortamount_2", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox28.DataBindings.Add(New Binding("Text", APBS, "amortperunit_1", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox27.DataBindings.Add(New Binding("Text", APBS, "amortperunit_2", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        TextBox29.DataBindings.Add(New Binding("Text", APBS, "reason", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox30.DataBindings.Add(New Binding("Text", APBS, "remarks", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox31.DataBindings.Add(New Binding("Text", APBS, "comments", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox35.DataBindings.Add(New Binding("Text", APBS, "amortremarks", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox36.DataBindings.Add(New Binding("Text", APBS, "approvalname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox37.DataBindings.Add(New Binding("Text", APBS, "approvalname2", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        TextBox32.DataBindings.Add(New Binding("Text", AttachmentBS, "docname", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        ComboBox4.DataBindings.Add(New Binding("Text", APBS, "paymententity", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox39.DataBindings.Add(New Binding("Text", APBS, "toolingsupplierdescription", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker3.DataBindings.Add(New Binding("Text", APBS, "expectedtoolingfinishdate", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox42.DataBindings.Add(New Binding("Text", APBS, "origin", True, DataSourceUpdateMode.OnPropertyChanged, ""))
    End Sub
    Private Sub enabledComponents()
        TextBox1.Enabled = Not IsNothing(APBS.Current)
        TextBox2.Enabled = Not IsNothing(APBS.Current)
        'TextBox3.Enabled = Not IsNothing(APBS.Current)
        TextBox4.Enabled = Not IsNothing(APBS.Current)
        'TextBox5.Enabled = Not IsNothing(APBS.Current)
        'TextBox6.Enabled = Not IsNothing(APBS.Current)
        Button3.Enabled = Not IsNothing(APBS.Current)

        'General Information
        DateTimePicker1.Enabled = Not IsNothing(APBS.Current)
        TextBox10.Enabled = Not IsNothing(APBS.Current)
        ' TextBox11.Enabled = Not IsNothing(APBS.Current)
        'TextBox12.Enabled = Not IsNothing(APBS.Current)
        ComboBox7.Enabled = Not IsNothing(APBS.Current)
        TextBox9.Enabled = Not IsNothing(APBS.Current)
        ComboBox1.Enabled = Not IsNothing(APBS.Current)
        'TextBox8.Enabled = Not IsNothing(APBS.Current)

        ComboBox5.Enabled = Not IsNothing(APBS.Current)
        Button5.Enabled = Not IsNothing(APBS.Current)

        ComboBox4.Enabled = Not IsNothing(APBS.Current)
        TextBox39.Enabled = Not IsNothing(APBS.Current)
        TextBox42.Enabled = Not IsNothing(APBS.Current)
        DateTimePicker3.Enabled = Not IsNothing(APBS.Current)


        'Related Information
        TextBox7.Enabled = Not IsNothing(APBS.Current)
        TextBox16.Enabled = Not IsNothing(APBS.Current)
        TextBox17.Enabled = Not IsNothing(APBS.Current)
        ComboBox2.Enabled = Not IsNothing(APBS.Current)

        TextBox19.Enabled = Not IsNothing(APBS.Current)
        TextBox14.Enabled = Not IsNothing(APBS.Current)
        TextBox13.Enabled = Not IsNothing(APBS.Current)

        'Amortization
        ComboBox8.Enabled = Not IsNothing(APBS.Current)
        TextBox24.Enabled = Not IsNothing(APBS.Current)
        TextBox15.Enabled = Not IsNothing(APBS.Current)
        TextBox21.Enabled = Not IsNothing(APBS.Current)
        TextBox20.Enabled = Not IsNothing(APBS.Current)
        TextBox23.Enabled = Not IsNothing(APBS.Current)
        TextBox22.Enabled = Not IsNothing(APBS.Current)
        TextBox26.Enabled = Not IsNothing(APBS.Current)
        TextBox25.Enabled = Not IsNothing(APBS.Current)
        TextBox27.Enabled = Not IsNothing(APBS.Current)
        TextBox28.Enabled = Not IsNothing(APBS.Current)

        TextBox29.Enabled = Not IsNothing(APBS.Current)
        TextBox30.Enabled = Not IsNothing(APBS.Current)
        TextBox31.Enabled = Not IsNothing(APBS.Current)
        TextBox34.Enabled = Not IsNothing(APBS.Current)
        TextBox35.Enabled = Not IsNothing(APBS.Current)
        'TextBox36.Enabled = Not IsNothing(APBS.Current)

        TextBox18.Enabled = False
        TextBox33.Enabled = False
        TextBox11.Enabled = False
        TextBox8.Enabled = False
    End Sub
    Private Sub APBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles APBS.ListChanged
        Dim drv As DataRowView = APBS.Current
        'Project Information
        ToolStripButton5.Visible = False
        ToolStripButton6.Visible = False
        ToolStripButton7.Visible = False
        ToolStripButton8.Visible = False
        ToolStripButton9.Visible = False
        If Not FromApproval Then

            'If Not IsNothing(drv) Then
            '    If IsDBNull(drv.Item("status")) Then
            '        ToolStripButton5.Visible = True
            '    Else
            '        'ToolStripButton2.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusDraft
            '        ToolStripButton5.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusDraft
            '        ToolStripButton6.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusNew Or drv.Item("status") = AssetPurchaseStatusEnum.StatusReSubmit
            '        ToolStripButton7.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusNew Or drv.Item("status") = AssetPurchaseStatusEnum.StatusReSubmit
            '        ToolStripButton8.Visible = False
            '        If HelperClass1.UserInfo.IsAdmin Then ToolStripButton8.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByControlling
            '    End If

            'End If
            If Not IsNothing(drv) Then
                If IsDBNull(drv.Item("status")) Then
                    ToolStripButton5.Visible = True
                    'Else
                    '    'ToolStripButton2.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusDraft
                    '    ToolStripButton5.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusDraft
                    '    ToolStripButton6.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusNew Or drv.Item("status") = AssetPurchaseStatusEnum.StatusReSubmit
                    '    ToolStripButton7.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusNew Or drv.Item("status") = AssetPurchaseStatusEnum.StatusReSubmit
                    '    ToolStripButton8.Visible = False
                    '    If HelperClass1.UserInfo.IsAdmin Then ToolStripButton8.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByControlling
                Else
                    If (drv.Item("status") = AssetPurchaseStatusEnum.StatusDraft) Then ToolStripButton5.Visible = True
                    ToolStripButton9.Visible = HelperClass1.UserInfo.IsAdmin
                End If

            End If
            enabledComponents()

            
        Else
            'Check Department
            'Debug.Print("check department")
            If Not IsNothing(HelperClass1.UserInfo.GroupName) Then
                If HelperClass1.UserInfo.GroupName.Contains("Controlling Dept") Then
                    ToolStripButton6.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByPurchasing
                    ToolStripButton7.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByPurchasing
                ElseIf HelperClass1.UserInfo.GroupName.Contains("Accounting Dept") Then
                    ToolStripButton8.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByControlling
                End If
            Else
                If Not IsNothing(drv) Then
                    If Not FromApprovalHistory Then
                        If IsDBNull(drv.Item("status")) Then
                            ToolStripButton5.Visible = True
                        Else
                            'ToolStripButton2.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusDraft
                            ToolStripButton5.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusDraft Or drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByControlling
                            ToolStripButton6.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusNew Or drv.Item("status") = AssetPurchaseStatusEnum.StatusReSubmit Or drv.Item("status") = AssetPurchaseStatusEnum.StatusFirstValidatedByPurchasing
                            ToolStripButton7.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusNew Or drv.Item("status") = AssetPurchaseStatusEnum.StatusReSubmit Or drv.Item("status") = AssetPurchaseStatusEnum.StatusFirstValidatedByPurchasing
                            ToolStripButton8.Visible = False
                            ToolStripButton9.Visible = HelperClass1.UserInfo.IsAdmin
                            If HelperClass1.UserInfo.IsAdmin Then ToolStripButton8.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByControlling Or drv.Item("status") = AssetPurchaseStatusEnum.StatusPendingPayment
                        End If
                        enabledComponents()
                    End If

                End If
            End If
            'If HelperClass1.UserInfo.GroupName.Contains("Controlling Dept") Then
            '    ToolStripButton6.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByPurchasing
            '    ToolStripButton7.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByPurchasing
            'ElseIf HelperClass1.UserInfo.GroupName.Contains("Accounting Dept") Then
            '    ToolStripButton8.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByControlling
            'Else
            '    If Not IsNothing(drv) Then
            '        If Not FromApprovalHistory Then
            '            If IsDBNull(drv.Item("status")) Then
            '                ToolStripButton5.Visible = True
            '            Else
            '                'ToolStripButton2.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusDraft
            '                ToolStripButton5.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusDraft
            '                ToolStripButton6.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusNew Or drv.Item("status") = AssetPurchaseStatusEnum.StatusReSubmit
            '                ToolStripButton7.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusNew Or drv.Item("status") = AssetPurchaseStatusEnum.StatusReSubmit
            '                ToolStripButton8.Visible = False
            '                If HelperClass1.UserInfo.IsAdmin Then ToolStripButton8.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByControlling
            '            End If
            '            enabledComponents()
            '        End If

            '    End If
            'End If

                'Select Case HelperClass1.UserInfo.GroupName
                '    Case "Controlling Dept"
                '        ToolStripButton6.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByPurchasing
                '        ToolStripButton7.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByPurchasing

                '    Case "Accounting Dept"
                '        ToolStripButton8.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByControlling
                '    Case Else
                '        If Not IsNothing(drv) Then
                '            If Not FromApprovalHistory Then
                '                If IsDBNull(drv.Item("status")) Then
                '                    ToolStripButton5.Visible = True
                '                Else
                '                    'ToolStripButton2.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusDraft
                '                    ToolStripButton5.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing Or drv.Item("status") = AssetPurchaseStatusEnum.StatusDraft
                '                    ToolStripButton6.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusNew Or drv.Item("status") = AssetPurchaseStatusEnum.StatusReSubmit
                '                    ToolStripButton7.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusNew Or drv.Item("status") = AssetPurchaseStatusEnum.StatusReSubmit
                '                    ToolStripButton8.Visible = False
                '                    If HelperClass1.UserInfo.IsAdmin Then ToolStripButton8.Visible = drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByControlling
                '                End If
                '                enabledComponents()
                '            End If

                '        End If

                'End Select
            End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click, Button3.Click, Button4.Click, Button8.Click, Button9.Click
        Dim myobj As Button = CType(sender, Button)
        Try
            Select Case myobj.Name
                Case "Button5"
                    Dim myform = New FormHelper(VendorHelperBS)
                    myform.DataGridView1.Columns(0).DataPropertyName = "description"
                    If myform.ShowDialog = DialogResult.OK Then
                        Dim drv As DataRowView = VendorHelperBS.Current
                        Dim mydrv As DataRowView = APBS.Current
                        mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
                        mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
                        mydrv.Row.EndEdit()
                        'Sync with Combobox BindingSource
                        Dim itemfound As Integer = VendorBS.Find("vendorcode", drv.Row.Item("vendorcode"))
                        VendorBS.Position = itemfound
                        displayassetpurchaseid()

                    End If
                Case "Button3"
                    Dim myform = New FormHelper(ToolingProjectHelperBS)
                    myform.DataGridView1.Columns(0).DataPropertyName = "projectdescription"
                    If myform.ShowDialog = DialogResult.OK Then
                        If Not IsNothing(ToolingProjectHelperBS.Current) Then
                            Dim drv As DataRowView = ToolingProjectHelperBS.Current
                            Dim mydrv As DataRowView = APBS.Current

                            TextBox3.Text = "" + drv.Row.Item("dept")
                            TextBox4.Text = "" + drv.Row.Item("ppps")
                            TextBox5.Text = "" + drv.Row.Item("familydescription")
                            TextBox6.Text = "" + drv.Row.Item("sbuname2")
                            TextBox1.Text = drv.Row.Item("projectcode")
                            TextBox2.Text = "" + drv.Row.Item("projectname") 'I put it at the botton on purpossed
                            mydrv.Row.Item("familyid") = drv.Row.Item("familyid")

                            mydrv.Row.Item("projectid") = drv.Row.Item("id")
                            mydrv.Row.EndEdit()

                        End If
                        ErrorProvider1.SetError(TextBox1, "")
                    End If
                Case "Button4"
                    Dim myform = New FormHelper(FamilyGroupHelperBS)
                    myform.DataGridView1.Columns(0).DataPropertyName = "description"
                    If myform.ShowDialog = DialogResult.OK Then
                        Dim drv As DataRowView = FamilyGroupHelperBS.Current
                        Dim mydrv As DataRowView = APBS.Current

                        Dim projecthelperdrv As DataRowView = ToolingProjectHelperBS.Current

                        mydrv.Row.Item("familyid") = drv.Row.Item("familyid")

                        TextBox5.Text = "" + drv.Row.Item("description")
                        TextBox6.Text = "" + drv.Row.Item("sbuname2")

                        projecthelperdrv.Row.Item("familydescription") = drv.Row.Item("description")
                        projecthelperdrv.Row.Item("sbuname2") = drv.Row.Item("sbuname2")
                    End If
                Case "Button8"
                    Dim myform = New FormHelper(ApprovalBS)
                    myform.DataGridView1.Columns(0).DataPropertyName = "username"
                    If myform.ShowDialog = DialogResult.OK Then
                        Dim drv As DataRowView = ApprovalBS.Current
                        Dim mydrv As DataRowView = APBS.Current


                        mydrv.Row.Item("approvalname") = drv.Row.Item("username")

                        TextBox36.Text = "" + drv.Row.Item("username")


                    End If
                Case "Button9"
                    Dim myform = New FormHelper(ApprovalBS2)
                    myform.DataGridView1.Columns(0).DataPropertyName = "username"
                    If myform.ShowDialog = DialogResult.OK Then
                        Dim drv As DataRowView = ApprovalBS2.Current
                        Dim mydrv As DataRowView = APBS.Current


                        mydrv.Row.Item("approvalname2") = drv.Row.Item("username")

                        TextBox37.Text = "" + drv.Row.Item("username")


                    End If
            End Select
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
        DataGridView1.Invalidate()
    End Sub

    Private Sub MyTextChanged(ByVal sender As Object, ByVal e As EventArgs)
        If TextBox1.Text.Trim <> "" Then
            If sender.GetType Is GetType(TextBox) Then
                Select Case DirectCast(sender, TextBox).Name
                    Case "TextBox1", "TextBox2", "TextBox5"
                        ErrorProvider1.SetError(sender, "")
                End Select
            End If

            'APBS.EndEdit() 'This method is not correct, you should get the current and do endedit as below
            If Not IsNothing(ToolingProjectHelperBS.Current) Then
                Dim TPdrv = ToolingProjectHelperBS.Current
                TPdrv.EndEdit()
            End If

            If Not IsNothing(APBS.Current) Then
                Dim drv As DataRowView = APBS.Current
                drv.EndEdit()
            End If
        End If

    End Sub


    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted, ComboBox5.SelectionChangeCommitted, ComboBox2.SelectionChangeCommitted, ComboBox3.SelectionChangeCommitted
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)
        '1. Force the Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        Dim drv As DataRowView = APBS.Current
        drv.EndEdit()
        If myobj.Name = "ComboBox5" Then
            ErrorProvider1.SetError(ComboBox5, "")
            displayassetpurchaseid()
        ElseIf myobj.Name = "ComboBox1" Then
            ErrorProvider1.SetError(ComboBox1, "")
        ElseIf myobj.Name = "ComboBox3" Then
            Dim cb3selected As DocumentType = ComboBox3.SelectedItem
            Dim drvr As DataRowView = AttachmentBS.Current
            drvr.Item("doctypename") = cb3selected.DocumentTypeDesc
            'drvr.EndEdit()
            DataGridView2.Invalidate()
        End If
    End Sub

 
    Private Sub ToolStripButton2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Try
            If Me.validate Then
                Try
                    'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                    Dim mydrv As DataRowView = APBS.Current
                    If Not DateTimePicker2.Checked Then

                        mydrv.Row.Item("sapcapdate") = DBNull.Value
                    End If

                    'Check Status if Pending Payment then Change it to Compeleted -- This modification moved to button Complete
                    'If mydrv.Row.Item("status") = AssetPurchaseStatusEnum.StatusPendingPayment Then
                    '    mydrv.Row.Item("status") = AssetPurchaseStatusEnum.StatusCompleted
                    'End If


                    APBS.EndEdit() 'This is important, use this for merge
                    ToolingProjectHelperBS.EndEdit()
                    ToolingListDTBS.EndEdit()
                    AttachmentBS.EndEdit()
                    TrackingBS.EndEdit()
                    'Delete Blank ToolingProjectHelperBS
                    For Each myrow As DataRowView In ToolingProjectHelperBS.List
                        If IsDBNull(myrow.Item("projectname")) Then
                            myrow.Delete()
                        ElseIf myrow.Item("projectname") = "" Then
                            myrow.Delete()

                        End If
                    Next

                    Dim ds2 As DataSet
                    ds2 = DS.GetChanges

                    If Not IsNothing(ds2) Then

                        Dim mymessage As String = String.Empty
                        Dim ra As Integer
                        Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)

                        'delete file first

                        For Each drv As DataRowView In AttachmentBS.List
                            'For Each drv As DataRow In ds2.Tables("attachment").Rows
                            Try
                                If Not drv.Item("docfullname") = "" Then
                                    Try
                                        If drv.Row.Item("docname", DataRowVersion.Original) <> "" Then
                                            Dim orifile = DS.Tables(11).Rows(0).Item("cvalue") & "\" & drv.Item("id") & drv.Row.Item("docext", DataRowVersion.Original)
                                            If FileIO.FileSystem.FileExists(orifile) Then
                                                FileIO.FileSystem.DeleteFile(orifile)
                                            End If

                                        End If
                                    Catch ex As Exception
                                        Logger.log(String.Format("** Delete Original File error {0}**", ex.Message))
                                    End Try
                                End If
                            Catch ex As Exception

                            End Try

                        Next


                        If Not DbAdapter1.AssetPurchaseTx(Me, mye) Then
                            MessageBox.Show(mye.message)
                            DS.Merge(ds2)
                            Exit Sub
                        End If
                        'DS.AcceptChanges()

                        DS.Merge(ds2)
                        Try
                            DS.AcceptChanges()  'AccetpChanges complaint because of invoiceid in ToolingPayment still old data (-1)
                            'Update:: After removed constraint, no more issue.
                        Catch ex As Exception
                            'After got error ,Invoice id in ToolingPayment changed into new one. strange!

                            'DS.AcceptChanges()
                        End Try

                        ' If Not ds2.Tables(1).HasErrors Then
                        If Not ds2.HasErrors Then
                            For Each dr As DataRow In ds2.Tables("ToolingProject").Rows
                                If dr.RowState = DataRowState.Modified Then
                                    If Not ToolingProjectDict.ContainsKey(dr.Item("projectcode")) Then
                                        ToolingProjectDict.Add(dr.Item("projectcode"), dr.Item("id", DataRowVersion.Current))
                                    End If
                                End If
                            Next
                            ProgressReport(2, "Copying File...")
                            'Copy File
                            For Each drv As DataRowView In AttachmentBS.List
                                If Not drv.Row.Item("docfullname") = "" Then
                                    Dim mytarget = DS.Tables(11).Rows(0).Item("cvalue") & "\" & drv.Row.Item("id") & drv.Row.Item("docext")
                                    Try
                                        FileIO.FileSystem.CopyFile(drv.Row.Item("docfullname"), mytarget, True)
                                    Catch ex As Exception
                                        Logger.log(String.Format("** CopyFile error {0}**", ex.Message))
                                    End Try
                                End If
                            Next
                            ProgressReport(2, "")
                            MessageBox.Show("Saved.")

                            myId = DS.Tables(0).Rows(0).Item("id")
                            Dim toolingprojectid = DS.Tables(0).Rows(0).Item("projectid")
                            Loadata()
                        Else
                            ProgressReport(2, "Error occurs. Please check.")
                        End If

                    End If
                Catch ex As Exception
                    MessageBox.Show(" Error:: " & ex.Message)
                End Try
            Else
                ProgressReport(2, "Error occurs. Please check.")
            End If
            'DataGridView1.Invalidate()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Overloads Function validate() As Boolean        
        Dim myret As Boolean = True
        MyBase.Validate()

        For Each drv As DataRowView In APBS.List
            If Not ValidateRecordAPBS(drv) Then
                myret = False
            End If
        Next

        For Each mydrv As DataRowView In AttachmentBS.List
            If Not validateRecordAttachment(mydrv) Then                
                myret = False
            End If
        Next
        Return myret
    End Function
    Private Function validateRecordAttachment(ByVal mydrv As DataRowView) As Boolean        
        Dim myret As Boolean = True
        Dim myerror As New StringBuilder
        Try
            If IsDBNull(mydrv.Row.Item("docname")) Then
                myerror.Append("File name cannot be blank.")
                myret = False
            End If

            If IsDBNull(mydrv.Row.Item("doctypeid")) Then
                If myerror.Length > 0 Then
                    myerror.Append(",")
                End If
                myerror.Append("Document Type cannot be blank.")
                myret = False
            End If
            mydrv.Row.RowError = myerror.ToString
        Catch ex As Exception
            myret = False
        End Try
        Return myret
    End Function
    Private Function ValidateRecordAPBS(ByVal drv As DataRowView) As Boolean
        Dim myerror As New StringBuilder
        Dim myret As Boolean = True
        Try
            Dim TPdrv As DataRowView = ToolingProjectHelperBS.Current
            Dim APdrv As DataRowView = APBS.Current
            If IsDBNull(drv.Row.Item("projectcode")) Then
                ErrorProvider1.SetError(TextBox1, "ProjectCode cannot be blank.")
                myret = False
            Else
                ErrorProvider1.SetError(TextBox1, "")
            End If
            If IsDBNull(drv.Row.Item("projectname")) Then
                ErrorProvider1.SetError(TextBox2, "Project Name cannot be blank.")
                myret = False
            Else
                ErrorProvider1.SetError(TextBox2, "")
                TPdrv.Row.Item("projectname") = APdrv.Row.Item("projectname")
            End If
            If IsDBNull(drv.Row.Item("familydescription")) Then
                ErrorProvider1.SetError(TextBox5, "Family cannot be blank.")
                myret = False
            Else
                ErrorProvider1.SetError(TextBox5, "")
                TPdrv.Row.Item("familyid") = APdrv.Row.Item("familyid")
                TPdrv.EndEdit()
            End If
            If IsDBNull(drv.Row.Item("vendorcode")) Then
                ErrorProvider1.SetError(ComboBox5, "Vendorcode cannot be blank.")
                myret = False
            Else
                ErrorProvider1.SetError(ComboBox5, "")                
            End If
            If Not IsNothing(TPdrv) Then
                TPdrv.Row.Item("PPPS") = APdrv.Row.Item("PPPS")
                TPdrv.Row.Item("dept") = APdrv.Row.Item("dept")
            End If
            If IsDBNull(drv.Row.Item("aeb")) Then
                ErrorProvider1.SetError(TextBox40, "AEB cannot be blank.")
                myret = False
            Else
                ErrorProvider1.SetError(TextBox40, "")
            End If
            If IsDBNull(drv.Row.Item("budgetamount")) Then
                ErrorProvider1.SetError(TextBox16, "Budget Amount cannot be blank.")
                myret = False
            Else
                ErrorProvider1.SetError(TextBox16, "")
            End If
            If IsDBNull(drv.Row.Item("reason")) Then
                ErrorProvider1.SetError(TextBox29, "Reason cannot be blank.")
                myret = False
            Else
                ErrorProvider1.SetError(TextBox29, "")
            End If
            If IsDBNull(drv.Row.Item("applicantname")) Then
                ErrorProvider1.SetError(ComboBox7, "Applicant Name cannot be blank.")
                myret = False
            Else
                ErrorProvider1.SetError(ComboBox7, "")
            End If
            If IsDBNull(drv.Row.Item("approvalname")) Then
                'ErrorProvider1.SetError(ComboBox4, "Approval Name cannot be blank.")
                ErrorProvider1.SetError(TextBox36, "Approval Name cannot be blank.")
                myret = False
            Else
                'ErrorProvider1.SetError(ComboBox4, "")
                ErrorProvider1.SetError(TextBox36, "")
            End If
            If IsDBNull(drv.Row.Item("typeofinvestment")) Then
                ErrorProvider1.SetError(ComboBox1, "Type of investment cannot be blank.")
                myret = False
            Else
                ErrorProvider1.SetError(ComboBox1, "")
            End If
            If Not IsDBNull(drv.Row.Item("approvalname2")) Then
                If drv.Row.Item("approvalname2") = "" Then
                    drv.Row.Item("approvalname2") = System.DBNull.Value
                End If
            End If
            

        Catch ex As Exception
            myret = False
            MessageBox.Show("Validate::" & ex.Message)
        End Try
        Return myret

    End Function

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        displayassetpurchaseid()
    End Sub


    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated
        If Not HelperClass1.UserInfo.IsFinance Then


            'Find Project in dictionary, if avail find it from toolingproject then show it , if not create it in toolingproject
            ErrorProvider1.SetError(sender, "")
            'If ToolingProjectDict.Count > 0 Then
            Dim drv1 As DataRowView = APBS.Current
            'Dim originalvalue = ""
            'Try
            '    originalvalue = drv1.Row.Item("projectcode", DataRowVersion.Original)
            'Catch ex As Exception

            'End Try
            TextBox1.Text = TextBox1.Text.Trim
            If TextBox1.Text <> "" Then 'And TextBox1.Text <> originalvalue Then
                If ToolingProjectDict.ContainsKey(TextBox1.Text) Then

                    Dim mydrv As DataRowView = APBS.Current
                    Dim pkey(0) As Object
                    pkey(0) = ToolingProjectDict(TextBox1.Text)
                    Dim dr = DS.Tables(4).Rows.Find(pkey)

                    If Not IsNothing(dr) Then
                        TextBox2.Text = "" & dr.Item("projectname")
                        TextBox3.Text = "" & dr.Item("dept")
                        TextBox4.Text = "" & dr.Item("ppps")
                        TextBox5.Text = "" & dr.Item("familydescription")
                        TextBox6.Text = "" + dr.Item("sbuname2")
                        mydrv.Row.Item("familyid") = dr.Item("familyid")
                        mydrv.Row.Item("projectid") = dr.Item("id")
                        'move toolingprojecthelper position
                        Dim pos = ToolingProjectHelperBS.Find("projectcode", TextBox1.Text)
                        ToolingProjectHelperBS.Position = pos
                    End If
                    'displayassetpurchaseid()
                    'Exit Sub
                Else
                    'No need to cancel 
                    Dim TPdrv As DataRowView = ToolingProjectHelperBS.AddNew() 'TPdrv rowstate is detached                
                    Dim drv As DataRowView = APBS.Current
                    TPdrv.Row.Item("projectcode") = drv.Row.Item("projectcode")
                    TPdrv.Row.Item("projectdescription") = ""
                    'Change TPdrv rowstate into attached in order to make relation with table AssetPurchase
                    TPdrv.EndEdit()
                    drv.Row.Item("projectid") = TPdrv.Row.Item("id")
                    drv.EndEdit()
                End If
            ElseIf TextBox1.Text = "" Then
                'MessageBox.Show("Project Code cannot be blank.")
                ErrorProvider1.SetError(TextBox1, "Project Code cannot be blank.")
                Application.DoEvents()
                TabControl1.SelectedTab = AssetsPurchaseTabPage
                TextBox1.Focus()
            End If
            displayassetpurchaseid()
        End If
    End Sub

    Private Sub FillDictionary()
        ToolingProjectDict = New Dictionary(Of String, Long)
        For Each dr As DataRow In DS.Tables("ToolingProject").Rows

            ToolingProjectDict.Add(dr.Item("projectcode"), dr.Item("id"))
        Next
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        Dim myapbs As DataRowView = APBS.Current
        Dim drv As DataRowView = ToolingProjectHelperBS.Current

        drv.Row.Item("projectdescription") = drv.Row.Item("projectcode") & " - " & TextBox2.Text.Trim 'drv.Row.Item("projectcode") & " - " & myapbs.Row.Item("projectname")
        drv.Row.Item("projectname") = TextBox2.Text 'myapbs.Row.Item("projectname")
        drv.Row.Item("dept") = myapbs.Row.Item("dept")
        drv.Row.Item("ppps") = myapbs.Row.Item("ppps")
    End Sub


    Private Sub initData()

        TypeOfInvestmentList = New List(Of TypeOfInvestment)
        TypeOfInvestmentList.Add(New TypeOfInvestment(1, "New Asset"))
        TypeOfInvestmentList.Add(New TypeOfInvestment(2, "Tooling Modification"))
        TypeOfInvestmentList.Add(New TypeOfInvestment(3, "Back Up"))
        TypeOfInvestmentList.Add(New TypeOfInvestment(4, "Increase Capacity"))

        DocumentTypeList = New List(Of DocumentType)
        DocumentTypeList.Add(New DocumentType(1, "AEB"))
        DocumentTypeList.Add(New DocumentType(2, "Invoices"))
        DocumentTypeList.Add(New DocumentType(3, "Supporting email(technical confirmation)"))
        DocumentTypeList.Add(New DocumentType(4, "Assets Purchase form with Approval Signature"))
        DocumentTypeList.Add(New DocumentType(5, "Tooling List"))
        DocumentTypeList.Add(New DocumentType(6, "Project Specification"))

        PaymentMethodList = New List(Of PaymentMethod)
        PaymentMethodList.Add(New PaymentMethod(1, "Amortization"))
        PaymentMethodList.Add(New PaymentMethod(2, "Invoice Investment"))


        'Get textbox,combobox,checkbox then put handler in one function
        TabControl1.SelectedTab = AssetsPurchaseTabPage
        For Each ctl As Control In TabControl1.SelectedTab.Controls
            Try
                If ctl.GetType Is GetType(GroupBox) Then
                    For Each ctl2 As Control In DirectCast(ctl, GroupBox).Controls
                        If ctl2.GetType Is GetType(TextBox) Then
                            AddHandler DirectCast(ctl2, TextBox).Validated, AddressOf MyTextChanged
                        ElseIf ctl2.GetType Is GetType(ComboBox) Then
                            If DirectCast(ctl2, ComboBox).Name = "ComboBox2" Then
                                'AddHandler DirectCast(ctl2, ComboBox).SelectionChangeCommitted, AddressOf MyTextChanged
                            End If
                            'No Need, just Force bindingsource to accept the value in SelectionChangeCommitted
                        ElseIf ctl2.GetType Is GetType(DateTimePicker) Then
                            'AddHandler DirectCast(ctl2, DateTimePicker).Validated, AddressOf MyTextChanged
                        End If
                    Next
                End If
                'If ctl.GetType Is GetType(TextBox) Then
                '    AddHandler DirectCast(ctl, TextBox).TextChanged, AddressOf MyTextChanged
                'ElseIf ctl.GetType Is GetType(ComboBox) Then
                '    AddHandler DirectCast(ctl, ComboBox).SelectedValueChanged, AddressOf MyTextChanged
                'ElseIf ctl.GetType Is GetType(DateTimePicker) Then
                '    AddHandler DirectCast(ctl, DateTimePicker).ValueChanged, AddressOf MyTextChanged
                'End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Next
    End Sub

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        If Not IsNothing(ToolingListDTBS) Then

            ShowTotalInfo(TabControl1.SelectedTab.Text)

            'If TabControl1.SelectedTab.Text = "Tooling List" Then
            '    ShowTotalInfo(TabControl1.SelectedTab.Text)
            '    ' ProgressReport(2, String.Format("Record Count : {0}", DataGridView1.RowCount))
            'ElseIf TabControl1.SelectedTab.Text = "Payment & History" Then
            '    shotTotalInvoice()
            'Else

            '    ProgressReport(2, "")
            'End If

        End If
    End Sub

    Private Sub ToolingListTabPage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolingListTabPage.Click

    End Sub

    'Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
    '    If TextBox1.Text = "" Then
    '        'MessageBox.Show("Project Code cannot be blank.")
    '        TabControl1.SelectedTab = AssetsPurchaseTabPage
    '    End If
    'End Sub


    Private Sub AssetsPurchaseTabPage_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles AssetsPurchaseTabPage.Validating
        If TextBox1.Text = "" Then
            e.Cancel = True
        End If
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click

    End Sub

    Private Sub displayassetpurchaseid()
        TextBox12.Text = ""
        If ComboBox5.SelectedIndex > 1 And TextBox1.Text <> "" Then
            TextBox12.Text = String.Format("{0}_{1}_{2:yyyyMMdd}", ComboBox5.SelectedValue, TextBox1.Text, DateTimePicker1.Value.Date)
        End If
    End Sub

    Private Sub DateTimePicker1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker1.Validated
        displayassetpurchaseid()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If Not IsNothing(ToolingListDTBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    'Try
                    If drv.Cells.Item("Cost").Value <> drv.Cells.Item("Balance").Value Then
                        If MessageBox.Show(String.Format("This Tooling {0} has payment. Delete this record?", drv.Cells.Item(0).Value), "Has Payment", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                            ToolingListDTBS.RemoveAt(drv.Index)
                        End If
                    Else
                        Try
                            ToolingListDTBS.RemoveAt(drv.Index)
                        Catch ex As Exception

                        End Try

                    End If
                    'Catch ex As Exception
                    '    ToolingListDTBS.RemoveAt(drv.Index)
                    'End Try

                Next
            End If
            DataGridView1.Invalidate()

            ShowTotalInfo("Tooling List")
        End If
    End Sub

    Private Sub BatchPaymentToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Validate TextBox

        'Validate ToolingList

        'Message Asking for creating Detail

        'Create based on Tooling List
    End Sub

    Private Sub TextBox16_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox16.TextChanged, TextBox19.TextChanged
        Try
            TextBox41.Text = String.Format("{0:#,##0.00}", TextBox16.Text * TextBox19.Text)
        Catch ex As Exception

        End Try

    End Sub


    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim sqlstr As String = String.Format("select tl.id,tl.assetpurchaseid,tm.sebmodelno,tm.suppliermodelreference,tm.suppliermoldno,tm.toolsdescription,tm.material,tm.cavities,tm.numberoftools,tm.dailycapacity,tm.cost,tm.purchasedate ,tm.location,tm.comments ,tl.lineno,tl.vendorcode,tl.toolinglistid,tm.dailycaps,tm.commontool ,0 as typeofinvestment from doc.toolinglist tl left join doc.assetpurchase ap on ap.id = tl.assetpurchaseid  " &
                                " left join doc.toolingproject tp on tp.id = ap.projectid left join doc.toolinglistmst tm on tm.toolinglistid = tl.toolinglistid where tp.projectcode = '{0}' order by lineno", TextBox1.Text)

        Dim myform = New FormViewToolingList(sqlstr)
        myform.ShowDialog()


    End Sub

    Private Sub ShowTotalInfo(ByVal TabName As String)
        Select Case TabName
            Case "Tooling List"
                Dim TotalCost As Decimal = 0
                Dim TotalBalance As Decimal = 0
                Dim TotalCostCNY As Decimal = 0
                Dim TotalBalanceCNY As Decimal = 0
                Try
                    If Not IsNothing(ToolingListDTBS) Then
                        For Each drv As DataRowView In ToolingListDTBS.List
                            TotalCost = TotalCost + drv.Item("cost")
                            TotalBalance = TotalBalance + drv.Item("balance")
                            If drv.Row.Item("originalcurrency") = "CNY" Then
                                TotalCostCNY = TotalCostCNY + drv.Item("originalcost")
                            End If
                            TotalBalanceCNY = TotalBalanceCNY + drv.Item("balancecny")
                        Next
                    End If
                Catch ex As Exception

                End Try
                ProgressReport(2, String.Format("Record Count : {0}, Total Tooling cost (USD): {1:#,##0.00}, Balance (USD) : {2:#,##0.00}, Total cost (CNY) : {3:#,##0.00}, Balance (CNY) : {4:#,##0.00}", DataGridView1.RowCount, TotalCost, TotalBalance, TotalCostCNY, TotalBalanceCNY))
            Case "Payment & History"
                Dim TotalInvoice = 0
                Dim TotalInvoiceCNY = 0
                Try
                    If Not IsNothing(ToolingInvoiceBS) Then
                        For Each drv As DataRowView In ToolingInvoiceBS.List
                            TotalInvoice = TotalInvoice + drv.Item("totalamount")
                            TotalInvoiceCNY = TotalInvoiceCNY + drv.Item("totalamountcny")
                        Next
                    End If
                Catch ex As Exception

                End Try
                ProgressReport(2, String.Format("Record Count : {0}, Total Invoice (in USD): {1:#,##0.00}, Total Invoice (CNY) = {2:#,##0.00}", DataGridView3.RowCount, TotalInvoice, TotalInvoiceCNY))
            Case "Proforma Purchase Order"
                Dim TotalCost As Decimal = 0
                Dim TotalBalance As Decimal = 0
                Try
                    If Not IsNothing(ToolingListDTBS) Then
                        For Each drv As DataRowView In ToolingListDTBS.List
                            TotalCost = TotalCost + drv.Item("originalcost")
                            TotalBalance = TotalBalance + drv.Item("balance")
                            OriginalCurrency = "" & drv.Item("originalcurrency")
                        Next
                        OriginalCost = TotalCost

                    End If
                Catch ex As Exception

                End Try
                ProgressReport(2, String.Format("Proforma Invoice, Total Cost: {0:#,##0.00}, Original Currency = {1}", TotalCost, OriginalCurrency))
            Case Else
                ProgressReport(2, "")
        End Select

    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Dim myform As New FormPrintAssetsPurchase(DS)
        myform.ShowDialog()
    End Sub

   

    Private Sub AttachmentBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles AttachmentBS.ListChanged
        CheckBox1.Enabled = Not IsNothing(AttachmentBS.Current)
        'ComboBox3.Enabled = (Not IsNothing(AttachmentBS.Current)) And Not (FromApproval)
        ComboBox3.Enabled = (Not IsNothing(AttachmentBS.Current)) And Not (HelperClass1.UserInfo.IsFinance) And Not (FromApprovalHistory)
        Button1.Enabled = Not IsNothing(AttachmentBS.Current)
        Button2.Enabled = Not IsNothing(AttachmentBS.Current)

    End Sub




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.validate()
        If Not myThreadDownload.IsAlive Then
            'get Folder

            Dim myfolder = New SaveFileDialog
            myfolder.FileName = "SaveFile"
            If myfolder.ShowDialog = Windows.Forms.DialogResult.OK Then
                ToolStripStatusLabel1.Text = ""
                SelectedFolder = IO.Path.GetDirectoryName(myfolder.FileName)
                myThreadDownload = New Thread(AddressOf DoDownload)
                myThreadDownload.Start()
            End If

        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
    Sub DoDownload()
        Dim i As Integer = 0
        For Each drv As DataRowView In AttachmentBS.List
            If drv.Row.Item("download") Then
                i = i + 1
                ProgressReport(1, String.Format("Downloading ::{0} of {1} {2}", i, AttachmentBS.Count, drv.Row.Item("docname")))
                Dim filesource As String = String.Format("{0}\{1}{2}", DS.Tables(11).Rows(0).Item("cvalue"), drv.Row.Item("id"), drv.Row.Item("docext"))
                Dim FileTarget As String = String.Format("{0}\{1}", SelectedFolder, drv.Row.Item("docname"))
                Try
                    FileIO.FileSystem.CopyFile(filesource, FileTarget, True)
                Catch ex As Exception

                End Try

            End If
        Next
        ProgressReport(1, "Done. Please check your folder ::" & SelectedFolder)
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If Not IsNothing(AttachmentBS) Then
            For Each drv As DataRowView In AttachmentBS.List
                drv.Row.Item("download") = CheckBox1.Checked
            Next
            AttachmentBS.EndEdit()
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim myfolder = New SaveFileDialog
        myfolder.FileName = "ToolingListTemplate"
        myfolder.DefaultExt = ".xlsx"
        If myfolder.ShowDialog = Windows.Forms.DialogResult.OK Then
            ToolStripStatusLabel1.Text = ""
            SelectedFolder = IO.Path.GetDirectoryName(myfolder.FileName)

            'Dim filesource As String = String.Format("\\172.22.10.44\SharedFolder\PriceCmmf\New\template\ToolingListTemplate.xlsx")
            Dim filesource As String = String.Format("{0}\ToolingListTemplate.xlsx", HelperClass1.template)
            Dim FileTarget As String = String.Format("{0}\{1}", SelectedFolder, "ToolingListTemplate.xlsx")
            Try
                FileIO.FileSystem.CopyFile(filesource, FileTarget, True)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub



    Private Sub TextBox40_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox40.Validated, TextBox2.Validated
        Dim obj As TextBox = DirectCast(sender, TextBox)
        obj.Text = obj.Text.Trim
    End Sub

    Private Sub ToolStripButton3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Dim drv As DataRowView = APBS.Current
        If Not IsNothing(drv) Then
            If drv.Item("typeofinvestment") <> TypeOfInvestmentEnum.NewAsset Then
                MessageBox.Show("Only 'New Asset' type can use this function.")
                Exit Sub
            End If
            If MessageBox.Show("Do you want to update to tooling id?", "Update Tooling Id", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Dim sqlstr = String.Format("insert into doc.toolingidhistory(toolingid,changedby,changeddate,latesttoolingid)" &
                    " select tl.toolinglistid,'{0}',now()::date,a.assetpurchaseid || to_char(tl.lineno,'FM_0009') as newtoolinglistid from doc.toolinglist tl" &
                    " left join doc.toolinglistmst tm on tm.toolinglistid = tl.toolinglistid" &
                    " left join doc.assetpurchase a on a.id = tl.assetpurchaseid" &
                    " left join doc.toolingproject tp on tp.id = a.projectid" &
                    " where tl.assetpurchaseid = {1} and not tm.commontool and substring(tl.toolinglistid,1,1) <> 'S';" &
                    " insert into doc.toolingidhistory(toolingid,changedby,changeddate,latesttoolingid)" &
                    " select tl.toolinglistid,'{0}',now()::date,'S' || a.assetpurchaseid || to_char(tl.lineno,'FM_0009') as newtoolinglistid from doc.toolinglist tl" &
                    " left join doc.toolinglistmst tm on tm.toolinglistid = tl.toolinglistid" &
                    " left join doc.assetpurchase a on a.id = tl.assetpurchaseid" &
                    " left join doc.toolingproject tp on tp.id = a.projectid" &
                    " where tl.assetpurchaseid = {1} and not tm.commontool and substring(tl.toolinglistid,1,1) = 'S';" &
                    " update doc.toolinglistmst m set toolinglistid = foo.newtoolinglistid" &
                    " from (select tl.toolinglistid,a.assetpurchaseid || to_char(tl.lineno,'FM_0009') as newtoolinglistid from doc.toolinglist tl" &
                    " left join doc.toolinglistmst tm on tm.toolinglistid = tl.toolinglistid" &
                    " left join doc.assetpurchase a on a.id = tl.assetpurchaseid" &
                    " left join doc.toolingproject tp on tp.id = a.projectid" &
                    " where tl.assetpurchaseid = {1} and not tm.commontool and substring(tl.toolinglistid,1,1) <> 'S' ) foo" &
                    " where(m.toolinglistid = foo.toolinglistid);" &
                     " update doc.toolinglistmst m set toolinglistid = foo.newtoolinglistid" &
                    " from (select tl.toolinglistid,'S' || a.assetpurchaseid || to_char(tl.lineno,'FM_0009') as newtoolinglistid from doc.toolinglist tl" &
                    " left join doc.toolinglistmst tm on tm.toolinglistid = tl.toolinglistid" &
                    " left join doc.assetpurchase a on a.id = tl.assetpurchaseid" &
                    " left join doc.toolingproject tp on tp.id = a.projectid" &
                    " where tl.assetpurchaseid = {1} and not tm.commontool and substring(tl.toolinglistid,1,1) = 'S' ) foo" &
                    " where(m.toolinglistid = foo.toolinglistid);", Environment.UserDomainName.ToLower + "\" + Environment.UserName.ToLower, drv.Item("id"))
                Dim ra As Integer
                Dim mymessage As String = String.Empty
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, mymessage) Then
                    MessageBox.Show(mymessage)
                Else
                    Loadata()
                    MessageBox.Show(String.Format("{0}", "Done."))
                End If
            End If
        End If

    End Sub

    Private Sub UpdateToolingIdToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateToolingIdToolStripMenuItem.Click
        Dim drv As DataRowView = ToolingListDTBS.Current
        If Not IsNothing(drv) Then
            If MessageBox.Show(String.Format("Do you want to update this '{0}' tooling id ?", drv.Item("toolinglistid")), "Update Tooling Id", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                'show Dialog Input New ToolingId
                Dim myform As New DialogInputNewToolingID
                If myform.ShowDialog = Windows.Forms.DialogResult.OK Then                    
                    Dim sqlstr = String.Format("insert into doc.toolingidhistory(toolingid,changedby,changeddate,latesttoolingid) values(" &
                   " '{1}','{0}',now()::date,'{2}');" &
                   " update doc.toolinglistmst m set toolinglistid ='{2}'" &
                   " where(m.toolinglistid = '{1}');", Environment.UserDomainName.ToLower + "\" + Environment.UserName.ToLower, drv.Item("toolinglistid"), myform.Result)
                    Dim ra As Integer
                    Dim mymessage As String = String.Empty
                    If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, mymessage) Then
                        MessageBox.Show(mymessage)
                    Else
                        Loadata()
                        MessageBox.Show(String.Format("{0}", "Done."))
                    End If
                End If


            End If
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        HelperClass1.previewdoc(AttachmentBS, HelperClass1.attachment)
        'Try
        '    If Not IsNothing(DocumentBS.Current) Then
        '        Dim drv As DataRowView = DocumentBS.Current
        '        Dim filesource As String = String.Format("{2}\{0}{1}", drv.Row.Item("id"), drv.Row.Item("docext"), HelperClass1.document)
        '        If FileIO.FileSystem.GetFileInfo(filesource).Length / 1048576 < 5 Then
        '            Dim mytemp = String.Format("{1}{0}", drv.Row.Item("docname"), IO.Path.GetTempPath())

        '            FileIO.FileSystem.CopyFile(filesource, mytemp, True)
        '            Dim p As New System.Diagnostics.Process
        '            'p.StartInfo.FileName = "\\172.22.10.44\SharedFolder\PriceCMMF\New\template\Supplier Management Task User Guide-Admin.pdf"
        '            p.StartInfo.FileName = mytemp
        '            p.Start()
        '        Else
        '            MessageBox.Show("File too big.Please download.")
        '        End If
        '    End If
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try
    End Sub

    Private Sub TextBox28_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox28.TextChanged, TextBox34.TextChanged
        Try
            TextBox11.Text = String.Format("{0:#,##0.00}", TextBox28.Text * TextBox34.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox27_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox27.TextChanged, TextBox34.TextChanged
        Try
            TextBox8.Text = String.Format("{0:#,##0.00}", TextBox27.Text * TextBox34.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox26_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox26.TextChanged, TextBox34.TextChanged
        Try
            TextBox33.Text = String.Format("{0:#,##0.00}", TextBox26.Text * TextBox34.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox25_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox25.TextChanged, TextBox34.TextChanged
        Try
            TextBox18.Text = String.Format("{0:#,##0.00}", TextBox25.Text * TextBox34.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub AddToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddToolStripMenuItem.Click
        'new AssetPurchaseTracking

        Dim mydrv As DataRowView = TrackingBS.AddNew
        mydrv.BeginEdit()
        Dim drv As DataRowView = APBS.Current
        mydrv.Row.Item("assetpurchaseid") = drv.Row.Item("id")
        Dim message = "Tracking No :"
        Dim DefaultRespond As String = String.Empty
        Dim myRespond As String = String.Empty
        myRespond = InputBox(message, "Add Tracking No", DefaultRespond)
        If myRespond <> "" Then
            mydrv.Item("trackingno") = myRespond
            mydrv.EndEdit()
        Else
            mydrv.CancelEdit()
        End If

    End Sub

    Private Sub UpdateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateToolStripMenuItem.Click
        Dim mydrv As DataRowView = TrackingBS.Current
        mydrv.BeginEdit()
        Dim drv As DataRowView = APBS.Current       
        Dim message = "Tracking No :"
        Dim DefaultRespond As String = mydrv.Row.Item("trackingno")
        Dim myRespond As String = String.Empty
        myRespond = InputBox(message, "Update Tracking No", DefaultRespond)
        If myRespond <> "" Then
            mydrv.Item("trackingno") = myRespond
            mydrv.EndEdit()
        Else
            mydrv.CancelEdit()
        End If
    End Sub

    Private Sub DeleteToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem1.Click
        If Not IsNothing(TrackingBS.Current) Then

            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                'For Each drv As DataRowView In ListBox1.SelectedItems
                TrackingBS.RemoveCurrent()
                'Next

            End If
        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        If MessageBox.Show("Do you want to submit this task for validation?", "Submit for Validation", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            Dim drv As DataRowView = APBS.Current
            Dim approvaldrv As DataRowView = AssetPurchaseActionBS.AddNew
            If IsDBNull(drv.Item("status")) Then
                drv.Item("status") = AssetPurchaseStatusEnum.StatusNew
            Else
                Select Case drv.Item("status")
                    Case AssetPurchaseStatusEnum.StatusDraft
                        drv.Item("status") = AssetPurchaseStatusEnum.StatusNew
                        approvaldrv.Row.Item("remark") = "Creator Submit"
                    Case AssetPurchaseStatusEnum.StatusRejectedByPurchasing
                        drv.Item("status") = AssetPurchaseStatusEnum.StatusReSubmit
                        approvaldrv.Row.Item("remark") = "Creator Resubmit"
                    Case AssetPurchaseStatusEnum.StatusRejectedByControlling
                        drv.Item("status") = AssetPurchaseStatusEnum.StatusReSubmit
                        approvaldrv.Row.Item("remark") = "Creator Resubmit"
                End Select
               
            End If
            drv.EndEdit()

            approvaldrv.Row.Item("status") = drv.Item("status")
            approvaldrv.Row.Item("modifiedby") = HelperClass1.UserId
            approvaldrv.Row.Item("assetpurchaseid") = drv.Item("id")
            approvaldrv.EndEdit()
            drv.EndEdit()
            ToolStripButton2.PerformClick()
        End If
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        If MessageBox.Show("Do you want to validate this task?", "Validate", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            Dim drv As DataRowView = APBS.Current
            Dim approvaldrv As DataRowView = AssetPurchaseActionBS.AddNew

            Select Case drv.Item("status")
                Case AssetPurchaseStatusEnum.StatusReSubmit

                    If IsDBNull(drv.Item("approvalname2")) Then
                        'Validate by Purchasing with No approvalname2
                        drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByPurchasing
                        approvaldrv.Row.Item("remark") = "Purchasing Validator 1"
                    Else
                        'Validate by Purchasing with approvalname2 assigned                       
                        drv.Item("status") = AssetPurchaseStatusEnum.StatusFirstValidatedByPurchasing
                        approvaldrv.Row.Item("remark") = "Purchasing Validator 1"
                    End If
                Case AssetPurchaseStatusEnum.StatusNew
                    If IsDBNull(drv.Item("approvalname2")) Then
                        'Validate by Purchasing with No approvalname2
                        drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByPurchasing
                        approvaldrv.Row.Item("remark") = "Purchasing Validator 1"
                    Else
                        'Validate by Purchasing with approvalname2 assigned                       
                        drv.Item("status") = AssetPurchaseStatusEnum.StatusFirstValidatedByPurchasing
                        approvaldrv.Row.Item("remark") = "Purchasing Validator 1"
                    End If

                Case AssetPurchaseStatusEnum.StatusFirstValidatedByPurchasing
                    'Validate by Purchasing ApprovalName2
                    drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByPurchasing
                    approvaldrv.Row.Item("remark") = "Purchasing Validator 2"

                Case AssetPurchaseStatusEnum.StatusValidatedByPurchasing
                    'Controlling Validate the Task
                    'Show popup to input comments

                    Dim mydialog As New DialogAssetPurchaseComments(drv)
                    If mydialog.ShowDialog() = DialogResult.OK Then
                        drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByControlling
                        approvaldrv.Row.Item("remark") = "Controlling Team Validator"
                    Else
                        Exit Sub
                    End If

                    
            End Select
            approvaldrv.Row.Item("status") = drv.Item("status")
            approvaldrv.Row.Item("modifiedby") = HelperClass1.UserId
            approvaldrv.Row.Item("assetpurchaseid") = drv.Item("id")
            approvaldrv.EndEdit()
            drv.EndEdit()
            'ToolStripButton2.PerformClick()
            ToolStripButton2_Click_1(Me, New System.EventArgs)
        End If
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        If MessageBox.Show("Do you want to reject this task?", "Reject", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            Dim drv As DataRowView = APBS.Current
            Dim approvaldrv As DataRowView = AssetPurchaseActionBS.AddNew
            Select Case drv.Item("status")
                Case AssetPurchaseStatusEnum.StatusNew                  
                    'Rejected by Purchasing
                    drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing
                    approvaldrv.Row.Item("remark") = "Purchasing Validator 1"
                Case AssetPurchaseStatusEnum.StatusFirstValidatedByPurchasing
                    'Rejected by Purchasing ApprovalName2
                    drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByPurchasing
                    approvaldrv.Row.Item("remark") = "Purchasing Validator 2"
                Case AssetPurchaseStatusEnum.StatusValidatedByPurchasing
                    'Controlling Rejected the Task
                    'Show popup to input comments
                    Dim mydialog As New DialogAssetPurchaseComments(drv)
                    If mydialog.ShowDialog() = DialogResult.OK Then
                        drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByControlling
                        approvaldrv.Row.Item("remark") = "Controlling Team Validator"
                    Else
                        Exit Sub
                    End If
                    'drv.Item("status") = AssetPurchaseStatusEnum.StatusRejectedByControlling
                    'approvaldrv.Row.Item("remark") = "Controlling Team Validator"
            End Select
            approvaldrv.Row.Item("assetpurchaseid") = drv.Item("id")
            approvaldrv.Row.Item("status") = drv.Item("status")
            approvaldrv.Row.Item("modifiedby") = HelperClass1.UserId
            approvaldrv.EndEdit()
            drv.EndEdit()
            ToolStripButton2_Click_1(Me, New System.EventArgs)
            'ToolStripButton2.PerformClick()
        End If
    End Sub

    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        If MessageBox.Show("Do you want to change the status into completed?", "Complete", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            Dim drv As DataRowView = APBS.Current

            'If mydialog.ShowDialog() = DialogResult.OK Then
            '    drv.Item("status") = AssetPurchaseStatusEnum.StatusValidatedByControlling
            '    approvaldrv.Row.Item("remark") = "Controlling Team Validator"
            'Else
            '    Exit Sub
            'End If




            Dim approvaldrv As DataRowView = AssetPurchaseActionBS.AddNew
            Select Case drv.Item("status")
                Case AssetPurchaseStatusEnum.StatusValidatedByControlling
                    Dim mydialog As New DialogAssetPurchaseFinance(drv)
                    If mydialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        drv.Item("status") = AssetPurchaseStatusEnum.StatusCompleted
                        approvaldrv.Row.Item("remark") = "Accounting Team"
                    Else
                        Exit Sub
                    End If                    
                Case AssetPurchaseStatusEnum.StatusPendingPayment
                    Dim mydialog As New DialogAssetPurchaseFinance(drv)
                    If mydialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        drv.Item("status") = AssetPurchaseStatusEnum.StatusCompleted
                        approvaldrv.Row.Item("remark") = "Admin"
                    Else
                        Exit Sub
                    End If
            End Select
            approvaldrv.Row.Item("assetpurchaseid") = drv.Item("id")
            approvaldrv.Row.Item("status") = drv.Item("status")
            approvaldrv.Row.Item("modifiedby") = HelperClass1.UserId
            approvaldrv.EndEdit()
            drv.EndEdit()
            ToolStripButton2_Click_1(Me, New System.EventArgs)
            'ToolStripButton2.PerformClick()
        End If
    End Sub

    Private Sub ToolStripButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton9.Click
        If MessageBox.Show("Do you want to Cancel this task?", "Cancel", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            Dim drv As DataRowView = APBS.Current
            drv.Row.Item("status") = AssetPurchaseStatusEnum.StatusCancelled

            Dim approvaldrv As DataRowView = AssetPurchaseActionBS.AddNew            
            approvaldrv.Row.Item("assetpurchaseid") = drv.Item("id")
            approvaldrv.Row.Item("status") = drv.Item("status")
            approvaldrv.Row.Item("modifiedby") = HelperClass1.UserId
            approvaldrv.EndEdit()
            drv.EndEdit()
            ToolStripButton2_Click_1(Me, New System.EventArgs)
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim mydialog = New DialogFinanceAssetNo
        If mydialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            'Create tracking no
            Dim drv As DataRowView = APBS.Current
            Dim mydrv As DataRowView = TrackingBS.AddNew
            mydrv.row.item("trackingno") = mydialog.FinanceAssetNo
            mydrv.Row.Item("assetpurchaseid") = drv.Row.Item("id")
            mydrv.EndEdit()
            TextBox7.Text = TextBox7.Text + IIf(TextBox7.Text = "", "", ", ") + mydialog.FinanceAssetNo
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If MessageBox.Show("Do you want to clear the Finance Asset No ?", "Are you sure?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.OK Then
            For i = 0 To TrackingBS.Count - 1
                TrackingBS.RemoveCurrent()
            Next
            TextBox7.Clear()
        End If
    End Sub

    Private Sub ContextMenuStrip1Attachment_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1Attachment.Opening

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        If MessageBox.Show("Do you want to Create Proforma Purchase Order?", "Proforma PO", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.OK Then
            Dim dr As DataRowView
            If ProformaInvoiceBS.Count > 0 Then
                If MessageBox.Show(String.Format("Any modification in existing Proforma Purchase Order? {0}Click button Yes to create new one.", vbCrLf), "Proforma Purchase Order", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    dr = CreateProformaInvoice()
                    dr.EndEdit()

                    ToolStripButton2.PerformClick()
                    GeneratePDF(dr.Row("proformainvoice"))
                End If
            Else
                dr = CreateProformaInvoice()
                dr.EndEdit()
                Dim apdr = APBS.Current
                apdr.Row.Item("printdate") = Today.Date
                ToolStripButton2.PerformClick()
                GeneratePDF(dr.Row("proformainvoice"))
            End If
            'AddNew ProformaInvoice


            'Generate PDF file
        End If
    End Sub

    Private Function CreateProformaInvoice() As DataRowView
        Dim lineNo As Integer
        Dim APdr As DataRowView = APBS.Current
        lineNo = ProformaInvoiceBS.Count
        Dim dr = ProformaInvoiceBS.AddNew
        dr.item("proformainvoice") = String.Format("{0}_{1:yyyyMMdd}_{2:0000}", APdr.Row.Item("vendorcode"), APdr.Row.Item("applicantdate"), lineNo + 1)
        dr.item("creator") = HelperClass1.UserId
        dr.item("creationdate") = Today.Date

        myProformaPO = New proformaPO With {.proformapurchaseorder = dr.item("proformainvoice"),
                                            .applicantdate = APdr.Row.Item("applicantdate"),
                                            .applicantname = APdr.Row.Item("applicantname"),
                                            .assetdescription = APdr.Row.Item("assetdescription"),
                                            .origin = "" & APdr.Row.Item("origin"),
                                            .suppliercode = "" & APdr.Row.Item("toolingsupplier"),
                                            .suppliername = "" & APdr.Row.Item("toolingsuppliername"),
                                            .totaltoolingcost = OriginalCost,
                                            .originalcurrency = OriginalCurrency}
        Select Case APdr.Item("paymententity")
            Case "SEB Asia"
                myProformaPO.suppliercode = "" & APdr.Row.Item("vendorcode")
                myProformaPO.suppliername = "" & APdr.Row.Item("vendorname")
                myProformaPO.fax = "" & APdr.Row.Item("faxnumber")
                myProformaPO.tel = "" & APdr.Row.Item("telephone")
                myProformaPO.address = "" & APdr.Row.Item("vaaddress")
                myProformaPO.deliveryaddress = "" & APdr.Row.Item("vaaddress")
                myProformaPO.template = "\templates\ProformaPOSEBAsiaTemplate.xltx"
            Case "SEB SZ"
                myProformaPO.suppliercode = "" & APdr.Row.Item("toolingsupplier")
                myProformaPO.suppliername = "" & APdr.Row.Item("toolingsuppliername")
                myProformaPO.fax = "" & APdr.Row.Item("fax")
                myProformaPO.tel = "" & APdr.Row.Item("tel")
                myProformaPO.address = "" & APdr.Row.Item("address")
                myProformaPO.deliveryaddress = "" & APdr.Row.Item("deliveryaddress")
                myProformaPO.template = "\templates\ProformaPOTemplate.xltx"
        End Select
        If Not IsDBNull(APdr.Row.Item("expectedtoolingfinishdate")) Then
            myProformaPO.expectedtoolingfinishdate = APdr.Row.Item("expectedtoolingfinishdate")        
        End If
        If Not IsDBNull(APdr.Row.Item("printdate")) Then
            myProformaPO.printtdate = APdr.Row.Item("printdate")
        Else
            myProformaPO.printtdate = Date.Today
        End If

        Return dr
    End Function

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPreviewPPO.Click
        Dim dr = ProformaInvoiceBS.Current
        If Not IsNothing(dr) Then
            Dim Filename = String.Format("{0}\{1}.pdf", HelperClass1.proformapo, dr.item("proformainvoice"))
            HelperClass1.previewdoc(Filename)
        End If
        
    End Sub

    Private Sub GeneratePDF(ByVal p1 As Object)
        Dim ReportName As String = String.Format("{0}.pdf", p1)
        Dim filename = String.Format("{0}\{1}", HelperClass1.proformapo, ReportName)
        Dim myPdfCallback As FormatReportDelegate
        myPdfCallback = AddressOf CreateNewPDF
        'Dim mypdf = New ExportToExcelFile(Me, HelperClass1.proformapo, ReportName, myPdfCallback, template:="\templates\ProformaPOTemplate.xltx")
        Dim mypdf = New ExportToExcelFile(Me, HelperClass1.proformapo, ReportName, myPdfCallback, template:=myProformaPO.template)
        mypdf.CreateForm(filename, New EventArgs)
    End Sub

    Private Sub CreateNewPDF(ByRef sender As Object, ByRef e As EventArgs)
        Dim owb = DirectCast(sender, Excel.Workbook)
        Dim osheet As Excel.Worksheet = owb.Worksheets(1)

        osheet.Range("applicantdate").Value = myProformaPO.applicantdate
        osheet.Range("printdate").Value = myProformaPO.printtdate
        osheet.Range("proformapurchaseorder").Value = myProformaPO.proformapurchaseorder
        If myProformaPO.expectedtoolingfinishdate = #12:00:00 AM# Then
            osheet.Range("expectedtoolingfinishdate").Value = ""
        Else
            osheet.Range("expectedtoolingfinishdate").Value = myProformaPO.expectedtoolingfinishdate
        End If

        osheet.Range("origin").Value = myProformaPO.origin
        osheet.Range("applicantname").Value = myProformaPO.applicantname
        osheet.Range("suppliercode").Value = myProformaPO.suppliercode
        osheet.Range("suppliername").Value = myProformaPO.suppliername
        osheet.Range("address").Value = myProformaPO.address
        osheet.Range("tel").Value = myProformaPO.tel
        osheet.Range("fax").Value = myProformaPO.fax
        osheet.Range("originalcurrency").Value = myProformaPO.originalcurrency
        osheet.Range("assetdescription").Value = myProformaPO.assetdescription
        osheet.Range("deliveryaddress").Value = myProformaPO.deliveryaddress
        osheet.Range("totaltoolingcost").Value = myProformaPO.totaltoolingcost

    End Sub


    Private Sub BtnDownloadProformaPO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDownloadProformaPO.Click
        Me.validate()
        Dim DownloadCheckbox As Boolean = False
        For Each drv As DataRowView In ProformaInvoiceBS.List
            If drv.Row.Item("download") Then
                DownloadCheckbox = True
            End If
        Next
        If DownloadCheckbox Then
            If Not myThreadDownload.IsAlive Then
                'get Folder

                Dim myfolder = New SaveFileDialog
                myfolder.FileName = "SaveFile"
                If myfolder.ShowDialog = Windows.Forms.DialogResult.OK Then
                    ToolStripStatusLabel1.Text = ""
                    SelectedFolder = IO.Path.GetDirectoryName(myfolder.FileName)
                    myThreadDownload = New Thread(AddressOf DoDownloadProformaPO)
                    myThreadDownload.Start()
                End If

            Else
                MessageBox.Show("Please wait until the current process is finished.")
            End If
        Else
            MessageBox.Show("Please select file to download.")
        End If
        
    End Sub

    Private Sub DoDownloadProformaPO()
        Dim i As Integer = 0
        For Each drv As DataRowView In ProformaInvoiceBS.List
            If drv.Row.Item("download") Then
                i = i + 1
                Dim filename = String.Format("{0}\{1}.pdf", HelperClass1.proformapo, drv.Row.Item("proformainvoice"))
                ProgressReport(1, String.Format("Downloading ::{0} of {1} {2}", i, ProformaInvoiceBS.Count, drv.Row.Item("proformainvoice")))
                Dim filesource As String = String.Format("{0}", filename)
                Dim FileTarget As String = String.Format("{0}\{1}.pdf", SelectedFolder, drv.Row.Item("proformainvoice"))
                Try
                    FileIO.FileSystem.CopyFile(filesource, FileTarget, True)
                Catch ex As Exception

                End Try

            End If
        Next
        ProgressReport(1, "Done. Please check your folder ::" & SelectedFolder)
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim ToolingSupplierController = New ToolingSupplierAdapter
        Dim bshelper = ToolingSupplierController.GetToolingSupplierBS
        Dim myform = New FormHelper(bshelper)
        myform.DataGridView1.Columns(0).DataPropertyName = "description"
        If myform.ShowDialog = DialogResult.OK Then
            Dim drv As DataRowView = bshelper.Current
            Dim mydrv As DataRowView = APBS.Current

            mydrv.Row.Item("toolingsupplier") = drv.Row.Item("toolingsupplierid")
            mydrv.Row.Item("toolingsuppliername") = drv.Row.Item("toolingsuppliername")
            TextBox39.Text = drv.Row.Item("description")

        End If
    End Sub
End Class

Public Class TypeOfInvestment
    Public Property TypeOfInvestmentid As Integer
    Public Property TypeOfInvestmentName As String

    Public Sub New(ByVal _TypeOfInvestmentid As Integer, ByVal _TypeOfInvestmentName As String)
        Me.TypeOfInvestmentid = _TypeOfInvestmentid
        Me.TypeOfInvestmentName = _TypeOfInvestmentName
    End Sub
End Class

Public Class DocumentType
    Public Property DocumentTypeId As Integer
    Public Property DocumentTypeDesc As String

    Public Sub New(ByVal _documenttypeid As Integer, ByVal _documenttypedesc As String)
        Me.DocumentTypeId = _documenttypeid
        Me.DocumentTypeDesc = _documenttypedesc
    End Sub
End Class

Public Class PaymentMethod
    Public Property PaymentMethodId As Integer
    Public Property PaymentMethodDesc As String
    Public Sub New(ByVal _paymentmethodid As Integer, ByVal _paymentmethoddesc As String)
        Me.PaymentMethodDesc = _paymentmethoddesc
        Me.PaymentMethodId = _paymentmethodid
    End Sub
End Class

Public Class proformaPO
    Public Property applicantdate As Date
    Public Property printtdate As Date
    Public Property proformapurchaseorder As String
    Public Property expectedtoolingfinishdate As Date
    Public Property origin As String
    Public Property applicantname As String
    Public Property suppliercode As String
    Public Property suppliername As String
    Public Property address As String
    Public Property fax As String
    Public Property tel As String
    Public Property assetdescription As String
    Public Property deliveryaddress As String
    Public Property totaltoolingcost As Decimal
    Public Property originalcurrency As String
    Public Property template As String
End Class
