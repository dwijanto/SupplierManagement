Imports SupplierManagement.SharedClass
Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports System.Reflection


Public Class FormSupplierDashboard
    Const PURCHASING_CHARTER As Integer = 34
    Const SOCIAL_AUDIT As Integer = 22
    Const CONTRACT_GENERAL_CONTRACT As Integer = 32
    Const CONTRACT_QUALITY_APPENDIX As Integer = 33
    Const CONTRACT_SUPPLY_CHAIN_APPENDIX As Integer = 35
    Const PROJECT_SPECIFICATION As Integer = 36
    Const AUTHORIZATION_LETTERS As Integer = 18
    Const SIF As Integer = 21
    Const IDENTITY_SHEET As Integer = 54

    Const TURNOVER As Integer = 0
    Const SEBASIA_PLATFORM As Integer = 1
    Const NQSU As Integer = 2
    Const LOGISTICS As Integer = 3
    Const PROJECT_STATUS As Integer = 4
    Const SAP_VENDOR_PAYMENT As Integer = 5
    Const VFP_PROGRAM As Integer = 6


    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Dim myThreadQuery As New System.Threading.Thread(AddressOf DoQuery)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myThread2 As New System.Threading.Thread(AddressOf DoWork)
    Dim InitDS As DataSet
    Dim FilterSupplier As ClassSupplierDashBoard
    Dim DS As DataSet
    Dim myDS As DataSet
    Dim TurnoverBS As BindingSource
    Dim SEBTurnoverBS As BindingSource
    Dim SEBTurnoverBS1 As BindingSource
    Dim SEBTurnoverBS2 As BindingSource
    Dim SEBTurnoverBS3 As BindingSource
    Dim SEBTurnoverBS4 As BindingSource
    Dim FilterQtyBS As BindingSource
    Dim FilterAmountBS As BindingSource
    Dim GroupFPCPBS As BindingSource
    Dim GroupSBUBS As BindingSource
    Dim GroupFamilyBS As BindingSource
    Dim GroupBrandBS As BindingSource
    Dim NQSUBS As BindingSource

    Dim PurchasingFPCP As BindingSource
    Dim PurchasingFPCP1 As BindingSource
    Dim PurchasingFPCP2 As BindingSource
    Dim PurchasingFPCP3 As BindingSource
    Dim PurchasingFPCP4 As BindingSource    
    Dim PurchasingFP As BindingSource
    Dim PurchasingFP1 As BindingSource
    Dim PurchasingFP2 As BindingSource
    Dim PurchasingFP3 As BindingSource
    Dim PurchasingFP4 As BindingSource

    Dim LogisticsBS As BindingSource
    Dim LogisticsBS1 As BindingSource
    Dim LogisticsBS2 As BindingSource
    Dim LogisticsBS3 As BindingSource
    Dim LogisticsBS4 As BindingSource

    Dim PDBS As BindingSource
    Dim PDBS1 As BindingSource
    Dim PDBS2 As BindingSource
    Dim PDBS3 As BindingSource
    Dim PDBS4 As BindingSource

    Dim SAPBS As BindingSource
    Dim ContractBS As BindingSource

    Dim ALetterBS As BindingSource
    Dim PSpecBS As BindingSource
    Dim CitiProgramBS As BindingSource
    Dim VendorCurrencyBS As BindingSource

    Dim toCriteria2 As String = String.Empty 'criteria Shortname or VendorCode
    Dim ProductTypeCriteria As String = String.Empty

    Dim SDDS As DataSet
    Dim SBUBS As BindingSource
    Dim FamilyBS As BindingSource
    Dim BrandBS As BindingSource
    Dim ToolingHeader As BindingSource
    Dim ToolingDetailsAll As BindingSource
    Dim ToolingDetails As BindingSource
    Dim PCharterBS As BindingSource
    Dim SocialAuditBS As BindingSource

    Dim VendorFactoryContactBS As BindingSource
    Dim FactoryBS As BindingSource
    Dim ContactBS As BindingSource
    Dim SIFIDBS As BindingSource
    Dim SIFBS As BindingSource
    Dim IDBS As BindingSource

    Dim ContractualTerms1BS As BindingSource
    Dim ContractualTerms2BS As BindingSource
    Dim ContractualTerms3BS As BindingSource

    Dim PriceCMMFBS As BindingSource

    Private _shortname As String
    Private _vendorcode As String
    Private _currentdate As Date
    Private SelectedTab As Boolean = False

    Public PurchaseAmountSQL As String

    Private bFAdapter As BudgetForecastAdapter

    Enum VendorQuery
        vendorcode
        shortname
    End Enum

    Dim VendorQueryType As VendorQuery
    Private Sub ShowData()
        ''ClassSEBTO1.ClearValue()
        ''ClassFilterTO1.ClearValue()
        ''ClassSIFTO1.ClearValue()
        'UcGroupBy1.ClearValue()
        'UcGroupBy2.ClearValue()
        'UcGroupBy3.ClearValue()
        'UcGroupBy4.ClearValue()
        'UcChart1.ClearValue()
        ''UcScorecardDetail1.ClearValue()
        'Button2.PerformClick()
        'If String.Format("{0:yyyyMM}", DateTimePicker2.Value) <> String.Format("{0:yyyyMM}", DateTimePicker3.Value) Then
        '    Button1.PerformClick()
        'End If
    End Sub

    Public Sub LetsClearTo()
        ClassSEBTO1.ClearValue()
        ClassFilterTO1.ClearValue()
        ClassSIFTO1.ClearValue()
        UcScorecardDetail1.ClearValue()
        UcGroupBy1.ClearValue() 'Set bs with nothing, displaygrid after that
        UcGroupBy2.ClearValue()
        UcGroupBy3.ClearValue()
        UcGroupBy4.ClearValue()
        UcChart1.ClearValue()

        Button2.PerformClick()
        If String.Format("{0:yyyyMM}", DateTimePicker2.Value) <> String.Format("{0:yyyyMM}", DateTimePicker3.Value) Then
            Button1.PerformClick()
        End If



    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        InitData()
        AddHandler UcSupplierDashboard1.LoadFinish, AddressOf LetsClearTo
        'AddHandler UcSupplierDashboard1.FinishQuery3, AddressOf ShowData
        AddHandler UcSupplierDashboard1.FinishQuery5, AddressOf RefreshSupplierBasicInformation 'Supplier Technology & SupplierPanel History
        AddHandler UcSupplierDashboard1.QueryShortVendorname, AddressOf ShowProgress
        AddHandler UcSupplierDashboard1.StatusVendorChanged, AddressOf ChangeVendorList
        ' Add any initialization after the InitializeComponent() call.


        UcSupplierDashboard1.SupplierInfo = UcSupplierInfo1
        FilterSupplier = New ClassSupplierDashBoard(UcSupplierDashboard1)
        UcSupplierDashboard1.currentDate = DateTimePicker1.Value.Date

        AddHandler UcContractualTermDGV1.ShowPopUp, AddressOf ShowContractPopUp

    End Sub

    Private Sub InitData()
        If Not myThreadQuery.IsAlive Then

            ToolStripStatusLabel1.Text = ""
            myThreadQuery = New Thread(AddressOf DoQuery)
            myThreadQuery.Start()
        Else
            MessageBox.Show(String.Format("{0} Please wait until the current process is finished.", System.Reflection.MethodInfo.GetCurrentMethod()))
        End If
    End Sub

    Sub DoWork()

    End Sub

    Sub DoQuery()
        Dim sb As New StringBuilder
        sb.Append("select distinct period from doc.turnover order by period desc limit 1; select latestupdate,auditname from doc.audit order by linenumber;select period from doc.nqsu order by period desc limit 1")

        Dim mymessage As String = String.Empty
        Dim result As Integer
        InitDS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString, InitDS) Then
            ProgressReport(5, result.ToString)
        Else
            MessageBox.Show(mymessage)
        End If
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'clear Filter
        'ClassFilterTO1.ClearValue()

        ProgressReport(1, "Processing.. Please wait.")
        ProgressReport(7, "Marquee")
        _currentdate = DateTimePicker1.Value.Date
        _shortname = UcSupplierDashboard1.ShortName
        _vendorcode = UcSupplierDashboard1.Vendorcode
        If (IsNothing(_shortname) And _vendorcode = 0) Then
            'MessageBox.Show("Please select Short Name or VendorCode")
            UcSupplierDashboard1.ErrorProvider1.SetError(UcSupplierDashboard1.TextBox1, "Please select Short Name or Vendor Name.")
            UcSupplierDashboard1.ErrorProvider1.SetError(UcSupplierDashboard1.TextBox2, "Please select Vendor Name or Vendor Name.")
            Exit Sub
        End If
        UcSupplierDashboard1.ErrorProvider1.SetError(UcSupplierDashboard1.TextBox1, "")
        UcSupplierDashboard1.ErrorProvider1.SetError(UcSupplierDashboard1.TextBox2, "")

        UctoHeader1.CurrentDate = _currentdate
        UcScorecardHeader1.CurrentDate = _currentdate
        UcSupplierDashboard1.GetPanelStatusSupplier()
        If UcSupplierDashboard1.ProductTypeFilter = "" Then
            UcSupplierDashboard1.ProductTypeFilter = "ALL"
        End If


        fillSIF()
    End Sub

    Private Sub fillSIF()
        'Dim myyear As Integer
        If Not myThreadQuery.IsAlive Then

            ToolStripStatusLabel1.Text = ""
            myThreadQuery = New Thread(AddressOf DoFillSIF)
            myThreadQuery.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If

    End Sub

    Private Sub DoFillSIF()
        ProgressReport(1, "Processing.. Please wait.")

        Dim sb As New StringBuilder
        Dim mycriteria As String = String.Empty
        Dim mycriteria2 As String = String.Empty
        Dim toCriteria As String = String.Empty
        Dim tocriteriafpcp As String = String.Empty
        Dim tocriteriafp As String = String.Empty
        Dim SebFilterApplied As String = String.Empty
        'Dim toCriteria2 As String = String.Empty 'criteria Shortname or VendorCode
        Dim toFilter As String = String.Empty
        Dim TextFilter As String = String.Empty
        Dim TextFilterForPriceCMMF As String = String.Empty
        Dim StatusFilter As String = String.Empty
        'Dim ProductTypeCriteria As String = String.Empty
        Dim BFCriteria As String = String.Empty

        StatusFilter = UcSupplierDashboard1.GetStatusCriteria.Replace("status", "paramname")
        TextFilter = UcSupplierDashboard1.getFilterCriteria.Replace("status", "paramname")
        BFCriteria = TextFilter
        TextFilterForPriceCMMF = TextFilter.Replace("familyname", "tui.familyname")
        SebFilterApplied = If(UcSupplierDashboard1.CheckBox1.Checked, " and (not sf.cmmf isnull)", "")
        ProductTypeCriteria = UcSupplierDashboard1.ProductTypeCriteria.Replace("''", "'")
        'BFCriteria = BFCriteria & ProductTypeCriteria

        'SIF
        If Not IsNothing(_shortname) Then
            VendorQueryType = VendorQuery.shortname
            mycriteria = String.Format(" where v.shortname = '{0}' ", _shortname)
            mycriteria2 = String.Format(" where myyear <= {0} ", Year(_currentdate))
            toCriteria = String.Format(" where year = {0} and period <= {1:yyyyMM} and v.shortname = '{2}' {3}", Year(_currentdate), _currentdate, _shortname, ProductTypeCriteria)
            tocriteriafp = String.Format(" where year = {0} and period <= {1:yyyyMM} and v.shortname = '{2}' {3}", Year(_currentdate), _currentdate, _shortname, " and (fpcp in ('FP') ) ")
            tocriteriafpcp = String.Format(" where year = {0} and period <= {1:yyyyMM} and v.shortname = '{2}' {3}", Year(_currentdate), _currentdate, _shortname, " and (fpcp in ('FP','CP') ) ")
            toCriteria2 = String.Format(" where v.shortname = '{0}'", _shortname)
            'sb.Append(String.Format("select shortname,myyear,sum(turnovery) as turnovery,sum(turnovery1) as turnovery1,sum(turnovery2) as turnovery2,sum(turnovery3) as turnovery3 ,sum(turnovery4) as turnovery4 from (select distinct v.shortname,s.* from doc.sif s left join doc.document d on d.id = s.documentid " &
            '          " left join doc.vendordoc vd on vd.documentid = d.id left join vendor v on v.vendorcode = vd.vendorcode {0} " &
            '          " )foo {1}" &
            '          " group by shortname,myyear order by myyear desc limit 1;", mycriteria, mycriteria2))
            sb.Append(String.Format("with sif as (select distinct shortname,s.* from   (select v.shortname,first_value(s.documentid)  over(partition by myyear order by myyear,s.documentid desc) as docid,myyear from doc.sif s" &
                                    " left join doc.document d on d.id = s.documentid  left join doc.vendordoc vd on vd.documentid = d.id  left join vendor v on v.vendorcode = vd.vendorcode {0} order by myyear desc) as foo" &
                                    " left join doc.sif s on s.documentid = foo.docid ) " &
                                    " select shortname,myyear,sum(turnovery) as turnovery,sum(turnovery1) as turnovery1,sum(turnovery2) as turnovery2,sum(turnovery3) as turnovery3 ,sum(turnovery4) as turnovery4 from sif " &
                                    " {1} group by shortname,myyear order by myyear desc limit 1;", mycriteria, mycriteria2))
        Else
            VendorQueryType = VendorQuery.vendorcode
            mycriteria = String.Format(" where v.vendorcode = {0} and myyear <= {1}  order by myyear desc limit 1", _vendorcode, Year(_currentdate))
            toCriteria = String.Format(" where year = {0} and period <= {1:yyyyMM} and v.vendorcode = {2} {3}", Year(_currentdate), _currentdate, _vendorcode, ProductTypeCriteria)

            tocriteriafp = String.Format(" where year = {0} and period <= {1:yyyyMM} and v.vendorcode = {2} {3}", Year(_currentdate), _currentdate, _vendorcode, " and (fpcp in ('FP') ) ")
            tocriteriafpcp = String.Format(" where year = {0} and period <= {1:yyyyMM} and v.vendorcode = {2} {3}", Year(_currentdate), _currentdate, _vendorcode, " and (fpcp in ('FP','CP') ) ")

            toCriteria2 = String.Format(" where v.vendorcode = {0} ", _vendorcode)
            sb.Append(String.Format("select v.shortname,vd.vendorcode,s.* from doc.sif s" &
                  " left join doc.document d on d.id = s.documentid" &
                  " left join doc.vendordoc vd on vd.documentid = d.id" &
                  " left join vendor v on v.vendorcode = vd.vendorcode {0};", mycriteria))

        End If


        'Turnover,PriceIndex
        'YearToDate
        sb.Append(String.Format("select d.year,sum(qty) as qty,sum(amount) as amount,sum(tovariance) as tovariance,sum(towavpymin1) as towavpymin1,sum(towlkpmin1) as towlkpmin1,sum(tovariance) as tovariance, sum(amount)/sum(towavpymin1) * 100 as piavg,sum(amount)/sum(towlkpmin1) * 100 as pilkp,sum(amount) / sum(towstd) * 100 as pistd" &
                      " from doc.turnover d " &
                      " left join vendor v on v.vendorcode = d.vendorcode " &
                       " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                      "{0} {1}  group by year;", toCriteria, StatusFilter))

        'Y-1 to Y-4
        For i = 1 To 4
            sb.Append(String.Format("select d.year,sum(qty) as qty,sum(amount) as amount,sum(tovariance) as tovariance,sum(towavpymin1) as towavpymin1,sum(towlkpmin1) as towlkpmin1,sum(tovariance) as tovariance, sum(amount)/sum(towavpymin1) * 100 as piavg,sum(amount)/sum(towlkpmin1) * 100 as pilkp,sum(amount) / sum(towstd) * 100 as pistd " &
                      " from doc.turnover d " &
                      " left join vendor v on v.vendorcode = d.vendorcode" &
                      " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                      " {0} and year = {1} {2} {3}" &
                      " group by year;", toCriteria2, Year(_currentdate) - i, ProductTypeCriteria, StatusFilter))
        Next

        'Filter qty
        sb.Append(String.Format("select * from crosstab('select {0},year,sum(qty) from doc.turnover tu" &
                  " left join vendor v on v.vendorcode = tu.vendorcode" &
                  " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                  " left join materialmaster mm on mm.cmmf = tu.cmmf" &
                  " left join family f on f.familyid = tu.comfam" &
                  " left join doc.sebplatform sf on sf.cmmf = tu.cmmf" &
                  "  {1}  and period <= {2:yyyyMM} and year >= {3} {5}" &
                  " group by {0},year  ','select m from generate_series({3},{4}) m order by m desc') as (sbu text, amount1 numeric,amount2 numeric,amount3 numeric,amount4 numeric,amount5 numeric);", "v." + VendorQueryType.ToString, toCriteria2.Replace("'", "''"), _currentdate, Year(_currentdate) - 4, Year(_currentdate), TextFilter))
        'Filter Amount
        sb.Append(String.Format("select * from crosstab('select {0},year,sum(amount) from doc.turnover tu" &
                  " left join vendor v on v.vendorcode = tu.vendorcode" &
                  " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                  " left join materialmaster mm on mm.cmmf = tu.cmmf" &
                  " left join family f on f.familyid = tu.comfam" &
                  " left join doc.sebplatform sf on sf.cmmf = tu.cmmf" &
                  "  {1}  and period <= {2:yyyyMM} and year >= {3} {5}" &
                  " group by {0},year  ','select m from generate_series({3},{4}) m order by m desc') as (sbu text, amount1 numeric,amount2 numeric,amount3 numeric,amount4 numeric,amount5 numeric);", "v." + VendorQueryType.ToString, toCriteria2.Replace("'", "''"), _currentdate, Year(_currentdate) - 4, Year(_currentdate), TextFilter))

        'Group By ProductType
        Dim myGroup() As String = {"tu.fpcp", "tu.sbu", "f.familyname", "tu.brand"}
        Dim myLabel() As String = {"ProductType", "SBU", "Family", "Brand"}
        purchaseAmountSQL = ""
        For i = 0 To 3
            TextFilter = myGroup(i)
            Dim mysql As String
            'sb.Append(String.Format("select *,(amount1/amount2) - 1 as trend1,(amount2/amount3) -1 as trend2,(amount3 /amount4 )- 1 as trend3,amount4/amount5 - 1 as trend4 from crosstab('select {5}::text,year,sum(amount) from doc.turnover tu" &
            '      " left join vendor v on v.vendorcode = tu.vendorcode" &
            '      " left join materialmaster mm on mm.cmmf = tu.cmmf" &
            '      " left join family f on f.familyid = tu.comfam" &
            '      "  {1}  and period <= {2:yyyyMM} and year >= {3} {6}" &
            '      " group by {5},year order by {5} ','select m from generate_series({3},{4}) m order by m desc') as (sbu text, amount1 numeric,amount2 numeric,amount3 numeric,amount4 numeric,amount5 numeric);", "v." + VendorQueryType.ToString, toCriteria2.Replace("'", "''"), _currentdate, Year(_currentdate) - 4, Year(_currentdate), TextFilter, ProductTypeCriteria.Replace("'", "''")))
            mysql = String.Format("select '{7}' || row_number() over (order by amount1 desc)::text as label, * from ( select " &
                                  " sbu, case when amount1 isnull then 0 else amount1 end as amount1,case when amount2 isnull then 0 else amount2 end as amount2," &
                                  " case when amount3 isnull then 0 else amount3 end as amount3,case when amount4 isnull then 0 else amount4 end as amount4," &
                                  " case when amount5 isnull then 0 else amount5 end as amount5" &
                                  ",(amount1/amount2) - 1 as trend1,(amount2/amount3) -1 as trend2,(amount3 /amount4 )- 1 as trend3,amount4/amount5 - 1 as trend4 from crosstab('select {5}::text,year,sum(amount) as amount from doc.turnover tu" &
                 " left join vendor v on v.vendorcode = tu.vendorcode" &
                 " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                 " left join materialmaster mm on mm.cmmf = tu.cmmf" &
                 " left join family f on f.familyid = tu.comfam" &
                 "  {1}  and period <= {2:yyyyMM} and year >= {3} {6} {8}" &
                 " group by {5},year having sum(amount) <> 0 order by {5}  ','select m from generate_series({3},{4}) m order by m desc') as (sbu text, amount1 numeric,amount2 numeric,amount3 numeric,amount4 numeric,amount5 numeric)) as foo order by amount1 desc;", "v." + VendorQueryType.ToString, toCriteria2.Replace("'", "''"), _currentdate, Year(_currentdate) - 4, Year(_currentdate), TextFilter, ProductTypeCriteria.Replace("'", "''"), myLabel(i), StatusFilter.Replace("'", "''"))
            sb.Append(mysql)
            PurchaseAmountSQL = PurchaseAmountSQL + IIf(PurchaseAmountSQL.Length > 0, " union all ", "") + "(" + mysql.Replace(";", "") + ")"
        Next
        PurchaseAmountSQL = PurchaseAmountSQL + IIf(PurchaseAmountSQL.Length > 0, ";", "")
        'NQSU
        'sb.Append(String.Format("select sbu,case when amount1 isnull then 0 else amount1 end as amount1,case when amount2 isnull then 0 else amount2 end as amount2," &
        '                        " case when amount3 isnull then 0 else amount3 end as amount3,case when amount4 isnull then 0 else amount4 end as amount4," &
        '                        " case when amount5 isnull then 0 else amount5 end as amount5 from (select * from crosstab('select * from ((select {2} ,year,(sum(criticaldefectytd) + sum(majordefectytd)) * 1000000 / sum(samplesizeytd)::numeric as nqsu from doc.nqsu tu " &
        '         " left join vendor v on v.vendorcode = tu.vendorcode" &
        '         " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
        '         " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
        '         "  {3} {5} and period = {0:yyyyMM} and samplesizeytd > 0 group by {2},year  )" &
        '         " union all" &
        '         " (with s as (select v.vendorcode,max(period) as period" &
        '         " from doc.nqsu  n" &
        '         " left join vendor v on v.vendorcode = n.vendorcode " &
        '         " {3} and Year < {0:yyyy} group by v.vendorcode,year order by v.vendorcode,period desc)" &
        '         " select {2},year,(sum(criticaldefectytd) + sum(majordefectytd)) * 1000000 / sum(samplesizeytd)::numeric as nqsu from s " &
        '         " left join vendor v on v.vendorcode = s.vendorcode " &
        '         " left join doc.nqsu n on n.vendorcode = s.vendorcode and n.period = s.period and samplesizeytd > 0 group by {2},year)) as foo order by foo.{4}'," &
        '         " 'select m from generate_series({1},{0:yyyy}) m order by m desc')" &
        '         " as (sbu text, amount1 numeric,amount2 numeric,amount3 numeric,amount4 numeric,amount5 numeric)) as foo ;", _currentdate, Year(_currentdate) - 4, "v." + VendorQueryType.ToString, toCriteria2.Replace("'", "''"), VendorQueryType.ToString, StatusFilter.Replace("'", "''")))

        sb.Append(String.Format("select sbu,case when amount1 isnull then 0 else amount1 end as amount1,case when amount2 isnull then 0 else amount2 end as amount2," &
                                " case when amount3 isnull then 0 else amount3 end as amount3,case when amount4 isnull then 0 else amount4 end as amount4," &
                                " case when amount5 isnull then 0 else amount5 end as amount5 from (select * from crosstab('select * from ((select {2} ,year,(sum(criticaldefectytd) + sum(majordefectytd)) * 1000000 / sum(samplesizeytd)::numeric as nqsu from doc.nqsu tu " &
                 " left join vendor v on v.vendorcode = tu.vendorcode" &
                 " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                 " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                 "  {3} {5} and period <= {0:yyyyMM} and samplesizeytd > 0 group by {2},year,period  order by {2},year desc,period desc limit 1 )" &
                 " union all" &
                 " (with s as (select v.vendorcode,max(period) as period" &
                 " from doc.nqsu  n" &
                 " left join vendor v on v.vendorcode = n.vendorcode " &
                 " {3} and Year < {0:yyyy} group by v.vendorcode,year order by v.vendorcode,period desc)" &
                 " select {2},year,(sum(criticaldefectytd) + sum(majordefectytd)) * 1000000 / sum(samplesizeytd)::numeric as nqsu from s " &
                 " left join vendor v on v.vendorcode = s.vendorcode " &
                 " left join doc.nqsu n on n.vendorcode = s.vendorcode and n.period = s.period and samplesizeytd > 0 group by {2},year)) as foo order by foo.{4}'," &
                 " 'select m from generate_series({1},{0:yyyy}) m order by m desc')" &
                 " as (sbu text, amount1 numeric,amount2 numeric,amount3 numeric,amount4 numeric,amount5 numeric)) as foo ;", _currentdate, Year(_currentdate) - 4, "v." + VendorQueryType.ToString, toCriteria2.Replace("'", "''"), VendorQueryType.ToString, StatusFilter.Replace("'", "''")))

        'Logistics     
        'YearToDate
        Dim groupby As String
        If VendorQueryType = VendorQuery.shortname Then
            groupby = "v.shortname"
        Else
            groupby = "v.vendorcode"
        End If
        sb.Append(String.Format("select  l.year,(sum(sasl) / sum(weight) )as sasl,(sum(ssl) / sum(weight)) as ssl,avg(leadtime) as lt,sum(orderno) as orderno,sum(shipment) as shipment, sum(sslnet)/sum(weight) as sslnet" &
                                " from doc.logistics l" &
                                " left join vendor v on v.vendorcode = l.vendorcode" &
                                 " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                                " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                                " {0} {1} {2}" &
                                " group by {3},year ;", toCriteria, ProductTypeCriteria, StatusFilter, groupby))

        'Y-1 to Y-4
        For i = 1 To 4
            sb.Append(String.Format("select  l.year,(sum(sasl) / sum(weight) )as sasl,(sum(ssl) / sum(weight)) as ssl,avg(leadtime) as lt,sum(orderno) as orderno,sum(shipment) as shipment, sum(sslnet)/sum(weight) as sslnet " &
                               " from doc.logistics l" &
                               " left join vendor v on v.vendorcode = l.vendorcode " &
                                " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                                " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                               " {0} and year = {1} {2} {3}" &
                               " group by {4},year ;", toCriteria2, Year(_currentdate) - i, ProductTypeCriteria, StatusFilter, groupby))
        Next


        'PDProject
        'YearToDate
        sb.Append(String.Format("select  pp.* " &
                                " from doc.pdproject pp" &
                                " left join vendor v on v.shortname = pp.shortname" &
                                " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                                " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                                " {0} and period = {1:yyyyMM} {2} {3}" &
                                " limit 1 ;", toCriteria2, _currentdate, ProductTypeCriteria, StatusFilter))

        'Y-1 to Y-4
        For i = 1 To 4
            sb.Append(String.Format("select  pp.* " &
                               " from doc.pdproject pp" &
                               " left join vendor v on v.shortname = pp.shortname " &
                               " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                                " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                               " {0} and year = {1} {2} order by period desc " &
                               " limit 1 ;", toCriteria2, Year(_currentdate) - i, ProductTypeCriteria, StatusFilter))
        Next

        'Contract term
        sb.Append(String.Format("select v.vendorcode,v.vendorname::text,pt.payt || ' - ' || case when pterm.details isnull then '' else pterm.details end,pdate.effectivedate,pr.paramname as status,v.vendorcode::text || ' - ' ||  pt.payt || ' ' ||  case when pterm.details isnull then '' else pterm.details end || ' - ' || case when pr.paramname isnull then '' else pr.paramname end " &
                                " || case when pdate.effectivedate isnull then '' else ' - (Effective Date ' ||" &
                                " to_char(pdate.effectivedate,'DD-Mon-YYYY') || ')'  End " &
                                "as sappaymentterm from doc.vendorpayt pt" &
                  " left join vendor v on v.vendorcode = pt.vendorcode" &
                  " left join paymentterm pterm on pterm.payt = pt.payt" &
                  " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                  " left join doc.paymenttermdate pdate on pdate.vendorcode = v.vendorcode and pdate.paymenttermid = pterm.paymenttermid" &
                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                  "{0} {1};", toCriteria2, StatusFilter))

        sb.Append(String.Format("with data as (select distinct v.shortname,first_value(vd.documentid) over (partition by v.shortname,doctypeid order by docdate desc) as id,doctypeid from doc.vendordoc vd" &
                  " left join doc.document d on d.id = vd.documentid " &
                  " left join doc.doctype dt on dt.id = d.doctypeid" &
                  " left join vendor v on v.vendorcode = vd.vendorcode" &
                  " where doctypeid in ({1},{2},{3})" &
                  " order by shortname)" &
                  " select * from data left join doc.document d on d.id = data.id left join doc.generalcontract gc on gc.documentid = d.id " &
                  " left join paymentterm p on p.paymenttermid = gc.paymentcode left join doc.qualityappendix qa on qa.documentid = d.id " &
                  " left join doc.supplychain sc on sc.documentid = d.id left join doc.docexpired de on de.documentid = d.id" &
                  "  where shortname = '{0}' ;", UcSupplierInfo1.Shortname, CONTRACT_GENERAL_CONTRACT, CONTRACT_QUALITY_APPENDIX, CONTRACT_SUPPLY_CHAIN_APPENDIX))
        '36 AuthLetter & 18 Project Spec
        sb.Append(String.Format("with data as (select v.shortname,vd.documentid as id,doctypeid from doc.vendordoc vd" &
                  " left join doc.document d on d.id = vd.documentid left join doc.doctype dt on dt.id = d.doctypeid " &
                  " left join vendor v on v.vendorcode = vd.vendorcode where doctypeid in ({1},{2}) order by shortname)" &
                  " select d.doctypeid,d.docdate,dp.projectname,d.remarks,v.version,de.expireddate,docdate::text ||  case when remarks isnull then '' else ' ' || remarks end || case when version isnull then '' else ' ' || version end || ' ' ||projectname || case when expireddate isnull then '' else ' ' || expireddate end as description, d.id,d.docname,d.docext ,ps.returnrate " &
                  " from data  left join doc.document d on d.id = data.id left join doc.project dp on dp.documentid = d.id" &
                  " left join doc.projectspecification ps on ps.documentid = d.id left join doc.docexpired de on de.documentid = d.id left join doc.version v on v.documentid = d.id" &
                  " where shortname = '{0}' order by d.doctypeid,docdate desc;", UcSupplierInfo1.Shortname, PROJECT_SPECIFICATION, AUTHORIZATION_LETTERS))
        sb.Append(String.Format("select cp.*,v.shortname from doc.citiprogram cp" &
                  " left join vendor v on v.vendorcode = cp.vendorcode {0} {1};", toCriteria2, ProductTypeCriteria))

        'Tooling Header
        'sb.Append(String.Format("select foo.project,doc.validnum(foo.amount1) as newasset,doc.validnum(foo.amount2) as toolingmodification, doc.validnum(foo.amount1) +doc.validnum(foo.amount2) as total from (select * from crosstab('select  trim(tp.projectname) as project,ap.typeofinvestment,sum(tl.cost) from doc.assetpurchase ap " &
        '          "inner join doc.toolingproject tp on tp.id = ap.projectid left join vendor v on v.vendorcode = ap.vendorcode left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
        '          " {0} group by project,ap.typeofinvestment order by project,typeofinvestment','select m from generate_series(1,2)m order by m asc') as (project text,amount1 numeric,amount2 numeric))foo;", toCriteria2.Replace("'", "''")))
        'sb.Append(String.Format("select foo.project,doc.validnum(foo.amount1) as newasset,doc.validnum(foo.amount2) as toolingmodification, doc.validnum(foo.amount1) +doc.validnum(foo.amount2) as total,applicant from (select * from crosstab('with asset as (select first_value(applicantdate) over(partition by ap.projectid ) as applicantdate, trim(tp.projectname) as project," &
        '                        " ap.typeofinvestment,tl.cost,ap.projectid, ap.vendorcode,ap.id from doc.assetpurchase ap inner join doc.toolingproject tp on tp.id = ap.projectid " &
        '                        " left join vendor v on v.vendorcode = ap.vendorcode left join doc.toolinglist tl on tl.assetpurchaseid = ap.id  " &
        '                        " {0}) select applicantdate, project,ap.typeofinvestment,sum(cost) from asset ap  " &
        '                        " group by applicantdate,project,ap.typeofinvestment order by project,typeofinvestment','select m from generate_series(1,2)m order by m asc') as (applicant date,project text,amount1 numeric,amount2 numeric))foo;", toCriteria2.Replace("'", "''")))

        'sb.Append(String.Format("select foo.project,doc.validnum(foo.amount1) as newasset,doc.validnum(foo.amount2) as toolingmodification, doc.validnum(foo.amount1) +doc.validnum(foo.amount2) as total,applicant from (select * from crosstab('with asset as (select first_value(applicantdate) over(partition by ap.projectid ) as applicantdate, trim(tp.projectname) as project," &
        '                        " ap.typeofinvestment,tl.cost,ap.projectid, ap.vendorcode,ap.id from doc.assetpurchase ap inner join doc.toolingproject tp on tp.id = ap.projectid " &
        '                        " left join vendor v on v.vendorcode = ap.vendorcode left join doc.toolinglist tl on tl.assetpurchaseid = ap.id  " &
        '                        " {0}) select min(applicantdate) as applicantdate, project,ap.typeofinvestment,sum(cost) from asset ap  " &
        '                        " group by project,ap.typeofinvestment order by project,typeofinvestment','select m from generate_series(1,2)m order by m asc') as (applicant date,project text,amount1 numeric,amount2 numeric))foo;", toCriteria2.Replace("'", "''")))

        sb.Append(String.Format("with asset as (select distinct first_value(applicantdate) over(partition by trim(tp.projectname) ) as applicantdate, trim(tp.projectname) as project from doc.assetpurchase ap inner join doc.toolingproject tp on tp.id = ap.projectid  left join vendor v on v.vendorcode = ap.vendorcode left join doc.toolinglist tl on tl.assetpurchaseid = ap.id   {0}  ) " &
                                " select foo.project,doc.validnum(foo.amount1) as newasset,doc.validnum(foo.amount2) as toolingmodification, doc.validnum(foo.amount1) +doc.validnum(foo.amount2) as total ,a.applicantdate as applicant from (select * from crosstab('with asset as (select first_value(applicantdate) over(partition by ap.projectid ) as applicantdate, trim(tp.projectname) as project, ap.typeofinvestment,tl.cost,ap.projectid, ap.vendorcode,ap.id from doc.assetpurchase ap inner join doc.toolingproject tp on tp.id = ap.projectid  left join vendor v on v.vendorcode = ap.vendorcode left join doc.toolinglist tl on tl.assetpurchaseid = ap.id    {1}) select project,ap.typeofinvestment,sum(cost) from asset ap   group by project,ap.typeofinvestment order by project,typeofinvestment','select m from generate_series(1,2)m order by m asc') as (project text,amount1 numeric,amount2 numeric))foo" &
                                " left join asset a on a.project = foo.project order by applicant desc;", toCriteria2, toCriteria2.Replace("'", "''")))

        sb.Append(String.Format("select  tl.toolinglistid,trim(tp.projectname) as project, tl.suppliermoldno,tl.toolsdescription,tl.cost,f.familyname::text,ap.applicantdate,doc.getinvestmentname(ap.typeofinvestment)as atypeofinvestment,ap.aeb,ap.budgetamount,ap.comments,ap.financeassetno,ap.applicantname, s.sbuname::text, ap.vendorcode, v.vendorname::text, ap.assetpurchaseid" &
                                " from doc.assetpurchase ap inner join doc.toolingproject tp on tp.id = ap.projectid left join family f on f.familyid = tp.familyid left join vendor v on v.vendorcode = ap.vendorcode left join doc.toolinglist tl on tl.assetpurchaseid = ap.id left join sbu s on s.sbuid = f.sbuid" &
                                " {0}" &
                                " order by applicantdate desc,atypeofinvestment,tl.toolinglistid,project;", toCriteria2))
        'Sustainable Development
        sb.Append(String.Format("select distinct d.docdate::date,d.remarks,ver.version, d.id,d.docname,d.docext from doc.vendordoc vd" &
                                " left join doc.document d on d.id = vd.documentid left join doc.version ver on ver.documentid = d.id left join vendor v on v.vendorcode = vd.vendorcode " &
                                " where(d.doctypeid = {0}) {1} order by docdate desc;", PURCHASING_CHARTER, toCriteria2.Replace("where", " and ")))

        sb.Append(String.Format("select distinct d.docdate::date,sa.auditby,sa.audittype,sa.auditgrade,sa.overallauditresult,sa.score,doc.getzetolname(d.id) as zetol,d.remarks, d.id,d.docname,d.docext from doc.vendordoc vd left join doc.document d on d.id = vd.documentid left join doc.socialaudit sa on sa.documentid = d.id" &
                                " left join vendor v on v.vendorcode = vd.vendorcode where d.doctypeid = {0} {1} order by docdate desc;", SOCIAL_AUDIT, toCriteria2.Replace("where", " and ")))

        'Factory&Contact 31, 32, 33

        'sb.Append(String.Format(" select v.vendorcode,v.vendorname::text,v.shortname::text,ssm.officersebname::text as ssm,v.ssmidpl as ssmid,v.ssmeffectivedate,pm.officersebname::text as pm,v.pmid as pmid,v.pmeffectivedate,o.officername::text,v.shortname2,true as status" &
        '          " from vendor v" &
        '          " left join officerseb ssm on ssm.ofsebid = v.ssmidpl" &
        '          " left join officerseb pm on pm.ofsebid = v.pmid" &
        '          " left join officer o on o.officerid = v.officerid {0}" &
        '          " order by v.vendorcode;", toCriteria2))

        'viewvendorpmeffectivedate replace viewvendorfamilypmeffectivedate
        sb.Append(String.Format(" select v.vendorcode,v.vendorname::text,v.shortname::text,spm.username::text as ssm,v.ssmidpl as ssmid,ve.spmeffectivedate as ssmeffectivedate," &
                                " pm.username::text as pm,v.pmid as pmid,ve.pmeffectivedate,gsm.username::text as gsm,vgsm.effectivedate ,v.shortname2,true as status " &
                                " from vendor v " &
                                " left join doc.viewvendorpmeffectivedate ve on ve.vendorcode  = v.vendorcode" &
                                " LEFT JOIN officerseb os ON os.ofsebid = ve.pmid " &
                                " LEFT JOIN masteruser pm ON pm.id = os.muid " &
                                " LEFT JOIN officerseb o ON o.ofsebid = os.parent " &
                                " LEFT JOIN masteruser spm ON spm.id = o.muid " &
                                " LEFT JOIN doc.vendorgsm vgsm ON vgsm.vendorcode = v.vendorcode  " &
                                " LEFT JOIN officerseb o1 ON o1.ofsebid = vgsm.gsmid " &
                                " LEFT JOIN masteruser gsm ON gsm.id = o1.muid" &
                                " {0} order by v.vendorcode;", toCriteria2))
        'sb.Append(String.Format("with data as (select 'FP'::text as pt,vf.vendorcode,vf.familyid,fpm.pmid,pmeffectivedate,spmeffectivedate from doc.vendorfamily vf" &
        '                        " left join doc.familypm fpm on fpm.familyid = vf.familyid" &
        '                        " union all select 'CP'::text as pt,vendorcode,null::integer as familyid,pmid,pmeffectivedate,spmeffectivedate from doc.vendorpm vp)," &
        '                        " sbu as (select  distinct familylv1,sbu,sbuname from materialmaster mm" &
        '                        " left join sbusap ss on ss.sbuid = mm.sbu" &
        '                        " where not familylv1 isnull order by familylv1) " &
        '                        " select  data.pt,data.familyid  || ' ' || f.familyname as familyname ,sbu.sbuname,v.vendorcode,v.vendorname::text,v.shortname::text,spm.username::text as ssm,v.ssmidpl as ssmid,v.ssmeffectivedate, pm.username::text as pm," &
        '                        " v.pmid as pmid,v.pmeffectivedate,gsm.username::text as gsm,vgsm.effectivedate ,v.shortname2,true as status  from data" &
        '                        " left join vendor v on v.vendorcode = data.vendorcode" &
        '                        " left join sbu on sbu.familylv1 = data.familyid" &
        '                        " left join family f on f.familyid = data.familyid" &
        '                        " LEFT JOIN officerseb os ON os.ofsebid = data.pmid  LEFT JOIN masteruser pm ON pm.id = os.muid  " &
        '                        " LEFT JOIN officerseb o ON o.ofsebid = os.parent  LEFT JOIN masteruser spm ON spm.id = o.muid  " &
        '                        " LEFT JOIN doc.vendorgsm vgsm ON vgsm.vendorcode = v.vendorcode   LEFT JOIN officerseb o1 ON o1.ofsebid = vgsm.gsmid  " &
        '                        " LEFT JOIN masteruser gsm ON gsm.id = o1.muid {0} order by vendorcode,pt desc,familyname;", toCriteria2))
        sb.Append(String.Format("select f.chinesename,f.englishname,f.englishaddress,f.chineseaddress,f.area,f.city ,p.paramname as province, c.paramname as country, f.id,f.factoryhdid,f.provinceid,f.countryid,f.modifiedby,v.vendorcode,fhd.customname from doc.vendorfactory vf" &
                                " left join vendor v on v.vendorcode = vf.vendorcode" &
                                " left join doc.factorydtl f on f.id = vf.factoryid" &
                                " left join doc.factoryhd fhd on fhd.id = f.factoryhdid" &
                                " left join doc.paramdt c on c.paramdtid = f.countryid" &
                                " left join doc.paramdt p on p.paramdtid = f.provinceid" &
                                "{0} order by f.englishname;", toCriteria2))

        sb.Append(String.Format("select v.vendorcode,c.* from doc.vendorcontact vc" &
                                " left join vendor v on v.vendorcode = vc.vendorcode" &
                                " left join doc.contact c on c.id = vc.contactid" &
                                " {0} order by c.contactname;", toCriteria2))

        'SIF & Identity 34,35,36
        sb.Append(String.Format("with myd as (select distinct d.docdate,d.id,d.doctypeid,dt.doctypename  from doc.vendordoc vd" &
                                " left join doc.document d on d.id = vd.documentid" &
                                " left join vendor v on v.vendorcode = vd.vendorcode " &
                                " left join doc.doctype dt on dt.id = d.doctypeid " &
                                " where(d.doctypeid = {1} Or d.doctypeid = {2})" &
                                " {0}" &
                                " order by docdate desc limit 1)" &
                                " select sl.name,sx.value,myd.docdate,doctypeid,doctypename from myd " &
                                " left join  doc.sitx sx on sx.documentid = myd.id " &
                                " left join doc.silabel sl on sl.id = sx.labelid " &
                                " where iscommon" &
                                " order by orderline;", toCriteria2.Replace("where", " and "), SIF, IDENTITY_SHEET))
        sb.Append(String.Format("with myd as (select distinct d.docdate,d.id,d.doctypeid,v.vendorcode  from doc.vendordoc vd " &
                                " left join doc.document d on d.id = vd.documentid " &
                                " left join vendor v on v.vendorcode = vd.vendorcode " &
                                " where(d.doctypeid = {1})" &
                                " {0}" &
                                " order by docdate desc limit 1)" &
                                " select sl.name,sx.value,myd.docdate from myd" &
                                " left join  doc.sitx sx on sx.documentid = myd.id " &
                                " left join doc.silabel sl on sl.id = sx.labelid " &
                                " where not iscommon" &
                                " order by orderline;", toCriteria2.Replace("where", " and "), SIF))

        sb.Append(String.Format("with myd as (select distinct d.docdate,d.id,d.doctypeid,v.vendorcode  from doc.vendordoc vd " &
                                " left join doc.document d on d.id = vd.documentid " &
                                " left join vendor v on v.vendorcode = vd.vendorcode " &
                                " where(d.doctypeid = {1})" &
                                " {0}" &
                                " order by docdate desc limit 1)" &
                                " select sl.name,sx.value,myd.docdate from myd" &
                                " left join  doc.sitx sx on sx.documentid = myd.id " &
                                " left join doc.silabel sl on sl.id = sx.labelid " &
                                " where not iscommon" &
                                " order by orderline;", toCriteria2.Replace("where", " and "), IDENTITY_SHEET))
        '***************** Purchasing Data Base on fpcp
        'PriceIndex
        'YearToDate
        'sb.Append(String.Format("select d.year,sum(qty) as qty,sum(amount) as amount,sum(tovariance) as tovariance,sum(towavpymin1) as towavpymin1,sum(towlkpmin1) as towlkpmin1,sum(tovariance) as tovariance, sum(amount)/sum(towavpymin1) * 100 as piavg,sum(amount)/sum(towlkpmin1) * 100 as pilkp,sum(amount) / sum(towstd) * 100 as pistd" &
        '              " from doc.turnover d " &
        '              " left join vendor v on v.vendorcode = d.vendorcode {0}" &
        '              " group by year;", tocriteriafpcp))
        sb.Append(String.Format("select sum(amount)/sum(towavpymin1) * 100 as piavg,sum(amount)/sum(towlkpmin1) * 100 as pilkp," &
                                " sum(originalamountwavgymin1fxwoamort)/sum(towaverpricey1fixedcurr)*100 as piavgfx" &
                      " from doc.turnover d " &
                      " left join vendor v on v.vendorcode = d.vendorcode" &
                       " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                      " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                      " {0} {1} {2}" &
                      " group by year;", tocriteriafpcp, ProductTypeCriteria, StatusFilter))

        sb.Append(String.Format("select sum(tovariance) as tovariance,sum(amount) / sum(towstd) * 100 as pistd," &
                        " sum(originalamountwstdfxwoamort)/sum(towstd) * 100 as pistdfx,sum(tovariancewstdfx) as tovariancefx" &
                      " from doc.turnover d " &
                      " left join vendor v on v.vendorcode = d.vendorcode" &
                       " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
                      " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                      " {0} {1}" &
                      " group by year;", tocriteriafp, StatusFilter))

        'Y-1 to Y-4
        For i = 1 To 4
            sb.Append(String.Format("select sum(amount)/sum(towavpymin1) * 100 as piavg,sum(amount)/sum(towlkpmin1) * 100 as pilkp, " &
                                    " sum(originalamountwavgymin1fxwoamort)/sum(towaverpricey1fixedcurr)*100 as piavgfx" &
                      " from doc.turnover d " &
                      " left join vendor v on v.vendorcode = d.vendorcode" &
                       " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
                      " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                      " {0} and year = {1} {2} {3}" &
                      " group by year;", toCriteria2, Year(_currentdate) - i, "  and (fpcp in ('FP','CP') )  ", ProductTypeCriteria, StatusFilter))
        Next


        'Y-1 to Y-4
        For i = 1 To 4
            sb.Append(String.Format("select sum(tovariance) as tovariance,sum(amount) / sum(towstd) * 100 as pistd, " &
                                    " sum(originalamountwstdfxwoamort)/sum(towstd) * 100 as pistdfx,sum(tovariancewstdfx) as tovariancefx" &
                      " from doc.turnover d " &
                      " left join vendor v on v.vendorcode = d.vendorcode" &
                      " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
                      " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                      " {0} and year = {1} {2} {3} " &
                      " group by year;", toCriteria2, Year(_currentdate) - i, "  and (fpcp in ('FP') )  ", StatusFilter))
        Next

        'ContractTerm
        sb.Append(String.Format("with data as (select distinct v.shortname,first_value(vd.documentid) over (partition by v.shortname,doctypeid order by docdate desc) as id," &
                  " doctypeid from doc.vendordoc vd left join doc.document d on d.id = vd.documentid  left join doc.doctype dt on dt.id = d.doctypeid " &
                  " left join vendor v on v.vendorcode = vd.vendorcode  where doctypeid = {1} order by shortname) select go.*,p.payt || ' - ' || p.details as paymentterm from data " &
                  " left join doc.document d on d.id = data.id left join doc.generalcontract gc on gc.documentid = d.id  " &
                  " left join paymentterm p on p.paymenttermid = gc.paymentcode left join doc.qualityappendix qa on qa.documentid = d.id  " &
                  " inner join doc.generalcontractother go on go.documentid = data.id left join paymentterm pt on pt.paymenttermid = go.paymentcode" &
                  " where shortname = '{0}';", UcSupplierInfo1.Shortname, CONTRACT_GENERAL_CONTRACT))
        sb.Append(String.Format("with data as (select distinct v.shortname,first_value(vd.documentid) over (partition by v.shortname,doctypeid order by docdate desc) as id," &
          " doctypeid from doc.vendordoc vd left join doc.document d on d.id = vd.documentid  left join doc.doctype dt on dt.id = d.doctypeid " &
          " left join vendor v on v.vendorcode = vd.vendorcode  where doctypeid = {1} order by shortname) select qo.* from data " &
          " left join doc.document d on d.id = data.id left join doc.generalcontract gc on gc.documentid = d.id  " &
          " left join paymentterm p on p.paymenttermid = gc.paymentcode left join doc.qualityappendix qa on qa.documentid = d.id  " &
          " inner join doc.qualityappendixother qo on qo.documentid = data.id " &
          " where shortname = '{0}';", UcSupplierInfo1.Shortname, CONTRACT_QUALITY_APPENDIX))

        sb.Append(String.Format("with data as (select distinct v.shortname,first_value(vd.documentid) over (partition by v.shortname,doctypeid order by docdate desc) as id," &
          " doctypeid from doc.vendordoc vd left join doc.document d on d.id = vd.documentid  left join doc.doctype dt on dt.id = d.doctypeid " &
          " left join vendor v on v.vendorcode = vd.vendorcode  where doctypeid = {1} order by shortname) select sc.*,'Lead Time' as leadtimehd,'SASL' as saslhd from data " &
          " left join doc.document d on d.id = data.id left join doc.generalcontract gc on gc.documentid = d.id  " &
          " left join paymentterm p on p.paymenttermid = gc.paymentcode " &
          " inner join doc.supplychainother sc on sc.documentid = d.id left join doc.docexpired de on de.documentid = d.id  " &
          " where shortname = '{0}';", UcSupplierInfo1.Shortname, CONTRACT_SUPPLY_CHAIN_APPENDIX))

        'PanelHistory



        'CMMF
        If SelectedTab Then
            'sb.Append(String.Format("with  tuinfo as (	select distinct tu.cmmf,c.commercialref,tu.sbu,tu.comfam,range,f.familyname, tu.brand from doc.turnover tu	left join cmmf c on c.cmmf = tu.cmmf	left join family f on f.familyid = c.comfam	left join pricelist pl on pl.cmmf = tu.cmmf	left join vendor v on v.vendorcode = pl.vendorcode	{1} and  not tu.cmmf isnull )," &
            '    " pl as (" &
            '       " select foo2.cmmf,tui.commercialref,tui.sbu,tui.comfam,tui.range,tui.familyname,tui.brand from(" &
            '           " select distinct cmmf from ((select distinct pl.cmmf from pricelist pl left join vendor v on v.vendorcode = pl.vendorcode {1}  group by pl.cmmf order by cmmf) union all (select distinct cmmf from doc.turnover tu " &
            '           " left join vendor v on v.vendorcode = tu.vendorcode {1}  and not cmmf isnull)) as foo order by cmmf" &
            '           ") as foo2 left join tuinfo tui on tui.cmmf = foo2.cmmf" &
            '    " )," &
            '    " fob as (select distinct * from (select cmmf ,{0} as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
            '    " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit" &
            '    " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} ) as foo)," &
            '    " f_1 as (select distinct * from (select cmmf ,{0} - 1 as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
            '    " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit" &
            '    " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} - 1 ) as foo)," &
            '    " f_2 as (select distinct * from (select cmmf ,{0} - 2 as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
            '    " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit" &
            '    " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} - 2 ) as foo)," &
            '    " f_3 as (select distinct * from (select cmmf ,{0} - 3 as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
            '    " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit" &
            '    " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} - 3 ) as foo)," &
            '    " f_4 as (select distinct * from (select cmmf ,{0} - 4 as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
            '    " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit" &
            '    " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} - 4 ) as foo)," &
            '    " lkp as(select cmmf,cp.vendorcode,myyear as year,lastprice from cmmfvendorprice cp left join vendor v on v.vendorcode = cp.vendorcode {1})," &
            '    " pld as (select distinct pl.cmmf,pl.vendorcode,date_part('Year',validfrom) as year from pricelist pl left join vendor v on v.vendorcode = pl.vendorcode {1})," &
            '    " std as (select distinct * from (select std.cmmf,date_part('Year',std.validfrom) as year,first_value(planprice1) over(partition by std.cmmf,date_part('Year',std.validfrom) order by std.validfrom desc )::numeric as planprice1," &
            '    " first_value(per) over(partition by std.cmmf,std.validfrom order by std.validfrom desc )::numeric as per	from standardcostad std	left join pld on pld.cmmf = std.cmmf and pld.year = date_part('Year',std.validfrom)	left join vendor v on v.vendorcode = pld.vendorcode" &
            '    " {1}) as foo )," &
            '    " ac as (select distinct cmmf from pricelist)," &
            '    " tu as (select cmmf,year,sum(qty) as sumqty,sum(amount) as sumamount from doc.turnover tu where tu.year = {0} 	group by cmmf,year)" &
            '    " select distinct pl.cmmf,mm.materialdesc::text,tu.sumqty,tu.sumamount,f.amount / f.perunit as amountperunit ,f_1.amount / f_1.perunit as amountperunit_1 ,f_2.amount / f_2.perunit as amountperunit_2 ,f_3.amount / f_3.perunit as amountperunit_3 ,f_4.amount / f_4.perunit as amountperunit_4 ," &
            '    " lkp.lastprice::double precision as lkp,lkp_1.lastprice::double precision as lkp_1,lkp_2.lastprice::double precision as lkp_2,lkp_3.lastprice::double precision as lkp_3,lkp_4.lastprice::double precision as lkp_4," &
            '    " std.planprice1::double precision / std.per::double precision as plantpriceperunit,std_1.planprice1::double precision / std_1.per::double precision as plantpriceperunit_1,std_2.planprice1::double precision / std_2.per::double precision as plantpriceperunit_2,std_3.planprice1::double precision / std_3.per::double precision as plantpriceperunit_3,std_4.planprice1::double precision / std_4.per::double precision as plantpriceperunit_4" &
            '    " ,pl.commercialref::text from pl" &
            '    " left join tu on tu.cmmf = pl.cmmf left join tuinfo tui on tui.cmmf = pl.cmmf left join materialmaster mm on mm.cmmf = pl.cmmf left join fob f on f.cmmf = pl.cmmf and f.year = {0} left join f_1 on f_1.cmmf = pl.cmmf and f_1.year = {0} - 1 left join f_2 on f_2.cmmf = pl.cmmf and f_2.year = {0} - 2 left join f_3 on f_3.cmmf = pl.cmmf and f_3.year = {0} - 3 left join f_4 on f_4.cmmf = pl.cmmf and f_4.year = {0} - 4 " &
            '    " left join lkp on lkp.cmmf = pl.cmmf and lkp.year = {0} left join lkp lkp_1 on lkp_1.cmmf = pl.cmmf and lkp_1.year = {0} - 1 left join lkp lkp_2 on lkp_2.cmmf = pl.cmmf and lkp_2.year = {0} - 2 left join lkp lkp_3 on lkp_3.cmmf = pl.cmmf and lkp_3.year = {0} - 3 left join lkp lkp_4 on lkp_4.cmmf = pl.cmmf and lkp_4.year = {0} - 4 left join std on std.cmmf = pl.cmmf and std.year = {0} left join std std_1 on std_1.cmmf = pl.cmmf and std_1.year = {0} - 1" &
            '    " left join std std_2 on std_2.cmmf = pl.cmmf and std_2.year = {0} - 2 left join std std_3 on std_3.cmmf = pl.cmmf and std_3.year = {0} - 3 left join std std_4 on std_4.cmmf = pl.cmmf and std_4.year = {0} - 4 {2} order by pl.cmmf ", Year(_currentdate), toCriteria2, " where true " & TextFilterForPriceCMMF.Replace("''", "'").Replace("tu.", "pl.")))

            'sb.Append(String.Format("with  tuinfo as (	select distinct tu.cmmf,c.commercialref,tu.sbu,tu.comfam,range,f.familyname, tu.brand ,fpcp from doc.turnover tu	left join cmmf c on c.cmmf = tu.cmmf	left join family f on f.familyid = c.comfam	left join pricelist pl on pl.cmmf = tu.cmmf	left join vendor v on v.vendorcode = pl.vendorcode	{1} and  not tu.cmmf isnull )," &
            '    " pl as (" &
            '       " select foo2.cmmf,tui.commercialref,tui.sbu,tui.comfam,tui.range,tui.familyname,tui.brand from(" &
            '           " select distinct cmmf from ((select distinct pl.cmmf from pricelist pl left join vendor v on v.vendorcode = pl.vendorcode {1}  group by pl.cmmf order by cmmf) union all (select distinct cmmf from doc.turnover tu " &
            '           " left join vendor v on v.vendorcode = tu.vendorcode {1}  and not cmmf isnull)) as foo order by cmmf" &
            '           ") as foo2 left join tuinfo tui on tui.cmmf = foo2.cmmf" &
            '    " )," &
            '    " fob as (select distinct * from (select cmmf ,{0} as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
            '    " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit" &
            '    " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} ) as foo)," &
            '    " f_1 as (select distinct * from (select cmmf ,{0} - 1 as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
            '    " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit" &
            '    " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} - 1 ) as foo)," &
            '    " f_2 as (select distinct * from (select cmmf ,{0} - 2 as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
            '    " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit" &
            '    " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} - 2 ) as foo)," &
            '    " f_3 as (select distinct * from (select cmmf ,{0} - 3 as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
            '    " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit" &
            '    " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} - 3 ) as foo)," &
            '    " f_4 as (select distinct * from (select cmmf ,{0} - 4 as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
            '    " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit" &
            '    " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} - 4 ) as foo)," &
            '    " lkp as(select cmmf,cp.vendorcode,myyear as year,lastprice from cmmfvendorprice cp left join vendor v on v.vendorcode = cp.vendorcode {1})," &
            '    " pld as (select distinct pl.cmmf,pl.vendorcode,date_part('Year',validfrom) as year from pricelist pl left join vendor v on v.vendorcode = pl.vendorcode {1})," &
            '    " std as (select distinct * from (select std.cmmf,date_part('Year',std.validfrom) as year,first_value(planprice1) over(partition by std.cmmf,date_part('Year',std.validfrom) order by std.validfrom desc )::numeric as planprice1," &
            '    " first_value(per) over(partition by std.cmmf,std.validfrom order by std.validfrom desc )::numeric as per	from standardcostad std	" &
            '    " ) as foo )," &
            '    " ac as (select distinct cmmf from pricelist)," &
            '    " tu as (select cmmf,year,sum(qty) as sumqty,sum(amount) as sumamount from doc.turnover tu left join vendor v on v.vendorcode = tu.vendorcode {1} and tu.year = {0} group by cmmf,year)" &
            '    " select distinct pl.cmmf,mm.materialdesc::text,tu.sumamount,null::text,f.amount / f.perunit as amountperunit ,f_1.amount / f_1.perunit as amountperunit_1 ,f_2.amount / f_2.perunit as amountperunit_2 ,f_3.amount / f_3.perunit as amountperunit_3 ,f_4.amount / f_4.perunit as amountperunit_4 ," &
            '    " null::text,lkp.lastprice::double precision as lkp,lkp_1.lastprice::double precision as lkp_1,lkp_2.lastprice::double precision as lkp_2,lkp_3.lastprice::double precision as lkp_3,lkp_4.lastprice::double precision as lkp_4," &
            '    " null::text,std.planprice1::double precision / std.per::double precision as plantpriceperunit,std_1.planprice1::double precision / std_1.per::double precision as plantpriceperunit_1,std_2.planprice1::double precision / std_2.per::double precision as plantpriceperunit_2,std_3.planprice1::double precision / std_3.per::double precision as plantpriceperunit_3,std_4.planprice1::double precision / std_4.per::double precision as plantpriceperunit_4" &
            '    " ,pl.commercialref::text,pl.cmmf::text || ' - ' || case when mm.materialdesc isnull then ' '  else mm.materialdesc::text end as cmmfdesc ,tu.sumqty from pl" &
            '    " left join doc.sebplatform sf on sf.cmmf = pl.cmmf" &
            '    " left join pld on pld.cmmf = pl.cmmf" &
            '    " left join vendor v on v.vendorcode = pld.vendorcode" &
            '    " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode	" &
            '    " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
            '    " left join tu on tu.cmmf = pl.cmmf left join tuinfo tui on tui.cmmf = pl.cmmf left join materialmaster mm on mm.cmmf = pl.cmmf left join fob f on f.cmmf = pl.cmmf and f.year = {0} left join f_1 on f_1.cmmf = pl.cmmf and f_1.year = {0} - 1 left join f_2 on f_2.cmmf = pl.cmmf and f_2.year = {0} - 2 left join f_3 on f_3.cmmf = pl.cmmf and f_3.year = {0} - 3 left join f_4 on f_4.cmmf = pl.cmmf and f_4.year = {0} - 4 " &
            '    " left join lkp on lkp.cmmf = pl.cmmf and lkp.year = {0} left join lkp lkp_1 on lkp_1.cmmf = pl.cmmf and lkp_1.year = {0} - 1 left join lkp lkp_2 on lkp_2.cmmf = pl.cmmf and lkp_2.year = {0} - 2 left join lkp lkp_3 on lkp_3.cmmf = pl.cmmf and lkp_3.year = {0} - 3 left join lkp lkp_4 on lkp_4.cmmf = pl.cmmf and lkp_4.year = {0} - 4 left join std on std.cmmf = pl.cmmf and std.year = {0} left join std std_1 on std_1.cmmf = pl.cmmf and std_1.year = {0} - 1" &
            '    " left join std std_2 on std_2.cmmf = pl.cmmf and std_2.year = {0} - 2 left join std std_3 on std_3.cmmf = pl.cmmf and std_3.year = {0} - 3 left join std std_4 on std_4.cmmf = pl.cmmf and std_4.year = {0} - 4 {2} order by pl.cmmf ", Year(_currentdate), toCriteria2, " where true " & TextFilterForPriceCMMF.Replace("''", "'").Replace("tu.", "pl.")))

            sb.Append(String.Format("with  tuinfo as (	select distinct tu.cmmf,c.commercialref,tu.sbu,tu.comfam,range,f.familyname, tu.brand ,fpcp from doc.turnover tu	left join cmmf c on c.cmmf = tu.cmmf	left join family f on f.familyid = c.comfam	left join pricelist pl on pl.cmmf = tu.cmmf	left join vendor v on v.vendorcode = pl.vendorcode	{1} and  not tu.cmmf isnull )," &
                " pl as (" &
                   " select foo2.cmmf,tui.commercialref,tui.sbu,tui.comfam,tui.range,tui.familyname,tui.brand from(" &
                       " select distinct cmmf from ((select distinct pl.cmmf from pricelist pl left join vendor v on v.vendorcode = pl.vendorcode {1}  group by pl.cmmf order by cmmf) union all (select distinct cmmf from doc.turnover tu " &
                       " left join vendor v on v.vendorcode = tu.vendorcode {1}  and not cmmf isnull)) as foo order by cmmf" &
                       ") as foo2 left join tuinfo tui on tui.cmmf = foo2.cmmf" &
                " )," &
                " fob as (select distinct * from (select cmmf ,{0} as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
                " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit, first_value(currency) over (partition by cmmf order by cmmf,validfrom desc) as crcy" &
                " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} ) as foo)," &
                " f_1 as (select distinct * from (select cmmf ,{0} - 1 as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
                " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit, first_value(currency) over (partition by cmmf order by cmmf,validfrom desc) as crcy" &
                " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} - 1 ) as foo)," &
                " f_2 as (select distinct * from (select cmmf ,{0} - 2 as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
                " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit, first_value(currency) over (partition by cmmf order by cmmf,validfrom desc) as crcy" &
                " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} - 2 ) as foo)," &
                " f_3 as (select distinct * from (select cmmf ,{0} - 3 as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
                " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit, first_value(currency) over (partition by cmmf order by cmmf,validfrom desc) as crcy" &
                " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} - 3 ) as foo)," &
                " f_4 as (select distinct * from (select cmmf ,{0} - 4 as year, first_value(amount) over (partition by cmmf order by cmmf,validfrom desc)as amount," &
                " first_value(perunit) over (partition by cmmf order by cmmf,validfrom desc) as perunit, first_value(currency) over (partition by cmmf order by cmmf,validfrom desc) as crcy" &
                " from pricelist pl left join vendor v on v.vendorcode =pl.vendorcode  {1} and date_part('Year',validfrom) <= {0} - 4 ) as foo)," &
                " lkp as(select cmmf,cp.vendorcode,myyear as year,lastprice,doc.getvendorcurr(cp.vendorcode, cp.lasttx) as crcy from cmmfvendorprice cp left join vendor v on v.vendorcode = cp.vendorcode {1})," &
                " pld as (select distinct pl.cmmf,pl.vendorcode,date_part('Year',validfrom) as year from pricelist pl left join vendor v on v.vendorcode = pl.vendorcode {1})," &
                " std as (select distinct * from (select std.cmmf,date_part('Year',std.validfrom) as year,first_value(planprice1) over(partition by std.cmmf,date_part('Year',std.validfrom) order by std.validfrom desc )::numeric as planprice1," &
                " first_value(per) over(partition by std.cmmf,std.validfrom order by std.validfrom desc )::numeric as per	from standardcostad std	" &
                " ) as foo )," &
                " ac as (select distinct cmmf from pricelist)," &
                " tu as (select cmmf,year,sum(qty) as sumqty,sum(amount) as sumamount from doc.turnover tu left join vendor v on v.vendorcode = tu.vendorcode {1} and tu.year = {0} group by cmmf,year)" &
                " select distinct pl.cmmf,mm.materialdesc::text,tu.sumamount,null::text,f.amount / f.perunit as amountperunit,f.crcy as fcrcy ,f_1.amount / f_1.perunit as amountperunit_1 ,f_1.crcy as f_1crcy,f_2.amount / f_2.perunit as amountperunit_2 ,f_2.crcy as f_2crcy,f_3.amount / f_3.perunit as amountperunit_3 ,f_3.crcy as f_3crcy,f_4.amount / f_4.perunit as amountperunit_4 ,f_4.crcy as f_4crcy," &
                " null::text,lkp.lastprice::double precision as lkp,lkp.crcy as lkpcrcy,lkp_1.lastprice::double precision as lkp_1,lkp_1.crcy as lkp_1crcy,lkp_2.lastprice::double precision as lkp_2,lkp_2.crcy as lkp_2crcy,lkp_3.lastprice::double precision as lkp_3,lkp_3.crcy as lkp_3crcy,lkp_4.lastprice::double precision as lkp_4,lkp_4.crcy as lkp_4crcy," &
                " null::text,std.planprice1::double precision / std.per::double precision as plantpriceperunit,std_1.planprice1::double precision / std_1.per::double precision as plantpriceperunit_1,std_2.planprice1::double precision / std_2.per::double precision as plantpriceperunit_2,std_3.planprice1::double precision / std_3.per::double precision as plantpriceperunit_3,std_4.planprice1::double precision / std_4.per::double precision as plantpriceperunit_4" &
                " ,pl.commercialref::text,pl.cmmf::text || ' - ' || case when mm.materialdesc isnull then ' '  else mm.materialdesc::text end as cmmfdesc ,tu.sumqty from pl" &
                " left join doc.sebplatform sf on sf.cmmf = pl.cmmf" &
                " left join pld on pld.cmmf = pl.cmmf" &
                " left join vendor v on v.vendorcode = pld.vendorcode" &
                " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode	" &
                " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                " left join tu on tu.cmmf = pl.cmmf left join tuinfo tui on tui.cmmf = pl.cmmf left join materialmaster mm on mm.cmmf = pl.cmmf left join fob f on f.cmmf = pl.cmmf and f.year = {0} left join f_1 on f_1.cmmf = pl.cmmf and f_1.year = {0} - 1 left join f_2 on f_2.cmmf = pl.cmmf and f_2.year = {0} - 2 left join f_3 on f_3.cmmf = pl.cmmf and f_3.year = {0} - 3 left join f_4 on f_4.cmmf = pl.cmmf and f_4.year = {0} - 4 " &
                " left join lkp on lkp.cmmf = pl.cmmf and lkp.year = {0} left join lkp lkp_1 on lkp_1.cmmf = pl.cmmf and lkp_1.year = {0} - 1 left join lkp lkp_2 on lkp_2.cmmf = pl.cmmf and lkp_2.year = {0} - 2 left join lkp lkp_3 on lkp_3.cmmf = pl.cmmf and lkp_3.year = {0} - 3 left join lkp lkp_4 on lkp_4.cmmf = pl.cmmf and lkp_4.year = {0} - 4 left join std on std.cmmf = pl.cmmf and std.year = {0} left join std std_1 on std_1.cmmf = pl.cmmf and std_1.year = {0} - 1" &
                " left join std std_2 on std_2.cmmf = pl.cmmf and std_2.year = {0} - 2 left join std std_3 on std_3.cmmf = pl.cmmf and std_3.year = {0} - 3 left join std std_4 on std_4.cmmf = pl.cmmf and std_4.year = {0} - 4 {2} order by pl.cmmf; ", Year(_currentdate), toCriteria2, " where true " & TextFilterForPriceCMMF.Replace("''", "'").Replace("tu.", "pl.")))
        Else
            sb.Append("select 1;") 'To reserve last table for cmmf
        End If

        'Vendor Currency
        sb.Append(String.Format("select vc.* ,v.shortname from doc.vendorcurr vc left join vendor v on v.vendorcode = vc.vendorcode {0}  order by effectivedate desc ;", toCriteria2))
        DS = New DataSet
        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            DS.CaseSensitive = True
            ProgressReport(1, "Processing.. Please wait. Fill Dataset..")
            ProgressReport(4, "Fill SIF")
        Else
            MessageBox.Show(mymessage)
        End If

        bFAdapter = New BudgetForecastAdapter(_vendorcode, _shortname)
        If bFAdapter.loaddata(BFCriteria) Then
            ProgressReport(9, "Fill Forecast budget")
        Else
            ProgressReport(9, "Loading Forecast Error.")
        End If

        ProgressReport(6, "Continuous")
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

                    Case 4
                        TurnoverBS = New BindingSource
                        SEBTurnoverBS = New BindingSource
                        SEBTurnoverBS1 = New BindingSource
                        SEBTurnoverBS2 = New BindingSource
                        SEBTurnoverBS3 = New BindingSource
                        SEBTurnoverBS4 = New BindingSource
                        FilterQtyBS = New BindingSource
                        FilterAmountBS = New BindingSource
                        GroupFPCPBS = New BindingSource
                        GroupSBUBS = New BindingSource
                        GroupFamilyBS = New BindingSource
                        GroupBrandBS = New BindingSource
                        NQSUBS = New BindingSource
                        LogisticsBS = New BindingSource
                        LogisticsBS1 = New BindingSource
                        LogisticsBS2 = New BindingSource
                        LogisticsBS3 = New BindingSource
                        LogisticsBS4 = New BindingSource
                        PDBS = New BindingSource
                        PDBS1 = New BindingSource
                        PDBS2 = New BindingSource
                        PDBS3 = New BindingSource
                        PDBS4 = New BindingSource
                        SAPBS = New BindingSource
                        ContractBS = New BindingSource

                        ALetterBS = New BindingSource
                        PSpecBS = New BindingSource
                        CitiProgramBS = New BindingSource
                        ToolingHeader = New BindingSource
                        ToolingDetailsAll = New BindingSource
                        ToolingDetails = New BindingSource
                        PCharterBS = New BindingSource
                        SocialAuditBS = New BindingSource

                        VendorFactoryContactBS = New BindingSource
                        FactoryBS = New BindingSource
                        ContactBS = New BindingSource
                        SIFIDBS = New BindingSource
                        SIFBS = New BindingSource
                        IDBS = New BindingSource

                        ContractualTerms1BS = New BindingSource
                        ContractualTerms2BS = New BindingSource
                        ContractualTerms3BS = New BindingSource

                        PriceCMMFBS = New BindingSource
                        VendorCurrencyBS = New BindingSource

                        TurnoverBS.DataSource = DS.Tables(0)
                        SEBTurnoverBS.DataSource = DS.Tables(1)
                        SEBTurnoverBS1.DataSource = DS.Tables(2)
                        SEBTurnoverBS2.DataSource = DS.Tables(3)
                        SEBTurnoverBS3.DataSource = DS.Tables(4)
                        SEBTurnoverBS4.DataSource = DS.Tables(5)
                        FilterQtyBS.DataSource = DS.Tables(6)
                        FilterAmountBS.DataSource = DS.Tables(7)
                        GroupFPCPBS.DataSource = DS.Tables(8)
                        GroupSBUBS.DataSource = DS.Tables(9)
                        GroupFamilyBS.DataSource = DS.Tables(10)
                        GroupBrandBS.DataSource = DS.Tables(11)
                        NQSUBS.DataSource = DS.Tables(12)
                        LogisticsBS.DataSource = DS.Tables(13)
                        LogisticsBS1.DataSource = DS.Tables(14)
                        LogisticsBS2.DataSource = DS.Tables(15)
                        LogisticsBS3.DataSource = DS.Tables(16)
                        LogisticsBS4.DataSource = DS.Tables(17)
                        PDBS.DataSource = DS.Tables(18)
                        PDBS1.DataSource = DS.Tables(19)
                        PDBS2.DataSource = DS.Tables(20)
                        PDBS3.DataSource = DS.Tables(21)
                        PDBS4.DataSource = DS.Tables(22)
                        SAPBS.DataSource = DS.Tables(23)
                        ContractBS.DataSource = DS.Tables(24)

                        Dim MyView As DataView = DS.Tables(25).AsDataView




                        MyView.RowFilter = "doctypeid = 36"
                        MyView.Sort = "docdate desc,projectname"
                        Dim ALetter = MyView.ToTable("ALetter")


                        MyView.RowFilter = "doctypeid = 18"
                        MyView.Sort = "docdate desc,projectname"
                        Dim PSpec = MyView.ToTable("PSpec")


                        VendorCurrencyBS.DataSource = DS.Tables(51)


                        UcContractualTermDGV1.SAPBS = SAPBS
                        UcContractualTermDGV1.ContractBS = ContractBS
                        ALetterBS.DataSource = ALetter
                        PSpecBS.DataSource = PSpec
                        CitiProgramBS.DataSource = DS.Tables(26)
                        UcContractualTermDGV1.AuthLeter = ALetterBS
                        UcContractualTermDGV1.ProjectSpec = PSpecBS
                        UcContractualTermDGV1.CitiProgram = CitiProgramBS
                        UcContractualTermDGV1.VendorCurrency = VendorCurrencyBS


                        Dim myrel As DataRelation
                        Dim toolingHDCol As DataColumn
                        Dim ToolingDTCol As DataColumn
                        toolingHDCol = DS.Tables(27).Columns("project")
                        ToolingDTCol = DS.Tables(28).Columns("project")
                        myrel = New DataRelation("ToolingRel", toolingHDCol, ToolingDTCol)
                        DS.Relations.Add(myrel)

                        ToolingHeader.DataSource = DS.Tables(27)
                        ToolingDetailsAll.DataSource = DS.Tables(28)
                        ToolingDetails.DataSource = ToolingHeader
                        ToolingDetails.DataMember = "ToolingRel"


                        PCharterBS.DataSource = DS.Tables(29)
                        SocialAuditBS.DataSource = DS.Tables(30)
                        


                        Dim VendorCol As DataColumn
                        Dim FactoryCol As DataColumn
                        Dim ContactCol As DataColumn

                        VendorCol = DS.Tables(31).Columns("vendorcode")
                        FactoryCol = DS.Tables(32).Columns("vendorcode")
                        myrel = New DataRelation("VendorFactoryRel", VendorCol, FactoryCol)
                        DS.Relations.Add(myrel)

                        ContactCol = DS.Tables(33).Columns("vendorcode")
                        myrel = New DataRelation("VendorContactRel", VendorCol, ContactCol)
                        DS.Relations.Add(myrel)

                        VendorFactoryContactBS.DataSource = DS.Tables(31)
                        FactoryBS.DataSource = VendorFactoryContactBS
                        FactoryBS.DataMember = "VendorFactoryRel"

                        ContactBS.DataSource = VendorFactoryContactBS
                        ContactBS.DataMember = "VendorContactRel"


                        SIFIDBS.DataSource = DS.Tables(34)
                        SIFBS.DataSource = DS.Tables(35)
                        IDBS.DataSource = DS.Tables(36)

                        ContractualTerms1BS.DataSource = DS.Tables(47)
                        ContractualTerms2BS.DataSource = DS.Tables(48)
                        ContractualTerms3BS.DataSource = DS.Tables(49)




                        If SelectedTab Then
                            Try
                                'PriceCMMFBS.DataSource = DS.Tables(47)
                                PriceCMMFBS.DataSource = DS.Tables(50)
                            Catch ex As Exception

                            End Try

                            
                        Else
                            PriceCMMFBS.DataSource = Nothing
                        End If


                        UcSustainableDevelopment1.CharterBS = PCharterBS
                        UcSustainableDevelopment1.SocialAuditBS = SocialAuditBS
                        UcSustainableDevelopment1.DisplayDataGrid()




                        Dim drv As DataRowView = TurnoverBS.Current
                        UcTooling1.ToolingHD = ToolingHeader
                        UcTooling1.ToolingDT = ToolingDetails
                        UcTooling1.DisplayDataGrid()

                        'Dim myproducttype As String = UcSupplierDashboard1.ProductTypeFilter
                        'If myproducttype = "" Then
                        '    myproducttype = "ALL"
                        'End If
                        UctoHeader1.ProductType = UcSupplierDashboard1.ProductTypeFilter
                        'UctoHeader1.ProductType = myproducttype
                        UcScorecardHeader1.ProductType = UcSupplierDashboard1.ProductTypeFilter
                        'UcScorecardHeader1.ProductType = myproducttype
                        ' If Not IsNothing(TurnoverBS.Current) Then
                        ClassSIFTO1.currentDate = _currentdate
                        ClassSIFTO1.SIFTOBS = TurnoverBS
                        ClassSIFTO1.SEBTOBS = SEBTurnoverBS
                        ClassSIFTO1.SEBTOBS1 = SEBTurnoverBS1
                        ClassSIFTO1.SEBTOBS2 = SEBTurnoverBS2
                        ClassSIFTO1.SEBTOBS3 = SEBTurnoverBS3
                        ClassSIFTO1.SEBTOBS4 = SEBTurnoverBS4
                        ClassSIFTO1.DisplayValue()
                        'Else
                        'ClassSIFTO1.ClearValue()
                        'ClassSEBTO1.ClearValue()
                        '' UcContractualTermDGV1.ClearValue()
                        'End If
                        ClassSEBTO1.currentDate = _currentdate
                        ClassSEBTO1.SEBTOBS = SEBTurnoverBS
                        ClassSEBTO1.SEBTOBS1 = SEBTurnoverBS1
                        ClassSEBTO1.SEBTOBS2 = SEBTurnoverBS2
                        ClassSEBTO1.SEBTOBS3 = SEBTurnoverBS3
                        ClassSEBTO1.SEBTOBS4 = SEBTurnoverBS4
                        ClassSEBTO1.DisplayValue()

                        PurchasingFPCP = New BindingSource
                        PurchasingFPCP1 = New BindingSource
                        PurchasingFPCP2 = New BindingSource
                        PurchasingFPCP3 = New BindingSource
                        PurchasingFPCP4 = New BindingSource
                        PurchasingFP = New BindingSource
                        PurchasingFP1 = New BindingSource
                        PurchasingFP2 = New BindingSource
                        PurchasingFP3 = New BindingSource
                        PurchasingFP4 = New BindingSource

                        PurchasingFPCP.DataSource = DS.Tables(37)
                        PurchasingFPCP1.DataSource = DS.Tables(39)
                        PurchasingFPCP2.DataSource = DS.Tables(40)
                        PurchasingFPCP3.DataSource = DS.Tables(41)
                        PurchasingFPCP4.DataSource = DS.Tables(42)

                        PurchasingFP.DataSource = DS.Tables(38)
                        PurchasingFP1.DataSource = DS.Tables(43)
                        PurchasingFP2.DataSource = DS.Tables(44)
                        PurchasingFP3.DataSource = DS.Tables(45)
                        PurchasingFP4.DataSource = DS.Tables(46)

                        'UcScorecardDetail1.SEBTOBS = SEBTurnoverBS
                        'UcScorecardDetail1.SEBTOBS1 = SEBTurnoverBS1
                        'UcScorecardDetail1.SEBTOBS2 = SEBTurnoverBS2
                        'UcScorecardDetail1.SEBTOBS3 = SEBTurnoverBS3
                        'UcScorecardDetail1.SEBTOBS4 = SEBTurnoverBS4

                        UcScorecardDetail1.SEBTOBS = PurchasingFPCP
                        UcScorecardDetail1.SEBTOBS1 = PurchasingFPCP1
                        UcScorecardDetail1.SEBTOBS2 = PurchasingFPCP2
                        UcScorecardDetail1.SEBTOBS3 = PurchasingFPCP3
                        UcScorecardDetail1.SEBTOBS4 = PurchasingFPCP4


                        UcScorecardDetail1.SEBTOBSFP = PurchasingFP
                        UcScorecardDetail1.SEBTOBSFP1 = PurchasingFP1
                        UcScorecardDetail1.SEBTOBSFP2 = PurchasingFP2
                        UcScorecardDetail1.SEBTOBSFP3 = PurchasingFP3
                        UcScorecardDetail1.SEBTOBSFP4 = PurchasingFP4

                        UcScorecardDetail1.NQSUBS = NQSUBS
                        UcScorecardDetail1.LogisticsBS = LogisticsBS
                        UcScorecardDetail1.LogisticsBS1 = LogisticsBS1
                        UcScorecardDetail1.LogisticsBS2 = LogisticsBS2
                        UcScorecardDetail1.LogisticsBS3 = LogisticsBS3
                        UcScorecardDetail1.LogisticsBS4 = LogisticsBS4
                        UcScorecardDetail1.PDBS = PDBS
                        UcScorecardDetail1.PDBS1 = PDBS1
                        UcScorecardDetail1.PDBS2 = PDBS2
                        UcScorecardDetail1.PDBS3 = PDBS3
                        UcScorecardDetail1.PDBS4 = PDBS4
                        UcScorecardDetail1.DisplayValue()


                        ClassFilterTO1.currentdate = _currentdate
                        ClassFilterTO1.QTYBS = FilterQtyBS
                        ClassFilterTO1.AmountBS = FilterAmountBS
                        ClassFilterTO1.DisplayFilter = UcSupplierDashboard1.StringFilter

                        ClassFilterTO1.DisplayValue()

                        UcGroupBy1.CurrentDate = _currentdate
                        'GroupFPCPBS.Sort = "amount1 desc"
                        UcGroupBy1.bs = GroupFPCPBS


                        UcGroupBy2.CurrentDate = _currentdate
                        GroupSBUBS.Sort = "amount1 desc"
                        UcGroupBy2.bs = GroupSBUBS
                        UcGroupBy2.Label2.Text = "SBU"

                        UcGroupBy3.CurrentDate = _currentdate
                        GroupFamilyBS.Sort = "amount1 desc"
                        UcGroupBy3.bs = GroupFamilyBS
                        UcGroupBy3.Label2.Text = "Family"

                        UcGroupBy4.CurrentDate = _currentdate
                        GroupBrandBS.Sort = "amount1 desc"
                        UcGroupBy4.bs = GroupBrandBS
                        UcGroupBy4.Label2.Text = "Brand"



                        'UcContractualTermDGV1.DisplayValue()


                        UcContractualTermDGV1.DisplayValue()
                        ToolStripStatusLabel1.Text = "Done."

                        'FactoryContact
                        UcFactoryContact1.VendorBS = VendorFactoryContactBS
                        UcFactoryContact1.FactoryBS = FactoryBS
                        UcFactoryContact1.ContactBS = ContactBS
                        UcFactoryContact1.DisplayValue()

                        'SIFIDentitySheet
                        UcsifIdentity1.sifidbs = SIFIDBS
                        UcsifIdentity1.sifbs = SIFBS
                        UcsifIdentity1.idbs = IDBS

                        UcsifIdentity1.showData()

                        UcPriceCMMF1.DataGridView1.AutoGenerateColumns = False
                        UcPriceCMMF1.VendorQueryType = VendorQueryType
                        UcPriceCMMF1.vendorcode = _vendorcode
                        UcPriceCMMF1.shortname = _shortname
                        'UcPriceCMMF1.DataGridView1.DataSource = PriceCMMFBS
                        UcPriceCMMF1.bs = PriceCMMFBS

                        UcPriceCMMF1.ResetFilter()

                    Case 5
                        'UcSupplierDashboard1.SDDS = SDDS

                        'DateTimePicker1.Value = CDate(String.Format("{0}-{1}-1", message.Substring(0, 4), message.Substring(4, 2)))
                        Dim dr As DataRow = InitDS.Tables(0).Rows(0)
                        _currentdate = CDate(String.Format("{0}-{1}-1", dr.Item(0).ToString.Substring(0, 4), dr.Item(0).ToString.Substring(4, 2)))
                        'DateTimePicker1.Value = CDate(String.Format("{0}-{1}-1", dr.Item(0).ToString.Substring(0, 4), dr.Item(0).ToString.Substring(4, 2)))
                        DateTimePicker1.Value = _currentdate
                        dr = InitDS.Tables(1).Rows(TURNOVER)
                        UctoHeader1.ToDate = dr.Item(0)
                        dr = InitDS.Tables(1).Rows(SEBASIA_PLATFORM)
                        UctoHeader1.CurrentDate = _currentdate
                        UctoHeader1.SebPlatformDate = dr.Item(0)
                        UctoHeader1.DisplayLatestUpdate()
                        UctoHeader1.initLabel()
                        dr = InitDS.Tables(1).Rows(NQSU)
                        UcScorecardHeader1.QualityDate = dr.Item(0)
                        dr = InitDS.Tables(1).Rows(LOGISTICS)
                        UcScorecardHeader1.LogisticDate = dr.Item(0)
                        dr = InitDS.Tables(1).Rows(PROJECT_STATUS)
                        UcScorecardHeader1.ProjectDate = dr.Item(0)
                        dr = InitDS.Tables(2).Rows(0)
                        UcScorecardHeader1.QLatestPeriod = dr.Item(0).ToString

                        UcScorecardHeader1.DisplayLatestUpdate()
                        UcScorecardHeader1.CurrentDate = _currentdate
                        UcScorecardHeader1.initLabel()

                        dr = InitDS.Tables(1).Rows(SAP_VENDOR_PAYMENT)
                        UcContractualTermDGV1.VendorPaymentDate = dr.Item(0)
                        dr = InitDS.Tables(1).Rows(VFP_PROGRAM)
                        UcContractualTermDGV1.VFPDate = dr.Item(0)
                        UcContractualTermDGV1.DisplayLatestUpdate()


                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 7
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                    Case 8
                        UcChart1.DS = myDS
                    Case 9

                        Label5.Text = String.Format("Latest import date of budget and Forecast: {0:dd-MMM-yyyy}", bFAdapter.latestBudgetPeriod)
                        Label6.Text = String.Format("Lastest import date of Supply Chain Turnover Report: {0:dd-MMM-yyyy}", bFAdapter.latestTOPeriod)
                        showDataGridView(UcBudgetForecast1, "Supplier Total", bFAdapter.DS.Tables(0), bFAdapter.DS.Tables(1))
                        showDataGridView(UcBudgetForecast2, "By Filter", bFAdapter.DS.Tables(2), bFAdapter.DS.Tables(3))
                        showDataGridViewByGroup(UcBudgetForecast3, "SBU", bFAdapter.DS.Tables(4), bFAdapter.DS.Tables(5))
                        showDataGridViewByGroup(UcBudgetForecast4, "Family", bFAdapter.DS.Tables(6), bFAdapter.DS.Tables(7))
                        showDataGridViewByGroup(UcBudgetForecast5, "Brand", bFAdapter.DS.Tables(8), bFAdapter.DS.Tables(9))
                        showDataGridViewByGroup(UcBudgetForecast6, "Range", bFAdapter.DS.Tables(10), bFAdapter.DS.Tables(11))

                        
                End Select
            Catch ex As Exception
                MessageBox.Show(String.Format("FormSupplierDashBoard {0}::{1}::Id {2} {3}.", Me.Name, System.Reflection.MethodInfo.GetCurrentMethod(), id, ex.Message))
            End Try
        End If

    End Sub
    Public Sub showDataGridView(ByVal UcBudgetForeCast As Object, ByVal Title As String, ByVal Table1 As DataTable, ByVal Table2 As DataTable)
        'UCBudgetForecast1
        UcBudgetForeCast.title = Title
        UcBudgetForeCast.init()
        'Fill Forecast budget
        UcBudgetForeCast.DataGridView1.AutoGenerateColumns = True
        'hide vendorcode
        UcBudgetForeCast.DataGridView1.DataSource = Table1
        UcBudgetForeCast.DataGridView1.Columns(0).Visible = False
        Dim dgvwidth As Long
        For i = 1 To UcBudgetForeCast.DataGridView1.ColumnCount - 1
            With UcBudgetForeCast.DataGridView1.Columns(i)
                .DefaultCellStyle.Format = "#,##0.00"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Width = 100
            End With
            If UcBudgetForeCast.DataGridView1.Columns(i).HeaderText = "." Or UcBudgetForeCast.DataGridView1.Columns(i).HeaderText = "" Then
                UcBudgetForeCast.DataGridView1.Columns(i).HeaderText = ""
                UcBudgetForeCast.DataGridView1.Columns(i).Width = 50
            End If
            dgvwidth += UcBudgetForeCast.DataGridView1.Columns(i).Width
        Next
        UcBudgetForeCast.DataGridView1.Width = dgvwidth + 70

        UcBudgetForeCast.DataGridView2.Location = New Point(UcBudgetForeCast.DataGridView1.Location.X + UcBudgetForeCast.DataGridView1.Width + 20, 4)
        UcBudgetForeCast.DataGridView2.AutoGenerateColumns = True
        'hide vendorcode
        UcBudgetForeCast.DataGridView2.DataSource = Table2
        UcBudgetForeCast.DataGridView2.Columns(0).Visible = False
        dgvwidth = 0
        For i = 1 To UcBudgetForeCast.DataGridView2.ColumnCount - 1
            With UcBudgetForeCast.DataGridView2.Columns(i)
                .DefaultCellStyle.Format = "#,##0.00"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Width = 100
            End With
            If UcBudgetForeCast.DataGridView2.Columns(i).HeaderText = " " Then
                UcBudgetForeCast.DataGridView2.Columns(i).Width = 50
            End If
            dgvwidth += UcBudgetForeCast.DataGridView2.Columns(i).Width
        Next
        UcBudgetForeCast.DataGridView2.Width = dgvwidth + 70
        UcBudgetForeCast.Width = UcBudgetForeCast.DataGridView2.Location.X + UcBudgetForeCast.DataGridView2.Width + 30
    End Sub
    Private Sub showDataGridViewByGroup(ByVal UcBudgetForecast As Object, ByVal Title As String, ByVal Table1 As DataTable, ByVal Table2 As DataTable)
        'UCBudgetForecast3
        'Fill Forecast budget
        Dim dgvWidth As Integer
        UcBudgetForecast.title = Title
        UcBudgetForecast.init()
        UcBudgetForecast.DataGridView1.AutoGenerateColumns = True
        'hide vendorcode
        UcBudgetForecast.DataGridView1.DataSource = Table1
        UcBudgetForecast.DataGridView1.Columns(0).HeaderText = Title
        UcBudgetForecast.DataGridView1.Columns(0).Width = 200
        dgvWidth = 200
        For i = 1 To UcBudgetForecast.DataGridView1.ColumnCount - 1
            With UcBudgetForecast.DataGridView1.Columns(i)
                .DefaultCellStyle.Format = "#,##0.00"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Width = 100
            End With
            If UcBudgetForecast.DataGridView1.Columns(i).HeaderText = "." Or UcBudgetForecast.DataGridView1.Columns(i).HeaderText = "" Then
                UcBudgetForecast.DataGridView1.Columns(i).HeaderText = ""
                UcBudgetForecast.DataGridView1.Columns(i).Width = 50
            End If
            dgvWidth += UcBudgetForecast.DataGridView1.Columns(i).Width
        Next
        UcBudgetForecast.DataGridView1.Width = dgvWidth + 70

        UcBudgetForecast.DataGridView2.Location = New Point(UcBudgetForecast.DataGridView1.Location.X + UcBudgetForecast.DataGridView1.Width + 20, 4)
        UcBudgetForecast.DataGridView2.AutoGenerateColumns = True
        'hide vendorcode
        UcBudgetForecast.DataGridView2.DataSource = Table2
        UcBudgetForecast.DataGridView2.Columns(0).HeaderText = Title
        UcBudgetForecast.DataGridView2.Columns(0).Width = 200
        dgvWidth = 200
        For i = 1 To UcBudgetForecast.DataGridView2.ColumnCount - 1
            With UcBudgetForecast.DataGridView2.Columns(i)
                .DefaultCellStyle.Format = "#,##0.00"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Width = 100
            End With
            If UcBudgetForecast.DataGridView2.Columns(i).HeaderText = " " Then
                UcBudgetForecast.DataGridView2.Columns(i).Width = 50
            End If
            dgvWidth += UcBudgetForecast.DataGridView2.Columns(i).Width
        Next
        UcBudgetForecast.DataGridView2.Width = dgvWidth + 70
        UcBudgetForecast.Width = UcBudgetForecast.DataGridView2.Location.X + UcBudgetForecast.DataGridView2.Width + 30
    End Sub
    Private Sub FormSupplierDashboard_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

    End Sub

    Private Sub FormSupplierDashboard_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'RemoveHandler UcSupplierDashboard1.Button1.Click, AddressOf LetsClearTo
    End Sub



    Private Sub UcSupplierDashboard1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ShowProgress()
        ProgressReport(1, "Processing... Please wait")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        UcChart1.ClearValue()
        _shortname = UcSupplierDashboard1.ShortName
        _vendorcode = UcSupplierDashboard1.Vendorcode
        If (IsNothing(_shortname) And _vendorcode = 0) Then
            'MessageBox.Show("Please select Short Name or VendorCode")
            UcSupplierDashboard1.ErrorProvider1.SetError(UcSupplierDashboard1.TextBox1, "Please select Short Name or Vendor Name.")
            UcSupplierDashboard1.ErrorProvider1.SetError(UcSupplierDashboard1.TextBox2, "Please select Vendor Name or Vendor Name.")
            Exit Sub
        End If
        UcSupplierDashboard1.ErrorProvider1.SetError(UcSupplierDashboard1.TextBox1, "")
        UcSupplierDashboard1.ErrorProvider1.SetError(UcSupplierDashboard1.TextBox2, "")
        'If String.Format("{0:yyyyMM}", DateTimePicker2.Value) = String.Format("{0:yyyyMM}", DateTimePicker3.Value) Then

        '    ErrorProvider1.SetError(DateTimePicker2, "Period select period.")
        '    ErrorProvider1.SetError(DateTimePicker3, "Period select period.")
        '    Exit Sub
        'End If
        'ErrorProvider1.SetError(DateTimePicker2, "")
        'ErrorProvider1.SetError(DateTimePicker3, "")
        If Not myThread.IsAlive Then
            'UcChart1.CheckBox1.Checked = False
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoChartQuery)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If

    End Sub

    Private Sub DoChartQuery()
        Dim sb As New StringBuilder
        Dim TextFilter As String = String.Empty
        TextFilter = UcSupplierDashboard1.getFilterCriteria
        ProductTypeCriteria = UcSupplierDashboard1.ProductTypeCriteria.Replace("''", "'")
        Dim period1 As Date
        Dim period2 As Date
        period1 = DateTimePicker2.Value
        period2 = DateTimePicker3.Value

        Dim mySqlstr As String = String.Empty
        Dim myMessage As String = String.Empty
        Dim myFieldList As New StringBuilder

        mySqlstr = String.Format("select count(period) from (select period from doc.turnover tu left join vendor v on v.vendorcode = tu.vendorcode {0} and period >= {1:yyyyMM} and period <= {2:yyyyMM}  group by period order by period) as foo", toCriteria2, period1, period2)
        Dim myresult As Integer
        'Filter Amount
        'Create dynamic list for  amount1 numeric,amount2 numeric,amount3 numeric,amount4 numeric,amount5 numeric
        'Get Value from this query "select period from doc.turnover tu left join vendor v on v.vendorcode = tu.vendorcode {1} and period >= {2:yyyyMM} and period <= {3:yyyyMM}  group by period order by period"

        If DbAdapter1.ExecuteScalar(mySqlstr, myresult, message:=myMessage) Then
            If myresult = 0 Then
                Exit Sub
            End If
            For i = 0 To myresult - 1
                If myFieldList.Length > 0 Then
                    myFieldList.Append(",")
                End If
                myFieldList.Append(String.Format("amount{0} numeric ", i + 1))
            Next
        Else
            MessageBox.Show(myMessage)
            Exit Sub
        End If

        sb.Append(String.Format("select * from crosstab('select {0},period,sum(amount) from doc.turnover tu" &
                     " left join vendor v on v.vendorcode = tu.vendorcode" &
                     " left join materialmaster mm on mm.cmmf = tu.cmmf" &
                     " left join family f on f.familyid = tu.comfam" &
                     "  {1}  and period >= {2:yyyyMM} and period <= {3:yyyyMM} {5}" &
                     " group by {0},period order by period ','select period from doc.turnover tu left join vendor v on v.vendorcode = tu.vendorcode {1} and period >= {2:yyyyMM} and period <= {3:yyyyMM}  group by period order by period') as (sbu text, {6});", "v." + VendorQueryType.ToString, toCriteria2.Replace("'", "''"), period1, period2, Year(_currentdate), TextFilter, myFieldList.ToString))
        sb.Append(String.Format("select period from doc.turnover tu left join vendor v on v.vendorcode = tu.vendorcode {0} and period >= {1:yyyyMM} and period <= {2:yyyyMM}  group by period order by period", toCriteria2, period1, period2))
        myDS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString, myDS, myMessage) Then
            'Show Chart
            ProgressReport(8, "DisplayChart")


        Else
            MessageBox.Show(myMessage)
            Exit Sub
        End If


    End Sub

    Private Sub ChangeVendorList(ByVal criteria As String)
        If criteria <> "" Then
            UcSupplierDashboard1.getSupplierInfo2(criteria.ToString.Replace("status in (", "").Replace(")", ""))
        Else
            UcSupplierDashboard1.ShowOriginal()
        End If
    End Sub

    Private Sub CheckTabPage9(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not myThreadQuery.IsAlive Then
            SelectedTab = False
            If CType(sender, TabControl).SelectedTab Is TabPage9 Then
                SelectedTab = True
            End If
        Else
            If CType(sender, TabControl).SelectedTab Is TabPage9 Then
                SelectedTab = True
            Else
                SelectedTab = False
            End If

        End If
    End Sub

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        CheckTabPage9(sender, e)
    End Sub



    Private Sub TabControl1_Selected(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlEventArgs) Handles TabControl1.Selected
        CheckTabPage9(sender, e)


    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        UcSupplierDashboard1.currentDate = DateTimePicker1.Value.Date
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Not myThreadQuery.IsAlive AndAlso Not myThread2.IsAlive AndAlso Not myThread.IsAlive Then
            If Not IsNothing(DS) Then
                If PriceCMMFBS.Count = 0 Then
                    If MessageBox.Show("No CMMF Price List Record Found. Continue?", "", Windows.Forms.MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If
                'Dim myform As New FormReportTOSC(PurchaseAmountSQL)
                Dim myform As New FormReportTOSC(PurchaseAmountSQL, bFAdapter.DS)
                myform.CurrentMonth = _currentdate
                myform.frontshortname = UcSupplierInfo1.Shortname
                myform.frontsuppliername = UcSupplierInfo1.TextBox3.Text
                myform.frontsuppliercode = UcSupplierInfo1.TextBox2.Text
                myform.frontproductype = IIf(UcSupplierDashboard1.ProductTypeFilter = "", "ALL", UcSupplierDashboard1.ProductTypeFilter)
                myform.frontsbu = UcSupplierDashboard1.TextBox9.Text
                myform.frontfamily = UcSupplierDashboard1.TextBox4.Text
                myform.frontbrand = UcSupplierDashboard1.TextBox3.Text
                myform.frontrangecode = UcSupplierDashboard1.TextBox6.Text
                myform.frontcmmf = UcSupplierDashboard1.TextBox5.Text
                myform.frontcommercialcode = UcSupplierDashboard1.TextBox8.Text
                myform.frontsebasiaplatform = IIf(UcSupplierDashboard1.CheckBox1.Checked, "Y", "N")
                myform.frontvendorstatus = UcSupplierDashboard1.statusFilter
                myform.filterApplied = ClassFilterTO1.TextBox19.Text

                myform.sify = validDec(ClassSIFTO1.TextBox1.Text)
                myform.sify1 = validDec(ClassSIFTO1.TextBox2.Text)
                myform.sify2 = validDec(ClassSIFTO1.TextBox3.Text)
                myform.sify3 = validDec(ClassSIFTO1.TextBox4.Text)
                myform.sify4 = validDec(ClassSIFTO1.TextBox5.Text)
                myform.TOSEBQTY = validDec(ClassSEBTO1.TextBox1.Text)
                myform.TOSEBQTY1 = validDec(ClassSEBTO1.TextBox2.Text)
                myform.TOSEBQTY2 = validDec(ClassSEBTO1.TextBox3.Text)
                myform.TOSEBQTY3 = validDec(ClassSEBTO1.TextBox4.Text)
                myform.TOSEBQTY4 = validDec(ClassSEBTO1.TextBox5.Text)
                myform.TOSEBAMT = validDec(ClassSEBTO1.TextBox10.Text)
                myform.TOSEBAMT1 = validDec(ClassSEBTO1.TextBox9.Text)
                myform.TOSEBAMT2 = validDec(ClassSEBTO1.TextBox8.Text)
                myform.TOSEBAMT3 = validDec(ClassSEBTO1.TextBox7.Text)
                myform.TOSEBAMT4 = validDec(ClassSEBTO1.TextBox6.Text)
                myform.FilterQty = validDec(ClassFilterTO1.TextBox1.Text)
                myform.FilterQty1 = validDec(ClassFilterTO1.TextBox2.Text)
                myform.FilterQty2 = validDec(ClassFilterTO1.TextBox3.Text)
                myform.FilterQty3 = validDec(ClassFilterTO1.TextBox4.Text)
                myform.FilterQty4 = validDec(ClassFilterTO1.TextBox5.Text)
                myform.FilterAmt = validDec(ClassFilterTO1.TextBox10.Text)
                myform.FilterAmt1 = validDec(ClassFilterTO1.TextBox9.Text)
                myform.FilterAmt2 = validDec(ClassFilterTO1.TextBox8.Text)
                myform.FilterAmt3 = validDec(ClassFilterTO1.TextBox7.Text)
                myform.FilterAmt4 = validDec(ClassFilterTO1.TextBox6.Text)

                myform.nqsu = UcScorecardDetail1.TextBox25.Text 'validDec(IIf(IsNumeric(UcScorecardDetail1.TextBox25.Text), UcScorecardDetail1.TextBox25.Text, 0))
                myform.nqsu1 = UcScorecardDetail1.TextBox24.Text 'validDec(IIf(IsNumeric(UcScorecardDetail1.TextBox24.Text), UcScorecardDetail1.TextBox24.Text, 0))
                myform.nqsu2 = UcScorecardDetail1.TextBox23.Text ' validDec(IIf(IsNumeric(UcScorecardDetail1.TextBox23.Text), UcScorecardDetail1.TextBox23.Text, 0))
                myform.nqsu3 = UcScorecardDetail1.TextBox22.Text 'validDec(IIf(IsNumeric(UcScorecardDetail1.TextBox22.Text), UcScorecardDetail1.TextBox22.Text, 0))
                myform.nqsu4 = UcScorecardDetail1.TextBox21.Text '  validDec(IIf(IsNumeric(UcScorecardDetail1.TextBox21.Text), UcScorecardDetail1.TextBox21.Text, DBNull.Value))

                myform.scsasl = validDec(UcScorecardDetail1.TextBox45.Text.Replace("%", "")) / 100
                myform.scsasl1 = validDec(UcScorecardDetail1.TextBox44.Text.Replace("%", "")) / 100
                myform.scsasl2 = validDec(UcScorecardDetail1.TextBox43.Text.Replace("%", "")) / 100
                myform.scsasl3 = validDec(UcScorecardDetail1.TextBox42.Text.Replace("%", "")) / 100
                myform.scsasl4 = validDec(UcScorecardDetail1.TextBox41.Text.Replace("%", "")) / 100

                myform.scssl = validDec(UcScorecardDetail1.TextBox40.Text.Replace("%", "")) / 100
                myform.scssl1 = validDec(UcScorecardDetail1.TextBox39.Text.Replace("%", "")) / 100
                myform.scssl2 = validDec(UcScorecardDetail1.TextBox38.Text.Replace("%", "")) / 100
                myform.scssl3 = validDec(UcScorecardDetail1.TextBox37.Text.Replace("%", "")) / 100
                myform.scssl4 = validDec(UcScorecardDetail1.TextBox36.Text.Replace("%", "")) / 100

                myform.scsslnet = validDec(UcScorecardDetail1.TextBox75.Text.Replace("%", "")) / 100
                myform.scsslnet1 = validDec(UcScorecardDetail1.TextBox74.Text.Replace("%", "")) / 100
                myform.scsslnet2 = validDec(UcScorecardDetail1.TextBox73.Text.Replace("%", "")) / 100
                myform.scsslnet3 = validDec(UcScorecardDetail1.TextBox72.Text.Replace("%", "")) / 100
                myform.scsslnet4 = validDec(UcScorecardDetail1.TextBox71.Text.Replace("%", "")) / 100

                myform.sclt = validDec(UcScorecardDetail1.TextBox35.Text)
                myform.sclt1 = validDec(UcScorecardDetail1.TextBox34.Text)
                myform.sclt2 = validDec(UcScorecardDetail1.TextBox33.Text)
                myform.sclt3 = validDec(UcScorecardDetail1.TextBox32.Text)
                myform.sclt4 = validDec(UcScorecardDetail1.TextBox31.Text)

                myform.scno = validDec(UcScorecardDetail1.TextBox30.Text)
                myform.scno1 = validDec(UcScorecardDetail1.TextBox29.Text)
                myform.scno2 = validDec(UcScorecardDetail1.TextBox28.Text)
                myform.scno3 = validDec(UcScorecardDetail1.TextBox27.Text)
                myform.scno4 = validDec(UcScorecardDetail1.TextBox26.Text)

                myform.scsh = validDec(UcScorecardDetail1.TextBox50.Text)
                myform.scsh1 = validDec(UcScorecardDetail1.TextBox49.Text)
                myform.scsh2 = validDec(UcScorecardDetail1.TextBox48.Text)
                myform.scsh3 = validDec(UcScorecardDetail1.TextBox47.Text)
                myform.scsh4 = validDec(UcScorecardDetail1.TextBox46.Text)

                myform.piavg = validDec(UcScorecardDetail1.TextBox1.Text)
                myform.piavg1 = validDec(UcScorecardDetail1.TextBox2.Text)
                myform.piavg2 = validDec(UcScorecardDetail1.TextBox3.Text)
                myform.piavg3 = validDec(UcScorecardDetail1.TextBox4.Text)
                myform.piavg4 = validDec(UcScorecardDetail1.TextBox5.Text)

                myform.pilkp = validDec(UcScorecardDetail1.TextBox10.Text)
                myform.pilkp1 = validDec(UcScorecardDetail1.TextBox9.Text)
                myform.pilkp2 = validDec(UcScorecardDetail1.TextBox8.Text)
                myform.pilkp3 = validDec(UcScorecardDetail1.TextBox7.Text)
                myform.pilkp4 = validDec(UcScorecardDetail1.TextBox6.Text)

                myform.pistd = validDec(UcScorecardDetail1.TextBox20.Text)
                myform.pistd1 = validDec(UcScorecardDetail1.TextBox19.Text)
                myform.pistd2 = validDec(UcScorecardDetail1.TextBox18.Text)
                myform.pistd3 = validDec(UcScorecardDetail1.TextBox17.Text)
                myform.pistd4 = validDec(UcScorecardDetail1.TextBox16.Text)

                myform.pvstd = validDec(UcScorecardDetail1.TextBox15.Text)
                myform.pvstd1 = validDec(UcScorecardDetail1.TextBox14.Text)
                myform.pvstd2 = validDec(UcScorecardDetail1.TextBox13.Text)
                myform.pvstd3 = validDec(UcScorecardDetail1.TextBox12.Text)
                myform.pvstd4 = validDec(UcScorecardDetail1.TextBox11.Text)

                myform.pdrp14 = UcScorecardDetail1.TextBox70.Text 'validDec(UcScorecardDetail1.TextBox70.Text)
                myform.pdrp141 = UcScorecardDetail1.TextBox69.Text ' validDec(UcScorecardDetail1.TextBox69.Text)
                myform.pdrp142 = UcScorecardDetail1.TextBox68.Text 'validDec(UcScorecardDetail1.TextBox68.Text)
                myform.pdrp143 = UcScorecardDetail1.TextBox67.Text 'validDec(UcScorecardDetail1.TextBox67.Text)
                myform.pdrp144 = UcScorecardDetail1.TextBox66.Text ' validDec(UcScorecardDetail1.TextBox66.Text)

                myform.pdrp5 = UcScorecardDetail1.TextBox65.Text 'validDec(UcScorecardDetail1.TextBox65.Text)
                myform.pdrp51 = UcScorecardDetail1.TextBox64.Text 'validDec(UcScorecardDetail1.TextBox64.Text)
                myform.pdrp52 = UcScorecardDetail1.TextBox63.Text 'validDec(UcScorecardDetail1.TextBox63.Text)
                myform.pdrp53 = UcScorecardDetail1.TextBox62.Text 'validDec(UcScorecardDetail1.TextBox62.Text)
                myform.pdrp54 = UcScorecardDetail1.TextBox61.Text 'validDec(UcScorecardDetail1.TextBox61.Text)

                myform.pdramp = UcScorecardDetail1.TextBox60.Text 'validDec(UcScorecardDetail1.TextBox60.Text)
                myform.pdramp1 = UcScorecardDetail1.TextBox59.Text 'validDec(UcScorecardDetail1.TextBox59.Text)
                myform.pdramp2 = UcScorecardDetail1.TextBox58.Text '(UcScorecardDetail1.TextBox58.Text)
                myform.pdramp3 = UcScorecardDetail1.TextBox57.Text 'validDec(UcScorecardDetail1.TextBox57.Text)
                myform.pdramp4 = UcScorecardDetail1.TextBox56.Text 'validDec(UcScorecardDetail1.TextBox56.Text)

                myform.pdpot = UcScorecardDetail1.TextBox55.Text 'validDec(UcScorecardDetail1.TextBox55.Text.Replace("%", "")) / 100
                myform.pdpot1 = UcScorecardDetail1.TextBox54.Text 'validDec(UcScorecardDetail1.TextBox54.Text.Replace("%", "")) / 100
                myform.pdpot2 = UcScorecardDetail1.TextBox53.Text ' validDec(UcScorecardDetail1.TextBox53.Text.Replace("%", "")) / 100
                myform.pdpot3 = UcScorecardDetail1.TextBox52.Text 'validDec(UcScorecardDetail1.TextBox52.Text.Replace("%", "")) / 100
                myform.pdpot4 = UcScorecardDetail1.TextBox51.Text ' validDec(UcScorecardDetail1.TextBox51.Text.Replace("%", "")) / 100
                myform.ToolingHDBS = ToolingHeader
                myform.ToolingDetailsALLBS = ToolingDetailsAll
                myform.SAPBS = SAPBS
                myform.ContractBS = ContractBS
                myform.CitiProgramBS = CitiProgramBS
                myform.VendorCurrencyBS = VendorCurrencyBS
                myform.AuthletterBS = ALetterBS
                myform.ProjectSpecBS = PSpecBS
                myform.PCharterBS = PCharterBS
                myform.SocialAuditBS = SocialAuditBS
                myform.PanelStatusBS = UcSupplierDashboard1.PanelStatus
                myform.SupplierTechnologyBS = UcSupplierDashboard1.SupplierTechnologyBS
                myform.SupplierPanelHistory = UcSupplierDashboard1.SupplierPanelHistoryBS
                myform.VendorBS = VendorFactoryContactBS
                myform.FactoryBS = FactoryBS
                myform.ContactBS = ContactBS
                myform.SIFIDBS = SIFIDBS
                myform.SIFBS = SIFBS
                myform.IDBS = IDBS

                myform.PriceCMMFBS = PriceCMMFBS
                myform.LKPFiter = UcPriceCMMF1.GetLKPFilter
                myform.CMMFFilter = UcPriceCMMF1.GetCMMFFilter

                myform.ShowDialog()
            Else
                MessageBox.Show("Click show data first.")
            End If
        Else
            MessageBox.Show(String.Format("Please wait until the current process is finished."))
        End If



    End Sub
    Private Function validDec(ByVal value As Object) As Decimal
        If Not IsNumeric(value) Then
            value = Nothing
        End If
        Return value
    End Function

    Private Sub RefreshSupplierBasicInformation()
        UcFactoryContact1.SupplierTechnologyBS = UcSupplierDashboard1.SupplierTechnologyBS
        UcFactoryContact1.PanelHistoryBS = UcSupplierDashboard1.SupplierPanelHistoryBS
        'UcFactoryContact1.DisplayValue()
    End Sub

    Private Sub ShowContractPopUp()

        Dim myform As New DialogContractTerms(ContractBS, ContractualTerms1BS, ContractualTerms2BS, ContractualTerms3BS)
        myform.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        If IsNothing(DS) Then
            MessageBox.Show("Click show data first.")
            Exit Sub
        End If
        If UcSupplierInfo1.Shortname = "" Then
            MessageBox.Show("Shortname is blank")
            Exit Sub
        End If
        'Dim myform As New DialogExportBuyerInput(PurchaseAmountSQL, UcSupplierInfo1.Shortname)
        Dim myform As New DialogExportBuyerInput2(PurchaseAmountSQL, UcSupplierInfo1.Shortname)
        myform.frontproductype = UcSupplierDashboard1.ProductTypeFilter
        myform.CurrentMonth = _currentdate
        myform.frontshortname = UcSupplierInfo1.Shortname
        myform.frontsuppliername = UcSupplierInfo1.TextBox3.Text
        myform.frontsuppliercode = UcSupplierInfo1.TextBox2.Text
        myform.frontproductype = IIf(UcSupplierDashboard1.ProductTypeFilter = "", "ALL", UcSupplierDashboard1.ProductTypeFilter)
        myform.frontsbu = UcSupplierDashboard1.TextBox9.Text
        myform.frontfamily = UcSupplierDashboard1.TextBox4.Text
        myform.frontbrand = UcSupplierDashboard1.TextBox3.Text
        myform.frontrangecode = UcSupplierDashboard1.TextBox6.Text
        myform.frontcmmf = UcSupplierDashboard1.TextBox5.Text
        myform.frontcommercialcode = UcSupplierDashboard1.TextBox8.Text
        myform.frontsebasiaplatform = IIf(UcSupplierDashboard1.CheckBox1.Checked, "Y", "N")
        myform.frontvendorstatus = UcSupplierDashboard1.statusFilter
        myform.filterApplied = ClassFilterTO1.TextBox19.Text

        myform.sify = validDec(ClassSIFTO1.TextBox1.Text)
        myform.sify1 = validDec(ClassSIFTO1.TextBox2.Text)
        myform.sify2 = validDec(ClassSIFTO1.TextBox3.Text)
        myform.sify3 = validDec(ClassSIFTO1.TextBox4.Text)
        myform.sify4 = validDec(ClassSIFTO1.TextBox5.Text)
        myform.TOSEBQTY = validDec(ClassSEBTO1.TextBox1.Text)
        myform.TOSEBQTY1 = validDec(ClassSEBTO1.TextBox2.Text)
        myform.TOSEBQTY2 = validDec(ClassSEBTO1.TextBox3.Text)
        myform.TOSEBQTY3 = validDec(ClassSEBTO1.TextBox4.Text)
        myform.TOSEBQTY4 = validDec(ClassSEBTO1.TextBox5.Text)
        myform.TOSEBAMT = validDec(ClassSEBTO1.TextBox10.Text)
        myform.TOSEBAMT1 = validDec(ClassSEBTO1.TextBox9.Text)
        myform.TOSEBAMT2 = validDec(ClassSEBTO1.TextBox8.Text)
        myform.TOSEBAMT3 = validDec(ClassSEBTO1.TextBox7.Text)
        myform.TOSEBAMT4 = validDec(ClassSEBTO1.TextBox6.Text)
        myform.FilterQty = validDec(ClassFilterTO1.TextBox1.Text)
        myform.FilterQty1 = validDec(ClassFilterTO1.TextBox2.Text)
        myform.FilterQty2 = validDec(ClassFilterTO1.TextBox3.Text)
        myform.FilterQty3 = validDec(ClassFilterTO1.TextBox4.Text)
        myform.FilterQty4 = validDec(ClassFilterTO1.TextBox5.Text)
        myform.FilterAmt = validDec(ClassFilterTO1.TextBox10.Text)
        myform.FilterAmt1 = validDec(ClassFilterTO1.TextBox9.Text)
        myform.FilterAmt2 = validDec(ClassFilterTO1.TextBox8.Text)
        myform.FilterAmt3 = validDec(ClassFilterTO1.TextBox7.Text)
        myform.FilterAmt4 = validDec(ClassFilterTO1.TextBox6.Text)

        myform.nqsu = UcScorecardDetail1.TextBox25.Text
        myform.nqsu1 = UcScorecardDetail1.TextBox24.Text
        myform.nqsu2 = UcScorecardDetail1.TextBox23.Text
        myform.nqsu3 = UcScorecardDetail1.TextBox22.Text
        myform.nqsu4 = UcScorecardDetail1.TextBox21.Text

        myform.scsasl = validDec(UcScorecardDetail1.TextBox45.Text.Replace("%", "")) / 100
        myform.scsasl1 = validDec(UcScorecardDetail1.TextBox44.Text.Replace("%", "")) / 100
        myform.scsasl2 = validDec(UcScorecardDetail1.TextBox43.Text.Replace("%", "")) / 100
        myform.scsasl3 = validDec(UcScorecardDetail1.TextBox42.Text.Replace("%", "")) / 100
        myform.scsasl4 = validDec(UcScorecardDetail1.TextBox41.Text.Replace("%", "")) / 100

        myform.scssl = validDec(UcScorecardDetail1.TextBox40.Text.Replace("%", "")) / 100
        myform.scssl1 = validDec(UcScorecardDetail1.TextBox39.Text.Replace("%", "")) / 100
        myform.scssl2 = validDec(UcScorecardDetail1.TextBox38.Text.Replace("%", "")) / 100
        myform.scssl3 = validDec(UcScorecardDetail1.TextBox37.Text.Replace("%", "")) / 100
        myform.scssl4 = validDec(UcScorecardDetail1.TextBox36.Text.Replace("%", "")) / 100

        myform.scsslnet = validDec(UcScorecardDetail1.TextBox75.Text.Replace("%", "")) / 100
        myform.scsslnet1 = validDec(UcScorecardDetail1.TextBox74.Text.Replace("%", "")) / 100
        myform.scsslnet2 = validDec(UcScorecardDetail1.TextBox73.Text.Replace("%", "")) / 100
        myform.scsslnet3 = validDec(UcScorecardDetail1.TextBox72.Text.Replace("%", "")) / 100
        myform.scsslnet4 = validDec(UcScorecardDetail1.TextBox71.Text.Replace("%", "")) / 100

        myform.sclt = validDec(UcScorecardDetail1.TextBox35.Text)
        myform.sclt1 = validDec(UcScorecardDetail1.TextBox34.Text)
        myform.sclt2 = validDec(UcScorecardDetail1.TextBox33.Text)
        myform.sclt3 = validDec(UcScorecardDetail1.TextBox32.Text)
        myform.sclt4 = validDec(UcScorecardDetail1.TextBox31.Text)

        myform.scno = validDec(UcScorecardDetail1.TextBox30.Text)
        myform.scno1 = validDec(UcScorecardDetail1.TextBox29.Text)
        myform.scno2 = validDec(UcScorecardDetail1.TextBox28.Text)
        myform.scno3 = validDec(UcScorecardDetail1.TextBox27.Text)
        myform.scno4 = validDec(UcScorecardDetail1.TextBox26.Text)

        myform.scsh = validDec(UcScorecardDetail1.TextBox50.Text)
        myform.scsh1 = validDec(UcScorecardDetail1.TextBox49.Text)
        myform.scsh2 = validDec(UcScorecardDetail1.TextBox48.Text)
        myform.scsh3 = validDec(UcScorecardDetail1.TextBox47.Text)
        myform.scsh4 = validDec(UcScorecardDetail1.TextBox46.Text)

        'myform.piavg = validDec(UcScorecardDetail1.TextBox1.Text)
        'myform.piavg1 = validDec(UcScorecardDetail1.TextBox2.Text)
        'myform.piavg2 = validDec(UcScorecardDetail1.TextBox3.Text)
        'myform.piavg3 = validDec(UcScorecardDetail1.TextBox4.Text)
        'myform.piavg4 = validDec(UcScorecardDetail1.TextBox5.Text)

        myform.piavg = validDec(UcScorecardDetail1.TextBox85.Text) 'piavgfx
        myform.piavg1 = validDec(UcScorecardDetail1.TextBox84.Text)
        myform.piavg2 = validDec(UcScorecardDetail1.TextBox83.Text)
        myform.piavg3 = validDec(UcScorecardDetail1.TextBox82.Text)
        myform.piavg4 = validDec(UcScorecardDetail1.TextBox81.Text)

        'myform.pilkp = validDec(UcScorecardDetail1.TextBox10.Text)
        'myform.pilkp1 = validDec(UcScorecardDetail1.TextBox9.Text)
        'myform.pilkp2 = validDec(UcScorecardDetail1.TextBox8.Text)
        'myform.pilkp3 = validDec(UcScorecardDetail1.TextBox7.Text)
        'myform.pilkp4 = validDec(UcScorecardDetail1.TextBox6.Text)

        myform.pilkp = validDec(UcScorecardDetail1.TextBox80.Text) ' pistdfx
        myform.pilkp1 = validDec(UcScorecardDetail1.TextBox79.Text)
        myform.pilkp2 = validDec(UcScorecardDetail1.TextBox78.Text)
        myform.pilkp3 = validDec(UcScorecardDetail1.TextBox77.Text)
        myform.pilkp4 = validDec(UcScorecardDetail1.TextBox76.Text)

        myform.pistd = validDec(UcScorecardDetail1.TextBox20.Text)
        myform.pistd1 = validDec(UcScorecardDetail1.TextBox19.Text)
        myform.pistd2 = validDec(UcScorecardDetail1.TextBox18.Text)
        myform.pistd3 = validDec(UcScorecardDetail1.TextBox17.Text)
        myform.pistd4 = validDec(UcScorecardDetail1.TextBox16.Text)

        myform.pvstd = validDec(UcScorecardDetail1.TextBox15.Text)
        myform.pvstd1 = validDec(UcScorecardDetail1.TextBox14.Text)
        myform.pvstd2 = validDec(UcScorecardDetail1.TextBox13.Text)
        myform.pvstd3 = validDec(UcScorecardDetail1.TextBox12.Text)
        myform.pvstd4 = validDec(UcScorecardDetail1.TextBox11.Text)

        myform.pdrp14 = UcScorecardDetail1.TextBox70.Text
        myform.pdrp141 = UcScorecardDetail1.TextBox69.Text
        myform.pdrp142 = UcScorecardDetail1.TextBox68.Text
        myform.pdrp143 = UcScorecardDetail1.TextBox67.Text
        myform.pdrp144 = UcScorecardDetail1.TextBox66.Text

        myform.pdrp5 = UcScorecardDetail1.TextBox65.Text
        myform.pdrp51 = UcScorecardDetail1.TextBox64.Text
        myform.pdrp52 = UcScorecardDetail1.TextBox63.Text
        myform.pdrp53 = UcScorecardDetail1.TextBox62.Text
        myform.pdrp54 = UcScorecardDetail1.TextBox61.Text

        myform.pdramp = UcScorecardDetail1.TextBox60.Text
        myform.pdramp1 = UcScorecardDetail1.TextBox59.Text
        myform.pdramp2 = UcScorecardDetail1.TextBox58.Text
        myform.pdramp3 = UcScorecardDetail1.TextBox57.Text
        myform.pdramp4 = UcScorecardDetail1.TextBox56.Text

        myform.pdpot = UcScorecardDetail1.TextBox55.Text
        myform.pdpot1 = UcScorecardDetail1.TextBox54.Text
        myform.pdpot2 = UcScorecardDetail1.TextBox53.Text
        myform.pdpot3 = UcScorecardDetail1.TextBox52.Text
        myform.pdpot4 = UcScorecardDetail1.TextBox51.Text
        myform.ToolingHDBS = ToolingHeader
        myform.ToolingDetailsALLBS = ToolingDetailsAll
        myform.SAPBS = SAPBS
        myform.ContractBS = ContractBS
        myform.CitiProgramBS = CitiProgramBS
        myform.AuthletterBS = ALetterBS
        myform.ProjectSpecBS = PSpecBS
        myform.PCharterBS = PCharterBS
        myform.SocialAuditBS = SocialAuditBS
        myform.PanelStatusBS = UcSupplierDashboard1.PanelStatus
        myform.SupplierTechnologyBS = UcSupplierDashboard1.SupplierTechnologyBS
        myform.SupplierPanelHistory = UcSupplierDashboard1.SupplierPanelHistoryBS
        myform.VendorBS = VendorFactoryContactBS
        myform.FactoryBS = FactoryBS
        myform.ContactBS = ContactBS
        myform.SIFIDBS = SIFIDBS
        myform.SIFBS = SIFBS
        myform.IDBS = IDBS

        myform.PriceCMMFBS = PriceCMMFBS
        myform.LKPFiter = UcPriceCMMF1.GetLKPFilter
        myform.CMMFFilter = UcPriceCMMF1.GetCMMFFilter

        myform.ShowDialog()
    End Sub
End Class