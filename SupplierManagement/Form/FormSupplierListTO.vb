Imports System.Text
Imports Microsoft.Office.Interop
Imports System.Threading
Public Class FormSupplierListTO
    Dim YearStart As Integer
    Dim YearEnd As Integer
    Dim DbAdapter1 As DbAdapter = DbAdapter.getInstance

    Dim SBUFilter As String = String.Empty
    Dim FamilyFilter As String = String.Empty
    Dim ShortNameFilter As String = String.Empty
    Dim SupplierNameFilter As String = String.Empty
    Dim ProductTypeFilter As String = String.Empty
    Dim AppliedSBUFilter As String = String.Empty
    Dim AppliedFamilyFilter As String = String.Empty
    Dim AppliedShortNameFilter As String = String.Empty
    Dim AppliedSupplierNameFilter As String = String.Empty
    Dim AppliedProductTypeFilter As String = String.Empty
    Dim GroupBy As String = "v.shortname3"
    Dim VendorInfo As String = "doc.getvendorcode(shortname3,'(Active supplier,Active tooling supplier)')"

    Dim AllFilter As New StringBuilder
    Dim NQSUFilter As New StringBuilder
    Dim LogisticsFilter As New StringBuilder
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim DS As DataSet


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If myThread.IsAlive Then
            ProgressReport(1, "Background process still running. Please try it again later.")           
            Exit Sub
        End If
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        YearStart = DateTimePicker1.Value.Year
        YearEnd = DateTimePicker2.Value.Year

        runreport(sender, e)
    End Sub
    Private Sub runreport(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim sb As New StringBuilder
        Dim mymessage As String = String.Empty
        Dim counter As Integer = 1
        ProgressReport(6, "Preparing Helper Data. Please wait..")

        GetFilter()


        'TurnOver
        For i = YearStart To YearEnd 'TO
            'Debug.Print(i)
            ''TO Quantity
            'sb.Append(String.Format("select period,year,fpcp,t.vendorcode,v.vendorname,v.shortname,sbu,comfam,f.familyname,'{0} {1} Actual Qty' as description, qty as value from doc.turnover t" &
            '                        " left join vendor v on v.vendorcode = t.vendorcode" &
            '                        " left join family f on f.familyid = t.comfam" &
            '                        " where Year = {1} {2};", counter, i, AllFilter.ToString))
            'counter += 1
            ''TO Amount
            'sb.Append(String.Format("select period,year,fpcp,t.vendorcode,v.vendorname,v.shortname,sbu,comfam,f.familyname,'{0} {1} Actual PurAmnt' as description, amount as value from doc.turnover t" &
            '                        " left join vendor v on v.vendorcode = t.vendorcode" &
            '                        " left join family f on f.familyid = t.comfam" &
            '                        " where Year = {1} {2};", counter, i, AllFilter.ToString))
            'counter += 1
            'TO Quantity
            If sb.Length > 0 Then
                sb.Append(" union all ")
            End If
            sb.Append(String.Format("select {3} as shortname,{4},'{0:00} {1} Actual Qty' as description, sum(qty) as value from doc.turnover t" &
                                    " left join vendor v on v.vendorcode = t.vendorcode" &
                                    " left join family f on f.familyid = t.comfam" &
                                    " where Year = {1} {2} group by {3}", counter, i, AllFilter.ToString, GroupBy, VendorInfo))
            counter += 1

            If sb.Length > 0 Then
                sb.Append(" union all ")
            End If

            'TO Amount
            sb.Append(String.Format("select {3},{4},'{0:00} {1} Actual PurAmnt' as description, sum(amount) as value from doc.turnover t" &
                                    " left join vendor v on v.vendorcode = t.vendorcode" &
                                    " left join family f on f.familyid = t.comfam" &
                                    " where Year = {1} {2} group by {3}", counter, i, AllFilter.ToString, GroupBy, VendorInfo))
            counter += 1
        Next
        'Budget Volume QTY
        'sb.Append(String.Format("with latest as ( select distinct datatypeid from doc.budgetforecast where datatypeid in (1,2,8) and period = '{1}-01-01' order by datatypeid desc limit 1)" &
        '                        " select to_char(t.period,'FMYYYYMM'),to_char(t.period,'FMYYYY'),t.fpcp,t.vendorcode,v.vendorname,v.shortname,t.sbu,t.comfam,f.familyname," &
        '                        " {0} || {2} || dm.datatypename as description,t.qty as value from doc.budgetforecast t" &
        '                        " left join doc.datatypemaster dm on dm.id = t.datatypeid " &
        '                        " left join vendor v on v.vendorcode = t.vendorcode " &
        '                        " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
        '                        " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2 " &
        '                        " left join materialmaster mm on mm.cmmf = t.cmmf " &
        '                        " left join family f on f.familyid = t.comfam " &
        '                        " left join doc.sebplatform sf on sf.cmmf = t.cmmf " &
        '                        " where period = '{1}-01-01' and groupid = 1 " &
        '                        " and t.datatypeid in (select * from latest)", String.Format("{0} ", counter), YearEnd, String.Format("{0} ", YearEnd)))
        'counter += 1
        ''Budget Amount
        'sb.Append(String.Format("with latest as ( select distinct datatypeid from doc.budgetforecast where datatypeid in (1,2,8) and period = '{1}-01-01' order by datatypeid desc limit 1)" &
        '                        " select to_char(t.period,'FMYYYYMM'),to_char(t.period,'FMYYYY'),t.fpcp,t.vendorcode,v.vendorname,v.shortname,t.sbu,t.comfam,f.familyname," &
        '                        " {0} || {2} || dm.datatypename as description,t.amount as value from doc.budgetforecast t" &
        '                        " left join doc.datatypemaster dm on dm.id = t.datatypeid " &
        '                        " left join vendor v on v.vendorcode = t.vendorcode " &
        '                        " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
        '                        " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2 " &
        '                        " left join materialmaster mm on mm.cmmf = t.cmmf " &
        '                        " left join family f on f.familyid = t.comfam " &
        '                        " left join doc.sebplatform sf on sf.cmmf = t.cmmf " &
        '                        " where period = '{1}-01-01' and groupid = 1 " &
        '                        " and t.datatypeid in (select * from latest)", String.Format("{0} ", counter), YearEnd, String.Format("{0} ", YearEnd)))
        'counter += 1

        If sb.Length > 0 Then
            sb.Append(" union all ")
        End If

        sb.Append(String.Format("(with latest as ( select distinct datatypeid from doc.budgetforecast where datatypeid in (1,2,8) and period = '{1}-01-01' order by datatypeid desc limit 1)" &
                               " select {3},{6}," &
                               " {0} || {2} || dm.datatypename || ' Volume' as description,sum(t.qty) as value from doc.budgetforecast t" &
                               " left join doc.datatypemaster dm on dm.id = t.datatypeid " &
                               " left join vendor v on v.vendorcode = t.vendorcode " &
                               " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
                               " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2 " &
                               " left join materialmaster mm on mm.cmmf = t.cmmf " &
                               " left join family f on f.familyid = t.comfam " &
                               " left join doc.sebplatform sf on sf.cmmf = t.cmmf " &
                               " where period = '{1}-01-01' and groupid = 1 " &
                               " and t.datatypeid in (select * from latest) {5} group by {4})", String.Format("'{0:00} '", counter), YearEnd, String.Format("'{0} '", YearEnd), GroupBy, String.Format("{0},datatypename", GroupBy), AllFilter, VendorInfo))
        counter += 1

        If sb.Length > 0 Then
            sb.Append(" union all ")
        End If

        'Budget Amount
        sb.Append(String.Format("(with latest as ( select distinct datatypeid from doc.budgetforecast where datatypeid in (1,2,8) and period = '{1}-01-01' order by datatypeid desc limit 1)" &
                                " select {3},{6}," &
                                " {0} || {2} || dm.datatypename || ' Amount (USD)'  as description,sum(t.amount) as value from doc.budgetforecast t" &
                                " left join doc.datatypemaster dm on dm.id = t.datatypeid " &
                                " left join vendor v on v.vendorcode = t.vendorcode " &
                                " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
                                " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2 " &
                                " left join materialmaster mm on mm.cmmf = t.cmmf " &
                                " left join family f on f.familyid = t.comfam " &
                                " left join doc.sebplatform sf on sf.cmmf = t.cmmf " &
                                " where period = '{1}-01-01' and groupid = 1 " &
                                " and t.datatypeid in (select * from latest) {5} group by {4})", String.Format("'{0:00} ' ", counter), YearEnd, String.Format("'{0} '", YearEnd), GroupBy, String.Format("{0},datatypename", GroupBy), AllFilter, VendorInfo))
        counter += 1
        For i = YearStart To YearEnd
            'NQSU
            'sb.Append(String.Format("(with s as (select v.vendorcode,max(period) as period from doc.nqsu  n " &
            '                        " left join vendor v on v.vendorcode = n.vendorcode   " &
            '                        " where Year = {0}" &
            '                        " and samplesizeytd > 0" &
            '                        " group by vendorcode,year order by v.vendorcode,period desc) " &
            '                        " select s.period,year,fpcp,s.vendorcode,v.vendorname,v.shortname,null::text as sbu,null::integer as comfam,null::text as familyname, {1} || 'NQSU' as description," &
            '                        " (criticaldefectytd + majordefectytd) * 1000000 / samplesizeytd::numeric as value" &
            '                        " from s " &
            '                        " left join vendor v on v.vendorcode = s.vendorcode  " &
            '                        " left join doc.nqsu n on n.vendorcode = s.vendorcode and n.period = s.period and samplesizeytd > 0 " &
            '                        ")", i, String.Format("{0} ", counter, GroupBy)))
            'counter += 1
            If sb.Length > 0 Then
                sb.Append(" union all ")
            End If

            sb.Append(String.Format("(with s as (select v.vendorcode,max(period) as period from doc.nqsu  n " &
                                    " left join vendor v on v.vendorcode = n.vendorcode   " &
                                    " where Year = {0} {4}" &
                                    " and samplesizeytd > 0" &
                                    " group by v.vendorcode,year order by v.vendorcode,period desc) " &
                                    " select {2},{3},'{1:00} {0} NQSU' as description," &
                                    " (sum(criticaldefectytd) + sum(majordefectytd)) * 1000000 / sum(samplesizeytd::numeric) as value" &
                                    " from s " &
                                    " left join vendor v on v.vendorcode = s.vendorcode  " &
                                    " left join doc.nqsu n on n.vendorcode = s.vendorcode and n.period = s.period and samplesizeytd > 0  group by {2}" &
                                    ") ", i, counter, GroupBy, VendorInfo, NQSUFilter))
            counter += 1

            If sb.Length > 0 Then
                sb.Append(" union all ")
            End If

            sb.Append(String.Format("(select {2},{4},'{1:00} {0} SSL NET' as description ,  sum(sslnet)/sum(weight) as value from doc.logistics t " &
                                    " left join vendor v on v.vendorcode = t.vendorcode " &
                                    " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
                                    " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2 " &
                                    " where year = {0} {3} {5} group by {2},year order by shortname3)", i, counter, GroupBy, ProductTypeFilter, VendorInfo, LogisticsFilter))
            counter += 1
        Next
        Dim sqlstr = sb.ToString
        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("SupplierListTO-{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 2

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\SupplierListTOTemplate.xltx")
            myreport.Run(Me, e)
        End If
        ProgressReport(5, "Preparing Helper Data. Please wait..")
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)

    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
        Dim oXl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        oXl = owb.Parent
        owb.Worksheets(2).select()


        'check availability data


        Dim osheet = owb.Worksheets(2)
        Dim orange = osheet.Range("A2")
        If osheet.cells(2, 2).text.ToString = "" Then
            Err.Raise(100, Description:="Data not available.")
        End If
        osheet.name = "RawData"
        owb.Names.Add("db", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C1),COUNTA('RawData'!R1))")


        owb.Worksheets(1).select()
        osheet = owb.Worksheets(1)
        osheet.name = "Supplier List TO"
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        osheet.pivottables("PivotTable1").SaveData = True

        osheet.PivotTables("PivotTable1").PivotFields("Sum of value").NumberFormat = "#,##0"
        'override percentage

        For i = 3 To 100
            If osheet.cells(8, i).value = "" Then
                Exit For
            End If
            If InStr(1, osheet.cells(8, i).value, "SSL", CompareMethod.Binary) Then
                Dim mm = osheet.cells(8, i).Address(False, False).ToString.Substring(0, 1)
                osheet.range(mm & ":" & mm).numberformat = "0%"
            End If
        Next

        owb.Worksheets(3).select()
        osheet = owb.Worksheets(3)
        osheet.range("yearstart") = CDate(String.Format("{0}-1-1", YearStart))
        osheet.range("yearend") = CDate(String.Format("{0}-1-1", YearEnd))
        osheet.range("sbu") = AppliedSBUFilter
        osheet.range("family") = AppliedFamilyFilter
        osheet.range("shortname") = AppliedShortNameFilter
        osheet.range("suppliername") = AppliedSupplierNameFilter
        osheet.range("producttype") = AppliedProductTypeFilter
        osheet.name = "Filter"
        'osheet = owb.Worksheets(2)
        'osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        'osheet.pivottables("PivotTable1").SaveData = True
        'owb.RefreshAll()

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
                        ToolStripStatusLabel2.Text = message
                    Case 4
                        runreport(Me, New System.EventArgs)
                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub GetFilter()
        Dim sb As New StringBuilder
        AllFilter.Clear()
        NQSUFilter.Clear()
        LogisticsFilter.Clear()

        AppliedSBUFilter = TextBox1.Text
        AppliedFamilyFilter = TextBox2.Text
        AppliedShortNameFilter = TextBox3.Text
        AppliedSupplierNameFilter = TextBox4.Text
        AppliedProductTypeFilter = TextBox5.Text


        'SBUFilter 
        Dim SBUarr() As String = AppliedSBUFilter.Split(",")
        For i = 0 To SBUarr.Length - 1
            If SBUarr(i).Length > 0 Then
                If sb.Length > 0 Then
                    sb.Append(",")
                End If
                sb.Append(String.Format("'{0}'", SBUarr(i).Trim.ToLower))
            End If            
        Next
        If sb.Length > 0 Then
            SBUFilter = String.Format("and lower(t.sbu) in ({0})", sb.ToString)
            AllFilter.Append(SBUFilter)
            'LogisticsFilter.Append(SBUFilter)
        End If
        sb.Clear()

        'FamilyFilter
        Dim FamilyArr() As String = AppliedFamilyFilter.Split(",")
        For i = 0 To FamilyArr.Length - 1
            If FamilyArr(i).Length > 0 Then
                If sb.Length > 0 Then
                    sb.Append(",")
                End If
                sb.Append(String.Format("'{0}'", FamilyArr(i).Trim.ToLower))
            End If            
        Next
        If sb.Length > 0 Then
            FamilyFilter = String.Format("and lower(f.familyname) in ({0})", sb.ToString)
            AllFilter.Append(FamilyFilter)
        End If
        sb.Clear()

        'ShortnameFilter
        Dim Shortnamearr() As String = AppliedShortNameFilter.Split(",")
        For i = 0 To Shortnamearr.Length - 1
            If Shortnamearr(i).Length > 0 Then
                If sb.Length > 0 Then
                    sb.Append(",")
                End If
                sb.Append(String.Format("'{0}'", Shortnamearr(i).Trim.ToLower))
            End If
           
        Next
        If sb.Length > 0 Then
            ShortNameFilter = String.Format(" and lower(v.shortname3) in ({0})", sb.ToString)
            AllFilter.Append(ShortNameFilter)
            NQSUFilter.Append(ShortNameFilter)
            LogisticsFilter.Append(ShortNameFilter)
        End If
        sb.Clear()

        'SuppliernameShortnameFilter
        Dim SupplierNameArr() As String = AppliedSupplierNameFilter.Split(",")
        For i = 0 To SupplierNameArr.Length - 1
            If SupplierNameArr(i).Length > 0 Then
                If sb.Length > 0 Then
                    sb.Append(",")
                End If
                sb.Append(String.Format("'{0}'", SupplierNameArr(i).Trim.ToLower))
            End If            
        Next
        If sb.Length > 0 Then
            SupplierNameFilter = String.Format("and lower(v.vendorname) in ({0})", sb.ToString)
            AllFilter.Append(SupplierNameFilter)
            NQSUFilter.Append(SupplierNameFilter)
            LogisticsFilter.Append(SupplierNameFilter)
        End If
        sb.Clear()

        'ProductTypeFilter
        Dim ProductTypearr() As String = AppliedProductTypeFilter.Split(",")
        For i = 0 To ProductTypearr.Length - 1
            If ProductTypearr(i).Length > 0 Then
                If sb.Length > 0 Then
                    sb.Append(",")
                End If
                sb.Append(String.Format("'{0}'", ProductTypearr(i).Trim.ToLower))
            End If
           
        Next
        If sb.Length > 0 Then
            ProductTypeFilter = String.Format("and lower(fpcp) in ({0})", sb.ToString)
            AllFilter.Append(ProductTypeFilter)
            NQSUFilter.Append(ProductTypeFilter)
            LogisticsFilter.Append(ProductTypeFilter)
        End If
        sb.Clear()

        'Groupby
        If AppliedSupplierNameFilter.Length > 0 Then
            GroupBy = "v.vendorcode"
            VendorInfo = "v.vendorname"
        End If


    End Sub

    Private Sub FormSupplierListTO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loaddata()
    End Sub

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
        ProgressReport(1, "Preparing Helper Data. Please wait..")
        Dim sqlstr As New StringBuilder
        'SBU
        sqlstr.Append("with t as (select distinct sbu from doc.turnover where not sbu isnull), b as (select distinct sbu from doc.budgetforecast where not sbu isnull)," &
                      " u as (select * from t union all select * from b) select distinct sbu as name from t order by sbu;")
        'Family
        sqlstr.Append("with t as (select distinct familyname from doc.turnover t left join family f on f.familyid = t.comfam where not familyname isnull)," &
                      " b as (select distinct familyname from doc.budgetforecast b left join family f on f.familyid = b.comfam where not sbu isnull), " &
                      " u as (select * from t union all select * from b) select distinct familyname as name from t order by familyname;")
        'Shortname
        sqlstr.Append("with t as (select distinct shortname3 as shortname from doc.turnover t " &
                      " left join vendor v on v.vendorcode = t.vendorcode where not shortname3 isnull), b as (select distinct shortname3 as shortname from doc.budgetforecast b" &
                      " left join vendor v on v.vendorcode = b.vendorcode where not shortname3 isnull), u as (select * from t union all select * from b)" &
                      " select distinct shortname as name from t order by shortname;")
        'Vendorname
        'sqlstr.Append("with t as (select distinct t.vendorcode  from doc.turnover t " &
        '              " left join vendor v on v.vendorcode = t.vendorcode where not vendorname isnull), b as (select distinct b.vendorcode from doc.budgetforecast b" &
        '              " left join vendor v on v.vendorcode = b.vendorcode where not vendorname isnull), u as (select * from t union  select * from b)" &
        '              " select v.vendorcode::text || ' - '  || v.vendorname as name ,v.vendorname from u  left join vendor v on v.vendorcode = u.vendorcode where not vendorname isnull  order by vendorname;")
        sqlstr.Append("select v.vendorcode::text || ' - '  || v.vendorname as name ,v.vendorname::text from vendor v where vendorcode > 10000000 order by v.vendorname;")
        sqlstr.Append("select 'FP' as name union all Select 'CP';")
        Dim mymessage As String = String.Empty
        DS = New DataSet
        If DbAdapter1.TbgetDataSet(sqlstr.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "SBU"
                DS.Tables(1).TableName = "Family"
                DS.Tables(2).TableName = "Shortname"
                DS.Tables(3).TableName = "Vendorname"
                DS.Tables(4).TableName = "ProductType"



            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(1, "Preparing Helper Data. Done.")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonResetSBU.Click, ButtonResetFamily.Click, ButtonResetShortname.Click, ButtonResetSuppliername.Click, ButtonResetProduct.Click

        Dim obj = DirectCast(sender, Button)
        Select Case obj.Name
            Case "ButtonResetSBU"
                TextBox1.Text = ""
            Case "ButtonResetFamily"
                TextBox2.Text = ""
            Case "ButtonResetShortname"
                TextBox3.Text = ""
            Case "ButtonResetSuppliername"
                TextBox4.Text = ""
            Case "ButtonResetProduct"
                TextBox5.Text = ""
        End Select
    End Sub

    Private Sub ButtonHSBU_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonHSBU.Click, ButtonHFamily.Click, ButtonHShortName.Click, ButtonHVendorName.Click, ButtonHProductType.Click

        If myThread.IsAlive Then
            MessageBox.Show("Please wait until the current process is finished.")
            Exit Sub
        End If
        Dim obj = DirectCast(sender, Button)
        Dim myhelper As New FormHelperMulti()
        myhelper.Key = 0
        Dim BS As New BindingSource
        Select Case obj.Name
            Case "ButtonHSBU"
                BS.DataSource = DS.Tables(0)
            Case "ButtonHFamily"
                BS.DataSource = DS.Tables(1)
            Case "ButtonHShortName"
                BS.DataSource = DS.Tables(2)
            Case "ButtonHVendorName"
                BS.DataSource = DS.Tables(3)
                myhelper.Key = 1
            Case "ButtonHProductType"
                BS.DataSource = DS.Tables(4)
        End Select
        'Dim myhelper As New FormHelperMulti(BS)
        myhelper.bs = BS

        myhelper.ShowDialog()
        Select Case obj.Name
            Case "ButtonHSBU"
                TextBox1.Text = TextBox1.Text & IIf(TextBox1.Text = "", "", ",") & myhelper.SelectedResult
            Case "ButtonHFamily"
                TextBox2.Text = TextBox2.Text & IIf(TextBox2.Text = "", "", ",") & myhelper.SelectedResult
            Case "ButtonHShortName"
                TextBox3.Text = TextBox3.Text & IIf(TextBox3.Text = "", "", ",") & myhelper.SelectedResult
            Case "ButtonHVendorName"
                TextBox4.Text = TextBox4.Text & IIf(TextBox4.Text = "", "", ",") & myhelper.SelectedResult
            Case "ButtonHProductType"
                TextBox5.Text = TextBox5.Text & IIf(TextBox5.Text = "", "", ",") & myhelper.SelectedResult
        End Select
    End Sub

    
End Class