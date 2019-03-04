Imports System.Text

Public Class BudgetForecastAdapter
    Inherits BaseAdapter
    Implements IAdapter

    Private sb As StringBuilder
    Private vendorcode As Long
    Private shortname As String
    Public latestBudgetPeriod As Date
    Public latestTOPeriod As Date

    Public Enum DataTypeEnum
        BudgetEnum = 1
        Forecast4Plus8 = 2
        TOEnum = 3
        Forecast8Plus4 = 8
    End Enum

    Public Sub New(ByVal vendorcode As Long, ByVal shortname As String)
        MyBase.New()
        Me.vendorcode = vendorcode
        Me.shortname = shortname
    End Sub


    Public Function loaddata() As Boolean Implements IAdapter.loaddata
        Return True
    End Function

    Public Function save() As Boolean Implements IAdapter.save
        Return True
    End Function

    Public Function loadData(ByVal criteria As String) As Boolean
        Dim myret As Boolean = False
        sb = New StringBuilder
        DS = New DataSet
        Dim myfilter As String
        Dim mygroup As String
        If vendorcode = 0 Then
            myfilter = String.Format("and shortname = ''{0}''", shortname)
            mygroup = "group by v.shortname"
        Else
            myfilter = String.Format("and v.vendorcode = {0}", vendorcode)
            mygroup = "group by v.vendorcode"
        End If

        latestBudgetPeriod = getLatestPeriod(DataTypeEnum.BudgetEnum)
        latestTOPeriod = getLatestPeriod(DataTypeEnum.TOEnum)

        Dim labelBudget As String = getLabelBudget()
        Dim LabelTo As String = getLabelTO()


        Dim fieldlist As String = String.Format("vendorcode bigint,{0}", labelBudget)
        Dim TOFieldList As String = String.Format("vendorcode bigint,{0}", LabelTo)

        criteria = criteria.Replace("tu", "bf")
        '0 add criteria
        sb.Append(GetSupplierTotal(myfilter, "", mygroup, fieldlist, latestBudgetPeriod, latestTOPeriod))
        '1 add criteria
        sb.Append(GetSupplierTODetail(myfilter, "", TOFieldList, latestTOPeriod))
        '2
        sb.Append(GetSupplierTotal(myfilter, criteria, mygroup, fieldlist, latestBudgetPeriod, latestTOPeriod))
        '3
        sb.Append(GetSupplierTODetail(myfilter, criteria, TOFieldList, latestTOPeriod))
        'SBU
        fieldlist = String.Format("sbuname character varying,{0}", labelBudget)
        TOFieldList = String.Format("sbuname character varying,{0}", LabelTo)
        '4 add criteria
        sb.Append(GetSupplierTotalGroupby(myfilter, "", mygroup & ",bf.sbu", fieldlist, latestBudgetPeriod, latestTOPeriod, "bf.sbu"))
        '5 add criteria
        sb.Append(GetSupplierTODetailGroupby(myfilter, "", TOFieldList, latestTOPeriod, "bf.sbu"))

        'family
        fieldlist = String.Format("familyname character varying,{0}", labelBudget)
        TOFieldList = String.Format("familyname character varying,{0}", LabelTo)
        '6 add criteria
        sb.Append(GetSupplierTotalGroupby(myfilter, "", mygroup & ",f.familyname", fieldlist, latestBudgetPeriod, latestTOPeriod, "f.familyname::character varying"))
        '7 add criteria
        sb.Append(GetSupplierTODetailGroupby(myfilter, "", TOFieldList, latestTOPeriod, "f.familyname::character varying"))
        '
        'brand
        fieldlist = String.Format("brand character varying,{0}", labelBudget)
        TOFieldList = String.Format("brand character varying,{0}", LabelTo)
        '8 add criteria
        sb.Append(GetSupplierTotalGroupby(myfilter, "", mygroup & ",bf.brand", fieldlist, latestBudgetPeriod, latestTOPeriod, "bf.brand"))
        '9 add criteria
        sb.Append(GetSupplierTODetailGroupby(myfilter, "", TOFieldList, latestTOPeriod, "bf.brand"))

        'brand
        fieldlist = String.Format("rangedesc character varying,{0}", labelBudget)
        TOFieldList = String.Format("rangedesc character varying,{0}", LabelTo)
        '10 add criteria
        sb.Append(GetSupplierTotalGroupby(myfilter, "", mygroup & ",r.rangedesc", fieldlist, latestBudgetPeriod, latestTOPeriod, "r.rangedesc"))
        '11 add criteria
        sb.Append(GetSupplierTODetailGroupby(myfilter, "", TOFieldList, latestTOPeriod, "r.rangedesc"))

        DS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString, DS) Then
            myret = True        

        End If


        Return myret
    End Function

    Public Function getLabelBudget() As String
        Dim mystring As String = String.Empty
        Dim mysb As New StringBuilder
        Try
            Dim sqlstr = "select dtl.datatypelabel from doc.datatypelabel dtl where not lineorder isnull order by lineorder"
            Dim DS As New DataSet
            If DbAdapter1.TbgetDataSet(sqlstr, DS) Then
                For Each dr As DataRow In DS.Tables(0).Rows
                    If mysb.Length > 0 Then
                        mysb.Append(",")
                    End If
                    mysb.Append(String.Format("""{0}"" numeric", dr.Item("datatypelabel")))
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return mysb.ToString
    End Function

    Public Function getLabelTO() As String
        Dim mystring As String = String.Empty
        Dim mysb As New StringBuilder
        Try
            Dim sqlstr = "(select datatypename from doc.datatypemaster" &
                " where groupid = 2 order by amountlineorder)" &
                " union all" &
                " (select datatypename from doc.datatypemaster" &
                " where(groupid = 2)" &
                " and not qtylineorder isnull" &
                " order by qtylineorder)"
            Dim DS As New DataSet
            Dim type As String = "Amount"
            If DbAdapter1.TbgetDataSet(sqlstr, DS) Then
                For Each dr As DataRow In DS.Tables(0).Rows

                    If dr.Item("datatypename") = "Spacer" Then
                        dr.Item("datatypename") = ""
                        type = ""
                    ElseIf type = "" Then
                        type = "Qty"
                    End If
                    If mysb.Length > 0 Then
                        mysb.Append(",")
                    End If
                    mysb.Append(String.Format("""{0} {1}"" numeric", dr.Item("datatypename"), type))
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return mysb.ToString
    End Function

    Public Function getLatestPeriod(ByVal datatype As Integer) As Date
        Dim mystring As String = String.Empty
        Dim mysb As New StringBuilder
        Dim ra As Object = Nothing
        Try
            Dim sqlstr = String.Format("select period from doc.budgetforecast bf" &
                         " where(datatypeid = {0}) order by period desc limit 1", datatype)
            Dim DS As New DataSet

            DbAdapter1.ExecuteScalar(sqlstr, ra)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return CDate(ra).Date
    End Function

    Private Function GetSupplierTotal(ByVal myfilter As String, ByVal criteria As String, ByVal mygroup As String, ByVal fieldlist As String, ByVal latestBudgetPeriod As Date, ByVal latestTOPeriod As Date) As String
        Dim sb As New StringBuilder

        'sb.Append(String.Format("select * from crosstab('with rs as((select max(bf.vendorcode) as vendorcode,(select lineorder from doc.datatypelabel where datatypelabelid = ''TO Amount'' )as type, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join sbusap s on s.sbuid = mm.sbu" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " where period = ''{5:yyyy-MM-dd}'' {0} " &
        '          " and groupid = 2" &
        '          " {1} {2})" &
        '          " union all (select max(bf.vendorcode) as vendorcode,(select lineorder from  doc.datatypelabel where datatypelabelid = ''TO Qty''  ), sum(qty)::numeric as qty from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid " &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join sbusap s on s.sbuid = mm.sbu" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " where period = ''{5:yyyy-MM-dd}'' {0}" &
        '          " and groupid = 2" &
        '          " {1} {2})" &
        '          " union all" &
        '          " (select max(bf.vendorcode) as vendorcode,amountlineorder, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " where period = ''{4:yyyy-MM-dd}'' {0} {1} " &
        '          " and groupid = 1" &
        '          " group by dm.amountlineorder)" &
        '          " union all" &
        '          " (select max(bf.vendorcode) as vendorcode,qtylineorder, sum(qty)::numeric as qty  from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " where period = ''{4:yyyy-MM-dd}'' {0} {1} " &
        '          " and groupid = 1" &
        '          " group by dm.qtylineorder))" &
        '          " select * from rs','select lineorder from doc.datatypelabel where not lineorder isnull order by lineorder') as({3});", myfilter, criteria, mygroup, fieldlist, latestBudgetPeriod, latestTOPeriod))


        'sb.Append(String.Format("select * from crosstab('with rs as((select 1 as vendorcode,(select lineorder from doc.datatypelabel where datatypelabelid = ''TO Amount'' )as type, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join sbusap s on s.sbuid = mm.sbu" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " where period = ''{5:yyyy-MM-dd}'' {0} " &
        '          " and groupid = 2" &
        '          " {1} {2})" &
        '          " union all (select 1 as vendorcode,(select lineorder from  doc.datatypelabel where datatypelabelid = ''TO Qty''  ), sum(qty)::numeric as qty from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid " &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join sbusap s on s.sbuid = mm.sbu" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " where period = ''{5:yyyy-MM-dd}'' {0}" &
        '          " and groupid = 2" &
        '          " {1} {2})" &
        '          " union all" &
        '          " (select 1 as vendorcode,amountlineorder, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " where period = ''{4:yyyy-MM-dd}'' {0} {1} " &
        '          " and groupid = 1" &
        '          " group by dm.amountlineorder)" &
        '          " union all" &
        '          " (select 1 as vendorcode,qtylineorder, sum(qty)::numeric as qty  from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " where period = ''{4:yyyy-MM-dd}'' {0} {1} " &
        '          " and groupid = 1" &
        '          " group by dm.qtylineorder))" &
        '          " select * from rs','select lineorder from doc.datatypelabel where not lineorder isnull order by lineorder') as({3});", myfilter, criteria, mygroup, fieldlist, latestBudgetPeriod, latestTOPeriod))

        sb.Append(String.Format("select * from crosstab('with rs as((select 1 as vendorcode,(select lineorder from doc.datatypelabel where datatypelabelid = ''TO Amount'' )as type, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
                 " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
                 " left join vendor v on v.vendorcode = bf.vendorcode" &
                 " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                 " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                 " left join materialmaster mm on mm.cmmf = bf.cmmf" &
                 " left join sbusap s on s.sbuid = mm.sbu" &
                 " left join family f on f.familyid = bf.comfam" &
                 " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
                 " where period = ''{5:yyyy-MM-dd}'' {0} " &
                 " and groupid = 2" &
                 " {1} {2})" &
                 " union all (select 1 as vendorcode,(select lineorder from  doc.datatypelabel where datatypelabelid = ''TO Qty''  ), sum(qty)::numeric as qty from doc.budgetforecast bf" &
                 " left join doc.datatypemaster dm on dm.id = bf.datatypeid " &
                 " left join vendor v on v.vendorcode = bf.vendorcode" &
                 " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                 " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                 " left join materialmaster mm on mm.cmmf = bf.cmmf" &
                 " left join sbusap s on s.sbuid = mm.sbu" &
                 " left join family f on f.familyid = bf.comfam" &
                 " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
                 " where period = ''{5:yyyy-MM-dd}'' {0}" &
                 " and groupid = 2" &
                 " {1} {2})" &
                 " union all" &
                 " (select 1 as vendorcode,amountlineorder, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
                 " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
                 " left join vendor v on v.vendorcode = bf.vendorcode" &
                 " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                 " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                 " left join materialmaster mm on mm.cmmf = bf.cmmf" &
                 " left join family f on f.familyid = bf.comfam" &
                 " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
                 " where period = ''{4:yyyy-MM-dd}'' {0} {1} " &
                 " and groupid = 1" &
                 " group by dm.amountlineorder)" &
                 " union all" &
                 " (select 1 as vendorcode,qtylineorder, sum(qty)::numeric as qty  from doc.budgetforecast bf" &
                 " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
                 " left join vendor v on v.vendorcode = bf.vendorcode" &
                 " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                 " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                 " left join materialmaster mm on mm.cmmf = bf.cmmf" &
                 " left join family f on f.familyid = bf.comfam" &
                 " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
                 " where period = ''{4:yyyy-MM-dd}'' {0} {1} " &
                 " and groupid = 1" &
                 " group by dm.qtylineorder))" &
                 " select * from rs','select lineorder from doc.datatypelabel where not lineorder isnull order by lineorder') as({3});", myfilter, criteria, mygroup, fieldlist, latestBudgetPeriod, latestTOPeriod))
        Return sb.ToString
    End Function

    Private Function GetSupplierTODetail(ByVal myfilter As String, ByVal criteria As String, ByVal TOFieldList As String, ByVal latestTOPeriod As Date) As String
        Dim sb As New StringBuilder
        'sb.Append(String.Format("select * from crosstab('with rs as((select max(bf.vendorcode) as vendorcode,amountlineorder, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " where period = ''{3:yyyy-MM-dd}'' {0} " &
        '          " and groupid = 2" &
        '          " {1} group by amountlineorder)" &
        '          " union all (select max(bf.vendorcode) as vendorcode,qtylineorder, sum(qty)::numeric as qty from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid " &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " where period = ''{3:yyyy-MM-dd}'' {0} " &
        '          " and groupid = 2" &
        '          " {1} group by qtylineorder))" &
        '          " select * from rs','select m from generate_series(1,{4})m') as({2});", myfilter, criteria, TOFieldList, latestTOPeriod, TOFieldList.Split(",").Length - 1))

        'sb.Append(String.Format("select * from crosstab('with rs as((select 1 as vendorcode,amountlineorder, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " where period = ''{3:yyyy-MM-dd}'' {0} " &
        '          " and groupid = 2" &
        '          " {1} group by amountlineorder)" &
        '          " union all (select 1 as vendorcode,qtylineorder, sum(qty)::numeric as qty from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid " &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " where period = ''{3:yyyy-MM-dd}'' {0} " &
        '          " and groupid = 2" &
        '          " {1} group by qtylineorder))" &
        '          " select * from rs','select m from generate_series(1,{4})m') as({2});", myfilter, criteria, TOFieldList, latestTOPeriod, TOFieldList.Split(",").Length - 1))

        sb.Append(String.Format("select * from crosstab('with rs as((select 1 as vendorcode,amountlineorder, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
                  " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
                  " left join vendor v on v.vendorcode = bf.vendorcode" &
                  " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                 " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                  " left join materialmaster mm on mm.cmmf = bf.cmmf" &
                  " left join family f on f.familyid = bf.comfam" &
                  " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
                  " where period = ''{3:yyyy-MM-dd}'' {0} " &
                  " and groupid = 2" &
                  " {1} group by amountlineorder)" &
                  " union all (select 1 as vendorcode,qtylineorder, sum(qty)::numeric as qty from doc.budgetforecast bf" &
                  " left join doc.datatypemaster dm on dm.id = bf.datatypeid " &
                  " left join vendor v on v.vendorcode = bf.vendorcode" &
                  " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                 " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                  " left join materialmaster mm on mm.cmmf = bf.cmmf" &
                  " left join family f on f.familyid = bf.comfam" &
                  " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
                  " where period = ''{3:yyyy-MM-dd}'' {0} " &
                  " and groupid = 2" &
                  " {1} group by qtylineorder))" &
                  " select * from rs','select m from generate_series(1,{4})m') as({2});", myfilter, criteria, TOFieldList, latestTOPeriod, TOFieldList.Split(",").Length - 1))


        Return sb.ToString
    End Function

    Private Function GetSupplierTotalGroupby(ByVal myfilter As String, ByVal criteria As String, ByVal mygroup As String, ByVal fieldlist As String, ByVal latestBudgetPeriod As Date, ByVal latestTOPeriod As Date, ByVal fieldname As String) As String
        Dim sb As New StringBuilder
        Dim myfieldname() = fieldname.Split(".")
        'sb.Append(String.Format("select * from crosstab('with rs as((select {6},(select lineorder from doc.datatypelabel where datatypelabelid = ''TO Amount'' )as type, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join sbusap s on s.sbuid = mm.sbu" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " left join range r on r.range = bf.range" &
        '          " where period = ''{5:yyyy-MM-dd}'' {0} " &
        '          " and groupid = 2" &
        '          " {1} {2})" &
        '          " union all (select {6},(select lineorder from  doc.datatypelabel where datatypelabelid = ''TO Qty''  ), sum(qty)::numeric as qty from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid " &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join sbusap s on s.sbuid = mm.sbu" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " left join range r on r.range = bf.range" &
        '          " where period = ''{5:yyyy-MM-dd}'' {0}" &
        '          " and groupid = 2" &
        '          " {1} {2})" &
        '          " union all" &
        '          " (select {6},amountlineorder, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " left join range r on r.range = bf.range" &
        '          " where period = ''{4:yyyy-MM-dd}'' {0} {1} " &
        '          " and groupid = 1" &
        '          " group by dm.amountlineorder,{6}" &
        '          " order by {6},dm.amountlineorder)" &
        '          " union all" &
        '          " (select {6},qtylineorder, sum(qty)::numeric as qty  from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " left join range r on r.range = bf.range" &
        '          " where period = ''{4:yyyy-MM-dd}'' {0} {1} " &
        '          " and groupid = 1" &
        '          " group by dm.qtylineorder,{6}" &
        '          " order by {6},dm.qtylineorder))" &
        '          " select * from rs order by {7}','select lineorder from doc.datatypelabel where not lineorder isnull order by lineorder') as({3});", myfilter, criteria, mygroup, fieldlist, latestBudgetPeriod, latestTOPeriod, fieldname, myfieldname(1)))

        sb.Append(String.Format("select * from crosstab('with rs as((select {6},(select lineorder from doc.datatypelabel where datatypelabelid = ''TO Amount'' )as type, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
                  " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
                  " left join vendor v on v.vendorcode = bf.vendorcode" &
                  " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                  " left join materialmaster mm on mm.cmmf = bf.cmmf" &
                  " left join sbusap s on s.sbuid = mm.sbu" &
                  " left join family f on f.familyid = bf.comfam" &
                  " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
                  " left join range r on r.range = bf.range" &
                  " where period = ''{5:yyyy-MM-dd}'' {0} " &
                  " and groupid = 2" &
                  " {1} {2})" &
                  " union all (select {6},(select lineorder from  doc.datatypelabel where datatypelabelid = ''TO Qty''  ), sum(qty)::numeric as qty from doc.budgetforecast bf" &
                  " left join doc.datatypemaster dm on dm.id = bf.datatypeid " &
                  " left join vendor v on v.vendorcode = bf.vendorcode" &
                  " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                  " left join materialmaster mm on mm.cmmf = bf.cmmf" &
                  " left join sbusap s on s.sbuid = mm.sbu" &
                  " left join family f on f.familyid = bf.comfam" &
                  " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
                  " left join range r on r.range = bf.range" &
                  " where period = ''{5:yyyy-MM-dd}'' {0}" &
                  " and groupid = 2" &
                  " {1} {2})" &
                  " union all" &
                  " (select {6},amountlineorder, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
                  " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
                  " left join vendor v on v.vendorcode = bf.vendorcode" &
                  " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                  " left join materialmaster mm on mm.cmmf = bf.cmmf" &
                  " left join family f on f.familyid = bf.comfam" &
                  " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
                  " left join range r on r.range = bf.range" &
                  " where period = ''{4:yyyy-MM-dd}'' {0} {1} " &
                  " and groupid = 1" &
                  " group by dm.amountlineorder,{6}" &
                  " order by {6},dm.amountlineorder)" &
                  " union all" &
                  " (select {6},qtylineorder, sum(qty)::numeric as qty  from doc.budgetforecast bf" &
                  " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
                  " left join vendor v on v.vendorcode = bf.vendorcode" &
                  " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                  " left join materialmaster mm on mm.cmmf = bf.cmmf" &
                  " left join family f on f.familyid = bf.comfam" &
                  " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
                  " left join range r on r.range = bf.range" &
                  " where period = ''{4:yyyy-MM-dd}'' {0} {1} " &
                  " and groupid = 1" &
                  " group by dm.qtylineorder,{6}" &
                  " order by {6},dm.qtylineorder))" &
                  " select * from rs order by {7}','select lineorder from doc.datatypelabel where not lineorder isnull order by lineorder') as({3});", myfilter, criteria, mygroup, fieldlist, latestBudgetPeriod, latestTOPeriod, fieldname, myfieldname(1)))

        Return sb.ToString
    End Function
    Private Function GetSupplierTODetailGroupby(ByVal myfilter As String, ByVal criteria As String, ByVal TOFieldList As String, ByVal latestTOPeriod As Date, ByVal FieldName As String) As String
        Dim sb As New StringBuilder
        Dim myfieldname() = FieldName.Split(".")
        'sb.Append(String.Format("select * from crosstab('with rs as((select {4},amountlineorder, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " left join range r on r.range = bf.range" &
        '          " where period = ''{3:yyyy-MM-dd}'' {0} " &
        '          " and groupid = 2" &
        '          " {1} group by amountlineorder,{4} order by {4},amountlineorder)" &
        '          " union all (select {4},qtylineorder, sum(qty)::numeric as qty from doc.budgetforecast bf" &
        '          " left join doc.datatypemaster dm on dm.id = bf.datatypeid " &
        '          " left join vendor v on v.vendorcode = bf.vendorcode" &
        '          " left join materialmaster mm on mm.cmmf = bf.cmmf" &
        '          " left join family f on f.familyid = bf.comfam" &
        '          " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
        '          " left join range r on r.range = bf.range" &
        '          " where period = ''{3:yyyy-MM-dd}'' {0} " &
        '          " and groupid = 2" &
        '          " {1} group by qtylineorder,{4} order by {4},qtylineorder))" &
        '          " select * from rs order by {5}','select m from generate_series(1,{6})m') as({2});", myfilter, criteria, TOFieldList, latestTOPeriod, FieldName, myfieldname(1), TOFieldList.Split(",").Length - 1))

        sb.Append(String.Format("select * from crosstab('with rs as((select {4},amountlineorder, sum(amount)::numeric as amount  from doc.budgetforecast bf" &
                  " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
                  " left join vendor v on v.vendorcode = bf.vendorcode" &
                  " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                  " left join materialmaster mm on mm.cmmf = bf.cmmf" &
                  " left join family f on f.familyid = bf.comfam" &
                  " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
                  " left join range r on r.range = bf.range" &
                  " where period = ''{3:yyyy-MM-dd}'' {0} " &
                  " and groupid = 2" &
                  " {1} group by amountlineorder,{4} order by {4},amountlineorder)" &
                  " union all (select {4},qtylineorder, sum(qty)::numeric as qty from doc.budgetforecast bf" &
                  " left join doc.datatypemaster dm on dm.id = bf.datatypeid " &
                  " left join vendor v on v.vendorcode = bf.vendorcode" &
                  " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                  " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                  " left join materialmaster mm on mm.cmmf = bf.cmmf" &
                  " left join family f on f.familyid = bf.comfam" &
                  " left join doc.sebplatform sf on sf.cmmf = bf.cmmf" &
                  " left join range r on r.range = bf.range" &
                  " where period = ''{3:yyyy-MM-dd}'' {0} " &
                  " and groupid = 2" &
                  " {1} group by qtylineorder,{4} order by {4},qtylineorder))" &
                  " select * from rs order by {5}','select m from generate_series(1,{6})m') as({2});", myfilter, criteria, TOFieldList, latestTOPeriod, FieldName, myfieldname(1), TOFieldList.Split(",").Length - 1))

        '" select * from rs order by {5}','select lineorder from doc.datatypelabel where not lineorder isnull order by lineorder') as({2});", myfilter, criteria, TOFieldList, latestTOPeriod, FieldName, myfieldname(1)))
        Return sb.ToString
    End Function

    Public Function GetBudgetForecast(ByVal period As Date, ByVal vendorcode As Long, ByVal DataTypeEnum As DataTypeEnum) As Decimal
        Dim sqlstr = String.Format("select  v.vendorcode,amountlineorder, sum(amount)::numeric as amount  from doc.budgetforecast bf " &
                     " left join doc.datatypemaster dm on dm.id = bf.datatypeid" &
                     " left join vendor v on v.vendorcode = bf.vendorcode " &
                     " where period = '{0:yyyy-MM-yy}' and v.vendorcode = {1} and  datatypeid = {2}", period, vendorcode, DataTypeEnum)
        Dim ra As Decimal
        DbAdapter1.ExecuteScalar(sqlstr, ra)
        Return ra
    End Function

    Public Function GetTurnover(ByVal yearref As Integer, ByVal vendorcode As Long) As Decimal
        Dim sqlstr = String.Format("select sum(amount) as amount from doc.turnover d " &
                     " left join vendor v on v.vendorcode = d.vendorcode " &
                     " where(Year = {0} and v.vendorcode = {1} )" &
                     " group by year;", yearref, vendorcode)
        Dim ra As Decimal
        DbAdapter1.ExecuteScalar(sqlstr, ra)
        Return ra
    End Function
End Class
