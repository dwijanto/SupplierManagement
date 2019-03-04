Imports System.Threading
Imports Microsoft.Office.Interop
Imports SupplierManagement.SharedClass
Imports System.Text
Imports SupplierManagement.PublicClass
Public Class FormReportTOSC
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Private DS As DataSet
    Private SqlStr As String

    Public Property sify As Decimal
    Public Property sify1 As Decimal
    Public Property sify2 As Decimal
    Public Property sify3 As Decimal
    Public Property sify4 As Decimal

    Public Property TOSEBQTY As Decimal
    Public Property TOSEBQTY1 As Decimal
    Public Property TOSEBQTY2 As Decimal
    Public Property TOSEBQTY3 As Decimal
    Public Property TOSEBQTY4 As Decimal

    Public Property TOSEBAMT As Decimal
    Public Property TOSEBAMT1 As Decimal
    Public Property TOSEBAMT2 As Decimal
    Public Property TOSEBAMT3 As Decimal
    Public Property TOSEBAMT4 As Decimal

    Public Property FilterQty As Decimal
    Public Property FilterQty1 As Decimal
    Public Property FilterQty2 As Decimal
    Public Property FilterQty3 As Decimal
    Public Property FilterQty4 As Decimal

    Public Property FilterAmt As Decimal
    Public Property FilterAmt1 As Decimal
    Public Property FilterAmt2 As Decimal
    Public Property FilterAmt3 As Decimal
    Public Property FilterAmt4 As Decimal

    Public Property nqsu As String
    Public Property nqsu1 As String
    Public Property nqsu2 As String
    Public Property nqsu3 As String
    Public Property nqsu4 As String

    Public Property scsasl As Decimal
    Public Property scsasl1 As Decimal
    Public Property scsasl2 As Decimal
    Public Property scsasl3 As Decimal
    Public Property scsasl4 As Decimal

    Public Property scssl As Decimal = Nothing
    Public Property scssl1 As Decimal
    Public Property scssl2 As Decimal
    Public Property scssl3 As Decimal
    Public Property scssl4 As Decimal


    Public Property scsslnet As Decimal = Nothing
    Public Property scsslnet1 As Decimal
    Public Property scsslnet2 As Decimal
    Public Property scsslnet3 As Decimal
    Public Property scsslnet4 As Decimal

    Public Property sclt As Decimal
    Public Property sclt1 As Decimal
    Public Property sclt2 As Decimal
    Public Property sclt3 As Decimal
    Public Property sclt4 As Decimal

    Public Property scno As Decimal
    Public Property scno1 As Decimal
    Public Property scno2 As Decimal
    Public Property scno3 As Decimal
    Public Property scno4 As Decimal

    Public Property scsh As Decimal
    Public Property scsh1 As Decimal
    Public Property scsh2 As Decimal
    Public Property scsh3 As Decimal
    Public Property scsh4 As Decimal

    Public Property piavg As Decimal
    Public Property piavg1 As Decimal
    Public Property piavg2 As Decimal
    Public Property piavg3 As Decimal
    Public Property piavg4 As Decimal

    Public Property pilkp As Decimal
    Public Property pilkp1 As Decimal
    Public Property pilkp2 As Decimal
    Public Property pilkp3 As Decimal
    Public Property pilkp4 As Decimal

    Public Property pistd As Decimal
    Public Property pistd1 As Decimal
    Public Property pistd2 As Decimal
    Public Property pistd3 As Decimal
    Public Property pistd4 As Decimal

    Public Property pvstd As Decimal
    Public Property pvstd1 As Decimal
    Public Property pvstd2 As Decimal
    Public Property pvstd3 As Decimal
    Public Property pvstd4 As Decimal

    Public Property pdrp14 As String
    Public Property pdrp141 As String
    Public Property pdrp142 As String
    Public Property pdrp143 As String
    Public Property pdrp144 As String

    Public Property pdrp5 As String
    Public Property pdrp51 As String
    Public Property pdrp52 As String
    Public Property pdrp53 As String
    Public Property pdrp54 As String

    Public Property pdramp As String
    Public Property pdramp1 As String
    Public Property pdramp2 As String
    Public Property pdramp3 As String
    Public Property pdramp4 As String

    Public Property pdpot As String
    Public Property pdpot1 As String
    Public Property pdpot2 As String
    Public Property pdpot3 As String
    Public Property pdpot4 As String

    Public CurrentMonth As Date
    Public frontshortname As String
    Public frontsuppliername As String
    Public frontsuppliercode As String
    Public frontproductype As String
    Public frontsbu As String
    Public frontfamily As String
    Public frontbrand As String
    Public frontrangecode As String
    Public frontcmmf As String
    Public frontcommercialcode As String
    Public frontsebasiaplatform As String
    Public frontvendorstatus As String
    Public filterApplied As String

    Public ToolingHDBS As BindingSource
    Public ToolingDetailsALLBS As BindingSource
    Public SAPBS As BindingSource
    Public ContractBS As BindingSource
    Public CitiProgramBS As BindingSource
    Public AuthletterBS As BindingSource
    Public ProjectSpecBS As BindingSource
    Public VendorCurrencyBS As BindingSource

    Public PCharterBS As BindingSource
    Public SocialAuditBS As BindingSource

    Public PanelStatusBS As BindingSource
    Public SupplierTechnologyBS As BindingSource
    Public SupplierPanelHistory As BindingSource

    Public VendorBS As BindingSource
    Public FactoryBS As BindingSource
    Public ContactBS As BindingSource

    Public SIFIDBS As BindingSource
    Public SIFBS As BindingSource
    Public IDBS As BindingSource

    Public PriceCMMFBS As BindingSource
    Public CMMFFilter As String
    Public LKPFiter As String

    Public Sub New(ByVal SQLStr As String)
        InitializeComponent()

        Me.SqlStr = SQLStr
    End Sub

    Public Sub New(ByVal SQLStr As String, ByVal DS As DataSet)
        InitializeComponent()

        Me.SqlStr = SQLStr
        Me.DS = DS
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        runreport(sender, e)


    End Sub
    Private Sub runreport(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim mymessage As String = String.Empty

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("DashboardTurnoverScorecard{1}{0:yyyyMMdd}.xlsx", Date.Today, frontshortname)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 13

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\DashboardTurnoverScorecard.xltx")
            myreport.Run(Me, e)
        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
        Dim oXl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        oXl = owb.Parent

        owb.Names.Add("torawdata", RefersToR1C1:="=OFFSET('TurnoverRawData'!R1C1,0,0,COUNTA('TurnoverRawData'!C1),COUNTA('TurnoverRawData'!R1))")


        ' .Add("rawdata", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C2),COUNTA('RawData'!R1))")
        owb.Worksheets(1).select()

        Dim osheet = owb.Worksheets(1)
        osheet.range("frontshortname") = frontshortname
        osheet.range("frontsuppliername") = frontsuppliername
        osheet.range("frontsuppliercode") = frontsuppliercode
        osheet.range("frontproductype") = frontproductype
        osheet.range("frontsbu") = frontsbu
        osheet.range("frontfamily") = frontfamily
        osheet.range("frontbrand") = frontbrand
        osheet.range("frontrangecode") = frontrangecode
        osheet.range("frontcmmf") = frontcmmf
        osheet.range("frontcommercialcode") = frontcommercialcode
        osheet.range("frontsebasiaplatform") = frontsebasiaplatform
        osheet.range("frontvendorstatus") = frontvendorstatus
        osheet.range("frontyearperiod") = String.Format("{0:dd-MMM-yyyy}", CurrentMonth)
        osheet.range("issuedate") = String.Format("{0:dd-MMM-yyyy}", Date.Today)

        owb.Worksheets(2).select()
        osheet = owb.Worksheets(2)
        osheet.range("filterapplied") = filterApplied
        osheet.range("sify") = sify
        osheet.range("sify1") = sify1
        osheet.range("sify2") = sify2
        osheet.range("sify3") = sify3
        osheet.range("sify4") = sify4
        osheet.range("TOSEBQTY") = TOSEBQTY
        osheet.range("TOSEBQTY1") = TOSEBQTY1
        osheet.range("TOSEBQTY2") = TOSEBQTY2
        osheet.range("TOSEBQTY3") = TOSEBQTY3
        osheet.range("TOSEBQTY4") = TOSEBQTY4
        osheet.range("TOSEBAMT") = TOSEBAMT
        osheet.range("TOSEBAMT1") = TOSEBAMT1
        osheet.range("TOSEBAMT2") = TOSEBAMT2
        osheet.range("TOSEBAMT3") = TOSEBAMT3
        osheet.range("TOSEBAMT4") = TOSEBAMT4
        osheet.range("FilterQty") = FilterQty
        osheet.range("FilterQty1") = FilterQty1
        osheet.range("FilterQty2") = FilterQty2
        osheet.range("FilterQty3") = FilterQty3
        osheet.range("FilterQty4") = FilterQty4
        osheet.range("FilterAmt") = FilterAmt
        osheet.range("FilterAmt1") = FilterAmt1
        osheet.range("FilterAmt2") = FilterAmt2
        osheet.range("FilterAmt3") = FilterAmt3
        osheet.range("FilterAmt4") = FilterAmt4





        If Not IsNothing(DS) Then
            owb.Worksheets(3).select()
            osheet = owb.Worksheets(3)
            'Supplier Total
            showTotal(osheet, DS.Tables(0), DS.Tables(1), 5, 3)
            'By Filter
            showTotal(osheet, DS.Tables(2), DS.Tables(3), 8, 3)
            showDetail(osheet, "sbuname", "SBU", 10, DS.Tables(4), DS.Tables(5), 13, 3)
            showDetail(osheet, "familyname", "Family", 20, DS.Tables(6), DS.Tables(7), 25, 3)
            showDetail(osheet, "brand", "Brand", 20, DS.Tables(8), DS.Tables(9), 47, 3)
            showDetail(osheet, "rangedesc", "Range", 20, DS.Tables(10), DS.Tables(11), 69, 3)
        End If

        owb.Worksheets(4).select()
        osheet = owb.Worksheets(4)

        osheet.range("nqsu") = nqsu
        osheet.range("nqsu1") = nqsu1
        osheet.range("nqsu2") = nqsu2
        osheet.range("nqsu3") = nqsu3
        osheet.range("nqsu4") = nqsu4

        osheet.range("scsasl") = scsasl
        osheet.range("scsasl1") = scsasl1
        osheet.range("scsasl2") = scsasl2
        osheet.range("scsasl3") = scsasl3
        osheet.range("scsasl4") = scsasl4

        osheet.range("scssl") = scssl
        osheet.range("scssl1") = scssl1
        osheet.range("scssl2") = scssl2
        osheet.range("scssl3") = scssl3
        osheet.range("scssl4") = scssl4

        osheet.range("scsslnet") = scsslnet
        osheet.range("scsslnet1") = scsslnet1
        osheet.range("scsslnet2") = scsslnet2
        osheet.range("scsslnet3") = scsslnet3
        osheet.range("scsslnet4") = scsslnet4

        osheet.range("sclt") = sclt
        osheet.range("sclt1") = sclt1
        osheet.range("sclt2") = sclt2
        osheet.range("sclt3") = sclt3
        osheet.range("sclt4") = sclt4

        osheet.range("scno") = scno
        osheet.range("scno1") = scno1
        osheet.range("scno2") = scno2
        osheet.range("scno3") = scno3
        osheet.range("scno4") = scno4

        osheet.range("scsh") = scsh
        osheet.range("scsh1") = scsh1
        osheet.range("scsh2") = scsh2
        osheet.range("scsh3") = scsh3
        osheet.range("scsh4") = scsh4

        osheet.range("piavg") = piavg
        osheet.range("piavg1") = piavg1
        osheet.range("piavg2") = piavg2
        osheet.range("piavg3") = piavg3
        osheet.range("piavg4") = piavg4

        osheet.range("pilkp") = pilkp
        osheet.range("pilkp1") = pilkp1
        osheet.range("pilkp2") = pilkp2
        osheet.range("pilkp3") = pilkp3
        osheet.range("pilkp4") = pilkp4

        osheet.range("pistd") = pistd
        osheet.range("pistd1") = pistd1
        osheet.range("pistd2") = pistd2
        osheet.range("pistd3") = pistd3
        osheet.range("pistd4") = pistd4

        osheet.range("pvstd") = pvstd
        osheet.range("pvstd1") = pvstd1
        osheet.range("pvstd2") = pvstd2
        osheet.range("pvstd3") = pvstd3
        osheet.range("pvstd4") = pvstd4

        osheet.range("pdrp14") = pdrp14
        osheet.range("pdrp141") = pdrp141
        osheet.range("pdrp142") = pdrp142
        osheet.range("pdrp143") = pdrp143
        osheet.range("pdrp144") = pdrp144

        osheet.range("pdrp5") = pdrp5
        osheet.range("pdrp51") = pdrp51
        osheet.range("pdrp52") = pdrp52
        osheet.range("pdrp53") = pdrp53
        osheet.range("pdrp54") = pdrp54

        osheet.range("pdramp") = pdramp
        osheet.range("pdramp1") = pdramp1
        osheet.range("pdramp2") = pdramp2
        osheet.range("pdramp3") = pdramp3
        osheet.range("pdramp4") = pdramp4

        osheet.range("pdpot") = pdpot
        osheet.range("pdpot1") = pdpot1
        osheet.range("pdpot2") = pdpot2
        osheet.range("pdpot3") = pdpot3
        osheet.range("pdpot4") = pdpot4

        owb.Worksheets(5).select()
        osheet = owb.Worksheets(5)
        If Not IsNothing(SAPBS) Then

            Dim rowint As Integer = 11
            For Each drv As DataRowView In SAPBS.List
                Dim colint As Integer = 1
                For i = 0 To 4
                    colint = colint + 1
                    osheet.cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                Next
                rowint = rowint + 1
            Next
            osheet.Cells.EntireColumn.AutoFit()

        End If

        If Not IsNothing(ContractBS) Then
            For Each drv As DataRowView In ContractBS.List
                Select Case drv.Item("doctypeid")
                    Case 32
                        osheet.range("contractpaymentterms") = "" & drv.Item("payt") & " " & drv.Item("details")
                        osheet.range("contractpaymenttermseffectivedate") = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("docdate"))
                        osheet.range("contractpaymenttermsexpireddate") = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("expireddate"))
                    Case 33
                        osheet.range("contractnqsu") = "" & drv.Item("nqsu")
                        osheet.range("nqsucomment") = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("remarks"))
                        osheet.range("nqsueffectivedate") = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("docdate"))
                        osheet.range("nqsuexpireddate") = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("expireddate"))

                    Case 35
                        osheet.range("contractsasl") = "" & drv.Item("sasl") & " %"
                        osheet.range("contractleadtime") = "" & drv.Item("leadtime")
                        osheet.range("saslcomment") = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("remarks"))
                        osheet.range("sasleffectivedate") = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("docdate"))
                        osheet.range("saslexpireddate") = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("expireddate"))
                End Select
            Next
        End If

        If Not IsNothing(VendorCurrencyBS) Then
            Dim i As Integer = 0
            Dim rowint As Integer = osheet.range("startrowVendorCurrency").row
            For Each drv As DataRowView In VendorCurrencyBS.List
                i = i + 1

                osheet.cells(rowint, 2).value = "" & drv.Row.Item("vendorcode").ToString
                osheet.cells(rowint, 3).value = "" & drv.Row.Item("crcy").ToString
                osheet.cells(rowint, 4).value = "" & drv.Row.Item("effectivedate").ToString

                If i > 4 Then
                    Exit For
                End If
            Next

        End If

        If Not IsNothing(CitiProgramBS) Then
            'Dim rowint As Integer = 37 '27
            Dim rowint As Integer = osheet.range("startrowCitiProgram").row
            For Each drv As DataRowView In CitiProgramBS.List
                'Dim colint As Integer = 1
                'For i = 1 To 2
                '    colint = colint + 1
                '    osheet.cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                'Next
                osheet.cells(rowint, 2).value = "" & drv.Row.Item("fpcp").ToString
                osheet.cells(rowint, 3).value = "" & drv.Row.Item("crcy").ToString
                osheet.cells(rowint, 4).value = "" & drv.Row.Item("joindate").ToString
                osheet.cells(rowint, 5).value = "" & drv.Row.Item("enddate").ToString
                rowint = rowint + 1
            Next
            osheet.Cells.EntireColumn.AutoFit()
        End If

        'DocumentType
        'Dim myRow As Integer = 43
        Dim myRow As Integer = osheet.range("startrowAuthProject").row
        If Not IsNothing(AuthletterBS) Then
            For Each drv As DataRowView In AuthletterBS.List
                Dim colint As Integer = 2
                osheet.cells(myRow, 2).value = "Authorization letter"
                'For i = 1 To 5
                '    colint = colint + 1
                '    osheet.cells(myRow, colint).value = "" & drv.Row.Item(i).ToString
                'Next
                osheet.cells(myRow, 3).value = "" & drv.Row.Item("docdate").ToString
                osheet.cells(myRow, 4).value = "" & drv.Row.Item("projectname").ToString
                osheet.cells(myRow, 6).value = "" & drv.Row.Item("remarks").ToString
                osheet.cells(myRow, 7).value = "" & drv.Row.Item("version").ToString
                osheet.cells(myRow, 8).value = "" & drv.Row.Item("expireddate").ToString
                myRow = myRow + 1
            Next
            osheet.Cells.EntireColumn.AutoFit()
        End If

        If Not IsNothing(ProjectSpecBS) Then
            For Each drv As DataRowView In ProjectSpecBS.List
                Dim colint As Integer = 2
                osheet.cells(myRow, 2).value = "Project Specification"
                'For i = 1 To 5
                '    colint = colint + 1
                '    osheet.cells(myRow, colint).value = "" & drv.Row.Item(i).ToString
                'Next
                osheet.cells(myRow, 3).value = "" & drv.Row.Item("docdate").ToString
                osheet.cells(myRow, 4).value = "" & drv.Row.Item("projectname").ToString
                osheet.cells(myRow, 5).value = "" & drv.Row.Item("returnrate").ToString
                osheet.cells(myRow, 6).value = "" & drv.Row.Item("remarks").ToString
                osheet.cells(myRow, 7).value = "" & drv.Row.Item("version").ToString
                osheet.cells(myRow, 8).value = "" & drv.Row.Item("expireddate").ToString
                myRow = myRow + 1
            Next
            osheet.Cells.EntireColumn.AutoFit()
            osheet.cells.EntireRow.autoFit()
        End If
        osheet.cells.EntireRow.autoFit()
        osheet.Rows("1:5").EntireRow.Hidden = True

        If Not IsNothing(ToolingHDBS) Then
            owb.Worksheets(6).select()
            osheet = owb.Worksheets(6)
            Dim rowint As Integer = 11
            For Each drv As DataRowView In ToolingHDBS.List
                Dim colint As Integer = 1
                For i = 0 To 4
                    colint = colint + 1
                    osheet.cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                Next
                rowint = rowint + 1
            Next
            osheet.Cells.EntireColumn.AutoFit()
        End If

        If Not IsNothing(ToolingDetailsALLBS) Then
            owb.Worksheets(7).select()
            osheet = owb.Worksheets(7)
            Dim rowint As Integer = 12
            For Each drv As DataRowView In ToolingDetailsALLBS.List
                Dim colint As Integer = 1
                For i = 0 To 16
                    colint = colint + 1
                    osheet.cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                Next
                rowint = rowint + 1
            Next
            osheet.Cells.EntireColumn.AutoFit()
        End If

        If Not IsNothing(PCharterBS) Then
            owb.Worksheets(8).select()
            osheet = owb.Worksheets(8)
            Dim rowint As Integer = 11
            For Each drv As DataRowView In PCharterBS.List
                Dim colint As Integer = 2
                osheet.cells(rowint, colint).value = "Responsible Purchasing Charter"
                For i = 0 To 2
                    colint = colint + 1
                    osheet.cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                Next
                rowint = rowint + 1
            Next
            osheet.Cells.EntireColumn.AutoFit()
        End If

        If Not IsNothing(SocialAuditBS) Then
            owb.Worksheets(8).select()
            osheet = owb.Worksheets(8)
            Dim rowint As Integer = 11
            For Each drv As DataRowView In SocialAuditBS.List
                Dim colint As Integer = 8
                osheet.cells(rowint, colint).value = "Social Audit Result"
                For i = 0 To 7
                    colint = colint + 1
                    osheet.cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                Next
                rowint = rowint + 1
            Next
            osheet.Cells.EntireColumn.AutoFit()
        End If

        If Not IsNothing(PanelStatusBS) Then
            owb.Worksheets(9).select()
            osheet = owb.Worksheets(9)
            'Dim rowint As Integer = 19 '11
            Dim rowint As Integer = osheet.range("startrowSupplierPanel").row '11
            For Each drv As DataRowView In PanelStatusBS.List
                Dim colint As Integer = 1
                For i = 0 To 3
                    colint = colint + 1
                    osheet.cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                Next
                rowint = rowint + 1
            Next
            osheet.Cells.EntireColumn.AutoFit()
        End If

        If Not IsNothing(VendorBS) Then
            owb.Worksheets(9).select()
            osheet = owb.Worksheets(9)

            'Dim drv As DataRowView = VendorBS.Item(0)
            'osheet.range("gsm").value = drv.Row.Item("gsm")
            'osheet.range("gsmsince").value = drv.Row.Item("effectivedate")
            'osheet.range("spm").value = drv.Row.Item("ssm")
            'osheet.range("spmsince").value = drv.Row.Item("ssmeffectivedate")
            'osheet.range("pm").value = drv.Row.Item("pm")
            'osheet.range("pmsince").value = drv.Row.Item("pmeffectivedate")
            Dim rowint As Integer = osheet.range("startrowgsm").row '11
            Dim i As Integer
            For Each drv As DataRowView In VendorBS.List
                i = i + 1
                Dim colint As Integer = 1
                'For i = 0 To 3
                '    colint = colint + 1
                '    osheet.cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                'Next
                osheet.cells(rowint, 2).value = "" & drv.Row.Item("vendorcode").ToString
                osheet.cells(rowint, 3).value = "" & drv.Row.Item("gsm").ToString
                osheet.cells(rowint, 4).value = "" & drv.Row.Item("effectivedate").ToString
                osheet.cells(rowint, 5).value = "" & drv.Row.Item("ssm").ToString
                osheet.cells(rowint, 6).value = "" & drv.Row.Item("ssmeffectivedate").ToString
                osheet.cells(rowint, 7).value = "" & drv.Row.Item("pm").ToString
                osheet.cells(rowint, 8).value = "" & drv.Row.Item("pmeffectivedate").ToString

                rowint = rowint + 1
                If i > 4 Then
                    Exit For
                End If
            Next


        End If

        If Not IsNothing(SupplierTechnologyBS) Then
            owb.Worksheets(9).select()
            osheet = owb.Worksheets(9)
            Dim vt As String = String.Empty
            Dim technologyname = String.Empty
            For Each drv As DataRowView In SupplierTechnologyBS.List
                technologyname = drv.Row.Item("technologyname")
                If drv.Row.Item("lineno") = 1 Then
                    osheet.range("cpmain").value = technologyname
                Else
                    If vt.Length > 0 Then
                        vt = vt + ","
                    End If
                    vt = vt + technologyname
                End If
            Next
            osheet.range("cptechnology").value = vt
        End If

        If Not IsNothing(SupplierPanelHistory) Then
            owb.Worksheets(9).select()
            osheet = owb.Worksheets(9)
            'Dim rowint As Integer = 28 '26
            Dim rowint As Integer = osheet.range("startrowVendorHistory").row
            For Each drv As DataRowView In SupplierPanelHistory.List
                Dim colint As Integer = 1
                For i = 0 To 5
                    colint = colint + 1
                    osheet.cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                Next
                rowint = rowint + 1
            Next
            osheet.Cells.EntireColumn.AutoFit()
        End If

        If Not IsNothing(FactoryBS) Then
            owb.Worksheets(10).select()
            osheet = owb.Worksheets(10)
            If FactoryBS.Count > 0 Then
                Dim mydrv As DataRowView = FactoryBS.Item(0)
                osheet.range("customname") = mydrv.Row.Item("customname")
            End If

            Dim rowint As Integer = 14
            For Each drv As DataRowView In FactoryBS.List
                Dim colint As Integer = 1
                For i = 0 To 7
                    colint = colint + 1
                    osheet.cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                Next
                rowint = rowint + 1
            Next
            osheet.Cells.EntireColumn.AutoFit()
        End If

        If Not IsNothing(ContactBS) Then
            owb.Worksheets(10).select()
            osheet = owb.Worksheets(10)

            Dim rowint As Integer = 36
            For Each drv As DataRowView In ContactBS.List
                Dim colint As Integer = 1
                For i = 2 To 8
                    colint = colint + 1
                    osheet.cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                Next
                rowint = rowint + 1
            Next
            osheet.Cells.EntireColumn.AutoFit()
        End If

        Dim myrowid As Integer = 12
        If Not IsNothing(SIFIDBS) Then

            owb.Worksheets(11).select()
            osheet = owb.Worksheets(11)

            For Each drv As DataRowView In SIFIDBS.List
                Dim colint As Integer = 1
                Dim inc As Integer = 1
                If drv.Item("doctypename") <> "SIF" Then
                    colint = 0
                    inc = 2
                Else
                    If SIFIDBS.Count > 0 Then
                        Dim mydrv As DataRowView = SIFIDBS.Item(0)
                        osheet.range("sifdate") = mydrv.Row.Item("docdate")
                    End If
                End If

                For i = 0 To 1
                    colint = colint + inc
                    osheet.cells(myrowid, colint).value = "" & drv.Row.Item(i).ToString
                Next
                myrowid = myrowid + 1
            Next
            'osheet.Cells.EntireColumn.AutoFit()
        End If


        If Not IsNothing(SIFBS) Then
            owb.Worksheets(11).select()
            osheet = owb.Worksheets(11)
            If SIFBS.Count > 0 Then
                Dim mydrv As DataRowView = SIFBS.Item(0)
                osheet.range("sifdate") = mydrv.Row.Item("docdate")
            End If

            For Each drv As DataRowView In SIFBS.List
                Dim colint As Integer = 1
                For i = 0 To 1
                    colint = colint + 1
                    osheet.cells(myrowid, colint).value = "" & drv.Row.Item(i).ToString
                Next
                myrowid = myrowid + 1
            Next
            'osheet.Cells.EntireColumn.AutoFit()
        End If

        If Not IsNothing(IDBS) Then
            owb.Worksheets(11).select()
            osheet = owb.Worksheets(11)
            If IDBS.Count > 0 Then
                Dim mydrv As DataRowView = IDBS.Item(0)
                osheet.range("idsdate") = mydrv.Row.Item("docdate")
            End If

            For Each drv As DataRowView In IDBS.List
                Dim colint As Integer = 0
                For i = 0 To 1
                    colint = colint + 2
                    osheet.cells(myrowid, colint).value = "" & drv.Row.Item(i).ToString
                Next
                myrowid = myrowid + 1
            Next
            'osheet.Cells.EntireColumn.AutoFit()
        End If

        If Not IsNothing(PriceCMMFBS) Then
            owb.Worksheets(12).select()
            osheet = owb.Worksheets(12)
            If PriceCMMFBS.Count > 0 Then
                osheet.range("cmmffilter").value = "'" & CMMFFilter
                osheet.range("lkpfilter").value = "" & LKPFiter
            End If
            Dim rowint As Long = 14
            For Each drv As DataRowView In PriceCMMFBS.List
                Dim colint As Integer = 1
                For i = 0 To 20
                    colint = colint + 1
                    osheet.cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                Next
                rowint = rowint + 1
            Next
            'osheet.Cells.EntireColumn.AutoFit()
        End If

        '

        owb.Worksheets(1).select()

    End Sub
    Private Sub showTotal(ByVal osheet As Excel.Worksheet, ByVal dataTable As DataTable, ByVal dataTable2 As DataTable, ByVal row As Integer, ByVal startcolumn As Integer)
        Dim c As Integer
        Dim col As Integer
        For Each dc As DataColumn In dataTable.Columns
            If dc.ColumnName <> "." And dc.ColumnName <> "vendorcode" Then
                osheet.Cells(row, startcolumn + col) = dc.ColumnName
                If dataTable.Rows.Count > 0 Then
                    osheet.Cells(row + 1, startcolumn + col) = dataTable.Rows(0).Item(c) 'IIf(IsDBNull(dataTable.Rows(0).Item(c)), 0, dataTable.Rows(0).Item(c))
                End If

                col = col + 1
            End If
            c = c + 1
        Next
        c = 0
        For Each dc As DataColumn In dataTable2.Columns
            If dc.ColumnName <> " " And dc.ColumnName <> "vendorcode" Then
                osheet.Cells(row, startcolumn + col) = dc.ColumnName
                If dataTable2.Rows.Count > 0 Then
                    osheet.Cells(row + 1, startcolumn + col) = dataTable2.Rows(0).Item(c) 'IIf(IsDBNull(dataTable2.Rows(0).Item(c)), 0, dataTable2.Rows(0).Item(c))
                End If

                col = col + 1
            End If
            c = c + 1
        Next
    End Sub
    Private Sub showDetail(ByVal osheet As Excel.Worksheet, ByVal myfilter As String, ByVal ColName As String, ByVal maxrow As Integer, ByVal dataTable As DataTable, ByVal dataTable2 As DataTable, ByVal row As Integer, ByVal startcolumn As Integer)
        Dim c As Integer
        Dim col As Integer
        For Each dc As DataColumn In dataTable.Columns
            If dc.ColumnName <> "." Then
                osheet.Cells(row, startcolumn + col) = IIf(col = 0, ColName, dc.ColumnName)
                Dim irow As Integer
                For Each dr As DataRow In dataTable.Rows
                    If irow > maxrow Then
                        Exit For
                    End If
                    osheet.Cells(row + 1 + irow, startcolumn + col) = dr.Item(c) 'IIf(IsDBNull(dr.Item(c)), 0, dr.Item(c))
                    irow += 1
                Next
                col = col + 1
                irow = 0
            End If
            c = c + 1

        Next
        c = 0
        For Each dc As DataColumn In dataTable2.Columns
            If dc.ColumnName <> " " And dc.ColumnName <> myfilter Then
                osheet.Cells(row, startcolumn + col) = dc.ColumnName
                Dim irow As Integer
                For Each dr As DataRow In dataTable2.Rows
                    osheet.Cells(row + 1 + irow, startcolumn + col) = dr.Item(c) 'IIf(IsDBNull(dr.Item(c)), 0, dr.Item(c))
                    irow += 1
                Next
                irow = 0
                col = col + 1
            End If
            c = c + 1
        Next
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



End Class