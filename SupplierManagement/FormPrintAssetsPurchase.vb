Imports System.Threading
Imports Microsoft.Office.Interop
Imports SupplierManagement.SharedClass
Imports System.Text
Imports SupplierManagement.PublicClass
Public Class FormPrintAssetsPurchase

    Dim myThread As New Threading.Thread(AddressOf DoWork)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Dim DS As DataSet
    Dim APBS As New BindingSource
    Dim TPBS As New BindingSource
    Dim VCBS As New BindingSource
    Dim TIBS As New BindingSource
    Dim APACTBS As New BindingSource

    Dim InvArr(,) As String = {{"invoicedate1", "invoiceno1", "invoicedescription1", "totalinvoiceamount1"},
                               {"invoicedate2", "invoiceno2", "invoicedescription2", "totalinvoiceamount2"},
                               {"invoicedate3", "invoiceno3", "invoicedescription3", "totalinvoiceamount3"},
                               {"invoicedate4", "invoiceno4", "invoicedescription4", "totalinvoiceamount4"},
                               {"invoicedate5", "invoiceno5", "invoicedescription5", "totalinvoiceamount5"},
                               {"invoicedate6", "invoiceno6", "invoicedescription6", "totalinvoiceamount6"},
                               {"invoicedate7", "invoiceno7", "invoicedescription7", "totalinvoiceamount7"},
                               {"invoicedate8", "invoiceno8", "invoicedescription8", "totalinvoiceamount8"},
                               {"invoicedate9", "invoiceno9", "invoicedescription9", "totalinvoiceamount9"},
                               {"invoicedate10", "invoiceno10", "invoicedescription10", "totalinvoiceamount10"},
                               {"invoicedate11", "invoiceno11", "invoicedescription11", "totalinvoiceamount11"},
                               {"invoicedate12", "invoiceno12", "invoicedescription12", "totalinvoiceamount12"},
                               {"invoicedate13", "invoiceno13", "invoicedescription13", "totalinvoiceamount13"},
                               {"invoicedate14", "invoiceno14", "invoicedescription14", "totalinvoiceamount14"},
                               {"invoicedate15", "invoiceno15", "invoicedescription15", "totalinvoiceamount15"},
                               {"invoicedate16", "invoiceno16", "invoicedescription16", "totalinvoiceamount16"},
                               {"invoicedate17", "invoiceno17", "invoicedescription17", "totalinvoiceamount17"},
                               {"invoicedate18", "invoiceno18", "invoicedescription18", "totalinvoiceamount18"},
                               {"invoicedate19", "invoiceno19", "invoicedescription19", "totalinvoiceamount19"},
                               {"invoicedate20", "invoiceno20", "invoicedescription20", "totalinvoiceamount20"}
                               }

    Public Sub New(ByVal ds As DataSet)

        ' This call is required by the designer.
        InitializeComponent()
        Me.DS = ds
        ' Add any initialization after the InitializeComponent() call.

        APBS.DataSource = ds.Tables("AssetPurchase")
        TPBS.DataSource = ds.Tables("ToolingProject")
        VCBS.DataSource = ds.Tables("Vendor")
        TIBS.DataSource = ds.Tables("ToolingInvoice")
        APACTBS.DataSource = ds.Tables("AssetPurchaseAction")


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

       
        runreport(sender, e)
       
    End Sub

    Private Sub runreport(ByVal sender As System.Object, ByVal e As System.EventArgs)


        Dim mymessage As String = String.Empty

        Dim sqlstr As String = "select foo.*,vs.sbuname from (select * from doc.countdocument union all select * from doc.vendormissingdocument) foo" &
                                " left join doc.vendorsbu vs on vs.vendorcode = foo.vendorcode;"

        Dim mysaveform As New SaveFileDialog
        mysaveform.DefaultExt = "xlsx"
        mysaveform.FileName = String.Format("FormAssetsPurchase{0:yyyyMMdd}", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 8 'because hidden

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\\172.22.10.44\Users_I\Logistic Dept\KPI & Reporting\templates\KPI Graph Final.xltx")
            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\SupplierDocument.xltm")

            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\FormAssetsPurchase.xltx")
            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\\172.22.10.44\SharedFolder\PriceCMMF\New\template\FormAssetsPurchase.xltx")

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, String.Format("{0}\FormAssetsPurchase.xltx", HelperClass1.template))
            myreport.CreateForm(Me, e)
        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        Dim oXl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        oXl = owb.Parent

        owb.Worksheets(1).select()
        Dim osheet As Excel.Worksheet = owb.Worksheets(1)
        'Dim orange = osheet.Range("applicationdate")
        Dim APdrv As DataRowView = APBS.Current

        Dim pos = TPBS.Find("id", APdrv.Item("projectid"))
        TPBS.Position = pos
        Dim TPdrv As DataRowView = TPBS.Current

        Dim pos1 = VCBS.Find("vendorcode", APdrv.Item("vendorcode"))
        VCBS.Position = pos1
        Dim VCdrv As DataRowView = VCBS.Current

        osheet.Range("applicationdate").Value = String.Format("{0:dd-MMM-yyyy}", APdrv.Row.Item("applicantdate"))
        osheet.Range("applicantname").Value = String.Format("{0}", APdrv.Row.Item("applicantname"))
        osheet.Range("approvalname").Value = String.Format("{0}", APdrv.Row.Item("approvalname"))
        osheet.Range("creator").Value = APdrv.Row.Item("username")
        osheet.Range("dept").Value = TPdrv.Row.Item("dept")
        osheet.Range("assetdescription").Value = APdrv.Row.Item("assetdescription")
        osheet.Range("productfamilysbu").Value = TPdrv.Row.Item("sbuname2")
        osheet.Range("familydescription").Value = String.Format("{0} - {1} ({2})", TPdrv.Row.Item("familyid"), TPdrv.Row.Item("familyname"), TPdrv.Row.Item("groupingcode"))
        osheet.Range("projectcodename").Value = TPdrv.Row.Item("projectcode")
        osheet.Range("projectcodename2").Value = TPdrv.Row.Item("projectname")
        osheet.Range("paymentmethod").Value = APdrv.Row.Item("paymentmethod")

        Select Case APdrv.Row.Item("typeofinvestment")
            Case 1
                osheet.Range("newasset").Value = "X"
            Case 2
                osheet.Range("toolingmodification").Value = "X"
            Case 3
                osheet.Range("backup").Value = "X"
            Case 4
                osheet.Range("increasecapacity").Value = "X"
        End Select
        osheet.Range("reasonforpurchase").Value = APdrv.Row.Item("reason")
        osheet.Range("suppliername").Value = VCdrv.Row.Item("vendorname")
        osheet.Range("suppliercode").Value = APdrv.Row.Item("vendorcode")

        osheet.Range("amortizationperiod1").Value = APdrv.Row.Item("amortperiod_1")
        osheet.Range("amorizationqty1").Value = APdrv.Row.Item("amortqty_1")
        osheet.Range("totalamortizationqty1").Value = APdrv.Row.Item("totalamortqty_1")
        osheet.Range("totalamorizationamount1").Value = APdrv.Row.Item("totalamortamount_1")
        osheet.Range("amortizationperunit1").Value = APdrv.Row.Item("amortperunit_1")

        osheet.Range("amortizationperiod2").Value = APdrv.Row.Item("amortperiod_2")
        osheet.Range("amorizationqty2").Value = APdrv.Row.Item("amortqty_2")
        osheet.Range("totalamortizationqty2").Value = APdrv.Row.Item("totalamortqty_2")
        osheet.Range("totalamorizationamount2").Value = APdrv.Row.Item("totalamortamount_2")
        osheet.Range("amortizationperunit2").Value = APdrv.Row.Item("amortperunit_2")
        Dim i As Integer = 0

        For Each drv As DataRowView In TIBS.List
            Try
                'osheet.Cells(35 + i, 1).value = drv.Row.Item("invoicedate")
                'osheet.Cells(35 + i, 2).value = drv.Row.Item("invoiceno")
                'osheet.Cells(35 + i, 3).value = drv.Row.Item("description")
                'osheet.Cells(35 + i, 4).value = drv.Row.Item("totalamount")
                osheet.Range(InvArr(i, 0)).Value = drv.Row.Item("invoicedate")
                osheet.Range(InvArr(i, 1)).Value = drv.Row.Item("invoiceno")
                osheet.Range(InvArr(i, 2)).Value = drv.Row.Item("description")
                osheet.Range(InvArr(i, 3)).Value = drv.Row.Item("totalamount")
                i = i + 1
            Catch ex As Exception
                MessageBox.Show("Your invoice number more than available row in the template.")
                Exit Sub
            End Try

        Next
        For a = i To 20
            osheet.Rows(String.Format("{0}:{0}", 35 + i, 35 + i)).delete()
        Next

        If Not IsDBNull(APdrv.Row.Item("budgetamount")) Then
            osheet.Range("budgeted1").Value = "X"
            'osheet.Range("budgetamount").Value = APdrv.Row.Item("budgetamount") * APdrv.Row.Item("exchangerate")
            osheet.Range("budgetamount").Value = APdrv.Row.Item("budgetamount")
            osheet.Range("currency").Value = APdrv.Row.Item("budgetcurr")
            osheet.Range("aebno").Value = APdrv.Row.Item("aeb")
            If Not IsDBNull(APdrv.Row.Item("toolingcost")) Then
                If APdrv.Row.Item("toolingcost") <= (APdrv.Row.Item("budgetamount") * APdrv.Row.Item("exchangerate")) Then
                    osheet.Range("withinbudget1").Value = "X"
                Else
                    osheet.Range("withinbudget2").Value = "X"
                End If
            End If


            osheet.Range("aebno").Value = APdrv.Row.Item("aeb")
        Else
            osheet.Range("budgeted2").Value = "X"
        End If
        osheet.Range("comment").Value = APdrv.Row.Item("comments")

        'Get Date Approval First Priority Status 4, if not found Status5
        APACTBS.Sort = "status asc"
        For Each drv As DataRowView In APACTBS.List
            If drv.Row.Item("status") = 4 Then
                osheet.Range("approvaldate").Value = String.Format("{0:dd-MMM-yyyy}", drv.Row.Item("latestupdate"))
                If IsDBNull(APdrv.Row.Item("approvalname2")) Then
                    osheet.Range("signature").Value = String.Format("Approved by {0}", APdrv.Row.Item("approvalname"))
                Else
                    osheet.Range("signature").Value = String.Format("Approved by {0} and ", APdrv.Row.Item("approvalname"), APdrv.Row.Item("approvalname2"))
                End If
                Exit For
            End If
            If drv.Row.Item("status") = 5 Then
                osheet.Range("approvaldate").Value = String.Format("{0:dd-MMM-yyyy}", drv.Row.Item("latestupdate"))
                If IsDBNull(APdrv.Row.Item("approvalname2")) Then
                    osheet.Range("signature").Value = String.Format("Approved by {0}", APdrv.Row.Item("approvalname"))
                Else
                    osheet.Range("signature").Value = String.Format("Approved by {0} and ", APdrv.Row.Item("approvalname"), APdrv.Row.Item("approvalname2"))
                End If
                Exit For
            End If
        Next

        'If IsDBNull(APdrv.Row.Item("approvalname2")) Then
        '    osheet.Range("signature").Value = String.Format("Approved by {0}", APdrv.Row.Item("approvalname"))
        'Else
        '    osheet.Range("signature").Value = String.Format("Approved by {0} and ", APdrv.Row.Item("approvalname"), APdrv.Row.Item("approvalname2"))
        'End If

    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        ''Throw New NotImplementedException
        'Dim oXl As Excel.Application = Nothing
        'Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        'oXl = owb.Parent


        'owb.Worksheets(8).select()

        'Dim osheet = owb.Worksheets(8)
        'osheet.name = "RawData"

        'owb.Names.Add("rawdata", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C2),COUNTA('RawData'!R1))")

        'owb.Worksheets(1).select()
        'osheet = owb.Worksheets(1)
        ''oXl.Run("ShowFG")
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        ''oXl.Run("ShowCP")
        'osheet.PivotTables("PivotTable2").PivotCache.Refresh()
        ''oXl.Run("OpenTabs")

        'osheet = owb.Worksheets(2)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(3)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(4)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(5)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(6)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(1)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(7)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()


        ''oXl.Run("BACK")
    End Sub

    Private Sub DoWork()
        ProgressReport(1, "Updating SBU... Please wait")
        ProgressReport(6, "")
        Dim myvendorsbu = New VendorSBU
        myvendorsbu.UpdateVendorSBU()
        myvendorsbu.UpdateShortnameInfo()
        ProgressReport(1, "")
        ProgressReport(4, "")
        ProgressReport(5, "")
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


    Private Sub FormReportDocumentCount_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub


End Class