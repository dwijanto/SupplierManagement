Imports System.Windows.Forms
Imports System.Threading
Imports System.Text
Imports Microsoft.Office.Interop
Public Class DialogExportBuyerInput2

    Enum ExportTypeEnum
        NEW_BUYER
        SELECTED_BUYER
    End Enum
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
    Public Property frontproductype As String

    Public CurrentMonth As Date
    Public frontshortname As String
    Public frontsuppliername As String
    Public frontsuppliercode As String

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
    Private DS As DataSet
    'Public Property shortname

    Dim ExportType As ExportTypeEnum

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myadapter As ExportBuyerInputAdapter
    Dim sb As New StringBuilder

    Dim SupplierPhotoBS As BindingSource
    Dim PhotoFolder As String
    Dim Sqlstr As String

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    Private Sub loaddata()
        'MessageBox.Show(myuser)
        ComboBox1.SelectedIndex = 0
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
    Public Sub New(ByVal _DS As DataSet, ByVal shortname As String)

        ' This call is required by the designer.
        InitializeComponent()
        Me.DS = _DS
        myadapter = New ExportBuyerInputAdapter(shortname)

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByVal sqlstr As String, ByVal shortname As String)

        ' This call is required by the designer.
        InitializeComponent()
        Me.Sqlstr = sqlstr
        myadapter = New ExportBuyerInputAdapter(shortname)

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub DialogExportBuyerInput_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DateTimePicker1.Value = Today.AddMonths(-11)
        loaddata()
    End Sub

    Private Sub DoWork()
        ProgressReport(6, "Marquee")
        If myadapter.LoadData() Then
            ProgressReport(4, "Init Data")
        End If
        ProgressReport(5, "Continuous")
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
                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = myadapter.BS
                            Select Case frontproductype
                                Case "FP"
                                    RadioButton1.Checked = True
                                Case "CP"
                                    RadioButton2.Checked = True
                                Case Else
                                    RadioButton3.Checked = True
                            End Select
                        Catch ex As Exception
                            message = ex.Message
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GET_SELECTED.Click, NEW_BUTTON.Click
        Dim myerror As New ErrorProvider
        myerror.SetError(RadioButton1, "")
        If Not (RadioButton1.Checked Or RadioButton2.Checked Or RadioButton3.Checked) Then
            myerror.SetError(RadioButton1, "Please select from Photo Source.")
            Exit Sub
        End If
        Select Case frontproductype
            Case "ALL", "FP", "CP"
            Case Else
                frontproductype = "ALL(General)"
        End Select
        Select Case sender.name
            Case "NEW_BUTTON"
                ExportType = ExportTypeEnum.NEW_BUYER
                'Select Case frontproductype
                '    Case "ALL", "FP", "CP"
                '    Case Else
                '        frontproductype = "ALL"
                'End Select

            Case "GET_SELECTED"
                ExportType = ExportTypeEnum.SELECTED_BUYER
                Dim drv = myadapter.BS.Current
                'frontproductype = drv.row.item("producttype")
        End Select
        'get_photos

        Dim radiobuttonselected As String
        If RadioButton1.Checked Then
            radiobuttonselected = "fp"
        ElseIf RadioButton2.Checked Then
            radiobuttonselected = "cp"
        Else
            radiobuttonselected = "general"
        End If
        Dim myphotos As New SupplierPhotosAdapter
        If myphotos.GetSupplierPhoto(frontshortname, radiobuttonselected) Then
            SupplierPhotoBS = New BindingSource
            SupplierPhotoBS.DataSource = myphotos.DS.Tables(0)
            PhotoFolder = myphotos.DS.Tables(1).Rows(0).Item(0)
        End If

        'Refresh ActionPlanBS based on criteria
        Dim mycriteria As String = String.Empty
        If CheckBox1.Checked Then
            Dim myarray As String() = {"finishdate", "startdate", "enddate"}
            mycriteria = String.Format("{0} >= '{1:yyyy-MM-dd}' and {0} <= '{2:yyyy-MM-dd}' and status = 'Closed'", myarray(ComboBox1.SelectedIndex), DateTimePicker1.Value.Date, DateTimePicker2.Value.Date)
        End If
        myadapter.GetActionPlan(mycriteria)
        export()
    End Sub

    '    Private Sub export(ByVal drv As Object, ByVal e As System.EventArgs)
    Private Sub export()
        Dim mysaveform As New SaveFileDialog
        'mysaveform.FileName = String.Format("SupplierDocumentFilterReport{0:yyyyMMdd}.xlsm", Date.Today)
        'mysaveform.FileName = String.Format("Buyer_Input_{0:yyyyMMdd}.xlsx", Date.Today)
        mysaveform.FileName = String.Format("SSS_{0}_{2}_{1:yyyyMMdd}.xlsx", frontshortname.Replace("""", ""), Date.Today, frontproductype)
        mysaveform.DefaultExt = ".xlsx"
        'Dim sqlstr As String
        Dim myTemplate As String = "\templates\BuyerInput.xltx"
        'sqlstr = ""
        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 9

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, myTemplate)
            myreport.Run(Me, New EventArgs)
        End If

    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        
    End Sub
    Private Sub PutPhotos(ByVal oxl As Excel.Application, ByVal osheet As Excel.Worksheet, ByVal phototype As Integer, ByVal phototypename As String)
        Dim query = From Photos In SupplierPhotoBS.List
            Where Photos.item("phototype") = phototype
        Dim j As Integer = 1
        Dim myrow As Integer
        If phototypename = "Factory" Then
            myrow = 40 '35
        Else
            myrow = 49
        End If
        For Each DRV As DataRowView In query
            Dim filename As String = String.Format("{0}\{1}{2}", PhotoFolder, DRV.Row.Item("id"), IO.Path.GetExtension(DRV.Item("filename")))
            If IO.File.Exists(filename) Then
                Dim Width_OP As Long
                Dim Height_OP As Long
                Dim Left_OP As Long
                Dim Top_OP As Long
                Dim New_Pic As Object
                Dim Original_Pic As Object
                Dim Name_OP As String

                'Dim Desc_Pic As Object
                'Desc_Pic = osheet.Shapes.Item(String.Format("{0}desc{1}", phototypename, j))

                'Desc_Pic.textframe.characters.text = DRV.Item("description")


                Original_Pic = osheet.Pictures(String.Format("{0}{1}", phototypename, j))
                Width_OP = Original_Pic.Width
                Height_OP = Original_Pic.Height
                Left_OP = Original_Pic.Left
                Top_OP = Original_Pic.Top
                Name_OP = Original_Pic.Name
                Original_Pic.Delete()
                New_Pic = osheet.Pictures.Insert(filename)
                New_Pic.ShapeRange.LockAspectRatio = -1 'msoTrue
                With New_Pic
                    .Top = Top_OP
                    .Name = Name_OP
                End With
                If New_Pic.Height >= New_Pic.Width Then
                    New_Pic.Height = Height_OP
                    New_Pic.Left = Left_OP + Math.Abs((Width_OP - New_Pic.Width)) / 2
                Else
                    New_Pic.Width = Width_OP
                    If New_Pic.height > Height_OP Then
                        New_Pic.Height = Height_OP
                        New_Pic.Left = Left_OP + Math.Abs((Width_OP - New_Pic.Width)) / 2
                    Else
                        New_Pic.Left = Left_OP
                    End If
                    New_Pic.Top = Top_OP + (Height_OP - New_Pic.Height) / 2
                End If
                j = j + 1
                If j > 8 Then
                    Exit For
                End If
            End If

        Next
    End Sub
    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        Dim oXl As Excel.Application = Nothing
        'Dim osheet As Excel.Worksheet = CType(sender, Excel.Worksheet)
        'Dim owB As Excel.Workbook
        Dim owb = CType(sender, Excel.Workbook)
        Dim osheet As Excel.Worksheet
        'owB = osheet.Parent
        oXl = owB.Parent
        owB.Names.Add("torawdata", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C1),COUNTA('RawData'!R1))")
        Select Case ExportType
            Case ExportTypeEnum.NEW_BUYER

            Case ExportTypeEnum.SELECTED_BUYER
                Dim drv As DataRowView = myadapter.BS.Current
                Dim dr As DataRow = drv.Row
                owB.Worksheets(3).select()
                osheet = owB.Worksheets(3)
                'Fill 

                osheet.Range("Lastvisitdate").Value = dr.Item("lastvisit")
                osheet.Range("Keystakestopic").Value = "" & dr.Item("keystakestopic")
                osheet.Range("Strengths1").Value = "" & dr.Item("strength1")
                osheet.Range("Strengths2").Value = "" & dr.Item("strength2")
                osheet.Range("Strengths3").Value = "" & dr.Item("strength3")
                osheet.Range("Weaknesses1").Value = "" & dr.Item("weaknessess1")
                osheet.Range("Weaknesses2").Value = "" & dr.Item("weaknessess2")
                osheet.Range("Weaknesses3").Value = "" & dr.Item("weaknessess3")
                osheet.Range("Opportunities1").Value = "" & dr.Item("opportunities1")
                osheet.Range("Opportunities2").Value = "" & dr.Item("opportunities2")
                osheet.Range("Opportunities3").Value = "" & dr.Item("opportunities3")
                osheet.Range("Threats1").Value = "" & dr.Item("threats1")
                osheet.Range("Threats2").Value = "" & dr.Item("threats2")
                osheet.Range("Threats3").Value = "" & dr.Item("threats3")
                osheet.Range("YearFYFcstBudgetSEBAsia").Value = "" & dr.Item("currbudget")
                osheet.Range("YearFYFcstBudgetTotal").Value = "" & dr.Item("totalbudget")
                osheet.Range("ProductDevelopment1").Value = "" & dr.Item("productdevelopment1")
                osheet.Range("ProductDevelopment2").Value = "" & dr.Item("productdevelopment2")
                osheet.Range("Negotiationresultcomment1").Value = "" & dr.Item("negotiationresult1")
                osheet.Range("Negotiationresultcomment2").Value = "" & dr.Item("negotiationresult2")
                osheet.Range("Negotiationresultcomment3").Value = "" & dr.Item("negotiationresult3")
                osheet.Range("GeneralCommentOutstandingissue1").Value = "" & dr.Item("outstandingissue1")
                osheet.Range("GeneralCommentOutstandingissue2").Value = "" & dr.Item("outstandingissue2")
                osheet.Range("GeneralCommentOutstandingissue3").Value = "" & dr.Item("outstandingissue3")
                osheet.Range("GeneralCommentOutstandingissue4").Value = "" & dr.Item("outstandingissue4")
                osheet.Range("GeneralCommentOutstandingissue5").Value = dr.Item("outstandingissue5")

                'Fill Action plan
                '
                owb.Worksheets(2).select()
                osheet = owb.Worksheets(2)
                'myadapter.ActionPlanBS.Filter = String.Format("headerid = {0}", dr.Item("headerid"))
                Dim k As Integer = 0
                For Each mydrv As DataRowView In myadapter.ActionPlanBS.List
                    Dim mydr As DataRow = mydrv.Row

                    For i = 3 To mydr.ItemArray.Length - 1
                        osheet.Cells(6 + k, 2 + (i - 3)).value = mydr.Item(i)
                    Next
                    k = k + 1
                Next
        End Select

        owB.Worksheets(6).select()
        osheet = owB.Worksheets(6)
        osheet.Range("frontproducttype").Value = frontproductype
        osheet.Range("frontshortname").Value = frontshortname
        osheet.Range("issuedate").Value = String.Format("{0:dd-MMM-yyyy}", Date.Today)
        osheet.Range("frontyearperiod").Value = String.Format("{0:dd-MMM-yyyy}", CurrentMonth)
        osheet.Range("sify").Value = sify
        osheet.Range("sify1").Value = sify1
        osheet.Range("sify2").Value = sify2
        osheet.Range("sify3").Value = sify3
        osheet.Range("sify4").Value = sify4
        osheet.Range("TOSEBQTY").Value = TOSEBQTY
        osheet.Range("TOSEBQTY1").Value = TOSEBQTY1
        osheet.Range("TOSEBQTY2").Value = TOSEBQTY2
        osheet.Range("TOSEBQTY3").Value = TOSEBQTY3
        osheet.Range("TOSEBQTY4").Value = TOSEBQTY4
        osheet.Range("TOSEBAMT").Value = TOSEBAMT
        osheet.Range("TOSEBAMT1").Value = TOSEBAMT1
        osheet.Range("TOSEBAMT2").Value = TOSEBAMT2
        osheet.Range("TOSEBAMT3").Value = TOSEBAMT3
        osheet.Range("TOSEBAMT4").Value = TOSEBAMT4


        osheet.Range("piavg").Value = piavg
        osheet.Range("piavg1").Value = piavg1
        osheet.Range("piavg2").Value = piavg2
        osheet.Range("piavg3").Value = piavg3
        osheet.Range("piavg4").Value = piavg4
        osheet.Range("pilkp").Value = pilkp
        osheet.Range("pilkp1").Value = pilkp1
        osheet.Range("pilkp2").Value = pilkp2
        osheet.Range("pilkp3").Value = pilkp3
        osheet.Range("pilkp4").Value = pilkp4
        osheet.Range("pistd").Value = pistd
        osheet.Range("pistd1").Value = pistd1
        osheet.Range("pistd2").Value = pistd2
        osheet.Range("pistd3").Value = pistd3
        osheet.Range("pistd4").Value = pistd4
        osheet.Range("pvstd").Value = pvstd
        osheet.Range("pvstd1").Value = pvstd1
        osheet.Range("pvstd2").Value = pvstd2
        osheet.Range("pvstd3").Value = pvstd3
        osheet.Range("pvstd4").Value = pvstd4

        osheet.Range("nqsu").Value = nqsu
        osheet.Range("nqsu1").Value = nqsu1
        osheet.Range("nqsu2").Value = nqsu2
        osheet.Range("nqsu3").Value = nqsu3
        osheet.Range("nqsu4").Value = nqsu4
        osheet.Range("scsasl").Value = scsasl
        osheet.Range("scsasl1").Value = scsasl1
        osheet.Range("scsasl2").Value = scsasl2
        osheet.Range("scsasl3").Value = scsasl3
        osheet.Range("scsasl4").Value = scsasl4

        osheet.Range("scssl").Value = scssl
        osheet.Range("scssl1").Value = scssl1
        osheet.Range("scssl2").Value = scssl2
        osheet.Range("scssl3").Value = scssl3
        osheet.Range("scssl4").Value = scssl4

        osheet.Range("scsslnet").Value = scsslnet
        osheet.Range("scsslnet1").Value = scsslnet1
        osheet.Range("scsslnet2").Value = scsslnet2
        osheet.Range("scsslnet3").Value = scsslnet3
        osheet.Range("scsslnet4").Value = scsslnet4

        osheet.Range("sclt").Value = sclt
        osheet.Range("sclt1").Value = sclt1
        osheet.Range("sclt2").Value = sclt2
        osheet.Range("sclt3").Value = sclt3
        osheet.Range("sclt4").Value = sclt4

        osheet.Range("scno").Value = scno
        osheet.Range("scno1").Value = scno1
        osheet.Range("scno2").Value = scno2
        osheet.Range("scno3").Value = scno3
        osheet.Range("scno4").Value = scno4

        osheet.Range("scsh").Value = scsh
        osheet.Range("scsh1").Value = scsh1
        osheet.Range("scsh2").Value = scsh2
        osheet.Range("scsh3").Value = scsh3
        osheet.Range("scsh4").Value = scsh4

        If Not IsNothing(SocialAuditBS) Then
            Dim rowint As Integer = 42
            Dim max As Integer = 0
            For Each drv As DataRowView In SocialAuditBS.List
                Dim colint As Integer = 0
                For i As Integer = 0 To 7
                    colint = colint + 1
                    osheet.Cells(rowint, colint).value = "" & drv.Row.Item(i).ToString
                Next
                rowint = rowint + 1
                max = max + 1
                If max > 2 Then
                    Exit For
                End If
            Next

        End If

        If Not IsNothing(PanelStatusBS) Then
            Dim rowint As Integer = 47 '11
            Dim max As Integer = 0
            For Each drv As DataRowView In PanelStatusBS.List
                Dim colint As Integer = 1
                osheet.Cells(rowint, 2).value = "" & drv.Row.Item(0).ToString
                osheet.Cells(rowint, 3).value = "" & drv.Row.Item(1).ToString & " " & drv.Row.Item(2).ToString & " " & drv.Row.Item(3).ToString

                rowint = rowint + 1
                max = max + 1
                If max > 2 Then
                    Exit For
                End If
            Next
        End If
        Dim paymenttermsap As String = String.Empty
        If Not IsNothing(SAPBS) Then
            Dim drv As DataRowView = SAPBS.Current
            paymenttermsap = "" & drv.Row.Item(2).ToString
            osheet.Range("contractpaymentterms").Value = paymenttermsap
        End If
        If Not IsNothing(ContractBS) Then
            For Each drv As DataRowView In ContractBS.List
                Select Case drv.Item("doctypeid")
                    Case 32
                        osheet.Range("contractpaymentterms").Value = paymenttermsap '"" & drv.Item("payt") & " " & drv.Item("details")
                        osheet.Range("contractpaymenttermseffectivedate").Value = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("docdate"))
                    Case 33
                        osheet.Range("contractnqsu").Value = "" & drv.Item("nqsu")
                        osheet.Range("nqsueffectivedate").Value = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("docdate"))
                    Case 35
                        osheet.Range("contractsasl").Value = "" & drv.Item("sasl") & " %"
                        osheet.Range("contractleadtime").Value = "" & drv.Item("leadtime")
                        osheet.Range("sasleffectivedate").Value = "" & String.Format("{0:dd-MMM-yyyy}", drv.Item("docdate"))
                End Select
            Next
        End If

        'Assign Photos
        'Factory Photos
        owB.Worksheets(1).select()
        osheet = owB.Worksheets(1)


        PutPhotos(oXl, osheet, 1, "Factory")
        PutPhotos(oXl, osheet, 2, "Product")

        owB.Worksheets(3).select()
        osheet = owB.Worksheets(3)
        Dim query = From Photos In SupplierPhotoBS.List
        Where Photos.item("phototype") = 1
        Dim myrow As Integer = 40
        For Each drv As DataRowView In query
            osheet.Cells(myrow, 2) = drv.Item("description")
            myrow = myrow + 1
        Next

        query = From Photos In SupplierPhotoBS.List
        Where Photos.item("phototype") = 2
        myrow = 49
        For Each drv As DataRowView In query
            osheet.Cells(myrow, 2) = drv.Item("description")
            myrow = myrow + 1
        Next


        owb.Worksheets("Data").select()
        osheet = owb.Worksheets("Data")
        osheet.Move(Before:=owb.Worksheets(4))
        owb.Worksheets("Buyerinputimport").select()
        'osheet = owb.Worksheets("Buyerinputimport")
        osheet = owb.Worksheets("Buyerinputimport")
        osheet.Move(Before:=owb.Worksheets(5))
        'oXl.ActiveWindow.SelectedSheets.Visible = False
        owb.Worksheets("Purchasing strategy").select()
        osheet = owb.Worksheets("Purchasing strategy")
        oXl.ActiveWindow.SelectedSheets.Visible = False
        owb.Worksheets("RawData").select()
        osheet = owb.Worksheets("RawData")
        oXl.ActiveWindow.SelectedSheets.Visible = False
        owb.Worksheets(1).select()
        osheet = owb.Worksheets(1)
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        ComboBox1.Enabled = CheckBox1.Checked
        DateTimePicker1.Enabled = CheckBox1.Checked
        DateTimePicker2.Enabled = CheckBox1.Checked
    End Sub
End Class
