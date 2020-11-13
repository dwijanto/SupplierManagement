Imports Microsoft.Office.Interop
Imports SupplierManagement.PublicClass
Public Class ExportVendorInformationModification
    Private hd As DataRowView
    Private dtl As BindingSource
    Private parent As Object
    Public Sub New(ByVal Parent As Object, ByVal Header As DataRowView, ByVal Detail As BindingSource)
        hd = Header
        dtl = Detail
        Me.parent = Parent
    End Sub

    Public Sub GenerateExcel()
        Dim mymessage As String = String.Empty
        Dim mysaveform As New SaveFileDialog
        mysaveform.DefaultExt = "xlsx"
        mysaveform.FileName = String.Format("FormVendorInformationModification{0:yyyyMMdd}", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport       
            Dim myreport As New ExportToExcelFile(parent, filename, reportname, mycallback, datasheet, String.Format("{0}\{1}", HelperClass1.template, "VendorInformationModification.xltx"))
            myreport.CreateForm(parent, New System.EventArgs)
        End If
    End Sub

    Private Sub FormattingReport(ByRef obj As Object, ByRef e As System.EventArgs)
        Dim owb As Excel.Workbook = DirectCast(obj, Excel.Workbook)
        Dim oSheet As Excel.Worksheet = Nothing

        owb.Worksheets(1).select()
        oSheet = owb.Worksheets(1)  

        osheet.Range("vendorcode").Value = hd.Row.Item("vendorcode")
        oSheet.Range("vendorname").Value = hd.Row.Item("vendorname")
        oSheet.Range("vendorfamilycode").Value = hd.Row.Item("familycode")
        oSheet.Range("vendorsubfamilycode").Value = hd.Row.Item("subfamilycode")
        oSheet.Range("vendorcurrency").Value = hd.Row.Item("currency")
        oSheet.Range("referenceyearlyturnover").Value = String.Format("{0}, {1:#,##0}", hd.Row.Item("yearreference"), hd.Row.Item("turnovervalue"))
        oSheet.Range("ecoqualitycontact").Value = String.Format("{0} / {1}", hd.Row.Item("ecoqualitycontactname"), hd.Row.Item("ecoqualitycontactemail"))

        For Each dtdrv As DataRowView In dtl.List
            Select Case dtdrv.Row("fieldid")
                Case 1
                    oSheet.Range("companynameold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("companynamenew").Value = dtdrv.Row.Item("newvalue")
                Case 2
                    oSheet.Range("companyaddressold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("companyaddressnew").Value = dtdrv.Row.Item("newvalue")
                Case 3
                    oSheet.Range("companyaddressold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("companyaddressnew").Value = dtdrv.Row.Item("newvalue")
                Case 4
                    oSheet.Range("phonenoold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("phonenonew").Value = dtdrv.Row.Item("newvalue")
                Case 5
                    oSheet.Range("faxold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("faxnew").Value = dtdrv.Row.Item("newvalue")
                Case 6
                    oSheet.Range("mailaddressold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("mailaddressnew").Value = dtdrv.Row.Item("newvalue")
                Case 7
                    oSheet.Range("incotermold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("incotermnew").Value = dtdrv.Row.Item("newvalue")
                Case 8
                    oSheet.Range("paymenttermold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("paymenttermnew").Value = dtdrv.Row.Item("newvalue")
                Case 9
                    oSheet.Range("paymenttermeffectivedateold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("paymenttermeffectivedatenew").Value = dtdrv.Row.Item("newvalue")
                Case 10
                    oSheet.Range("benificiarynameold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("benificiarynamenew").Value = dtdrv.Row.Item("newvalue")
                Case 11
                    oSheet.Range("banknameold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("banknamenew").Value = dtdrv.Row.Item("newvalue")
                Case 12
                    oSheet.Range("bankaddressold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("bankaddressnew").Value = dtdrv.Row.Item("newvalue")
                Case 13
                    oSheet.Range("bankswiftcodeold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("bankswiftcodenew").Value = dtdrv.Row.Item("newvalue")
                Case 14
                    oSheet.Range("bankaccountnoold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("bankaccountnonew").Value = dtdrv.Row.Item("newvalue")
                Case 15
                    oSheet.Range("confirmationold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("confirmationnew").Value = dtdrv.Row.Item("newvalue")
                Case 16
                    oSheet.Range("cnapscodeold").Value = dtdrv.Row.Item("oldvalue")
                    oSheet.Range("cnapscodenew").Value = dtdrv.Row.Item("newvalue")

            End Select
        Next

        oSheet.Range("spm").Value = hd.Row.Item("spmusername")
        'oSheet.Range("spmdate").Value = hd.Row.Item("approvaldept2")
        oSheet.Range("pd").Value = hd.Row.Item("pdusername")
        'oSheet.Range("pddate").Value = hd.Row.Item("vendorname")
        oSheet.Range("dbmanager").Value = hd.Row.Item("dbusername")
        'oSheet.Range("dbmanagerdate").Value = hd.Row.Item("vendorname")
        oSheet.Range("fc").Value = hd.Row.Item("fcusername")
        'oSheet.Range("fcdate").Value = hd.Row.Item("vendorname")
        oSheet.Range("vp").Value = hd.Row.Item("vpusername")
        'oSheet.Range("vpdate").Value = hd.Row.Item("vendorname")

    End Sub

End Class
