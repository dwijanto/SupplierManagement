Imports System.Threading
Imports SupplierManagement.PublicClass
Imports System.Text
Imports SupplierManagement.SharedClass
Imports Microsoft.Office.Interop

Public Class FormImportToolingList
    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    'Public Property DS As DataSet
    Dim ToolingListBS As BindingSource
    Dim DS As DataSet
    Dim DV As DataView
    Public mymessage As String = String.Empty

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByRef DS As DataSet)
        InitializeComponent()
        Me.DS = DS
        ToolingListBS = New BindingSource
        ToolingListBS.DataSource = DS.Tables("ToolingListDT")
        'DV = DS.Tables("ToolingListDT").AsDataView
        'DV.Sort = "suppliermoldno"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not mythread.IsAlive Then
            'Get file
            If openfiledialog1.ShowDialog = DialogResult.OK Then
                mythread = New Thread(AddressOf doWork)
                mythread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub doWork()
        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String


        'OpenExcelFile -> saveas txt
        Dim myret As Boolean = False
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Dim errormsg As String = String.Empty
        Dim FileNameTXT As String = String.Empty
        Dim Filename As String = String.Empty
        Try
            'Create Object Excel 
            Me.ProgressReport(5, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.ScreenUpdating = False
            oXl.Visible = False
            oXl.DisplayAlerts = False
            Me.ProgressReport(5, "Opening File")
            oWb = oXl.Workbooks.Open(openfiledialog1.FileName)
            Filename = openfiledialog1.FileName
            'FileNameTXT = IO.FileInfo(filename)openfiledialog1.FileName.Replace(".xlsx", ".txt")
            FileNameTXT = IO.Path.GetDirectoryName(Filename) & "\" & IO.Path.GetFileNameWithoutExtension(openfiledialog1.FileName) & ".txt"
            'oWb.SaveAs(Filename:=FileNameTXT, FileFormat:=Excel.XlFileFormat.xlTextMSDOS)
            oWb.SaveAs(Filename:=FileNameTXT, FileFormat:=Excel.XlFileFormat.xlUnicodeText)
            oXl.ScreenUpdating = True
            Dim lineno As Integer = 0
            If Not IsNothing(ToolingListBS.Current) Then
                ToolingListBS.MoveLast()
                Dim dr = ToolingListBS.Current
                lineno = dr.item("lineno")
            End If
            ' Using objTFParser = New FileIO.TextFieldParser(openfiledialog1.FileName)
            Using objTFParser = New FileIO.TextFieldParser(FileNameTXT)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    ProgressReport(1, "Read Data")
                    ProgressReport(2, "Read Data")

                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count >= 1 Then
                            If myrecord(4) <> "" Then 'Supplier Mould Number cannot be blank
                                'ToolingListBS Filtered by AssetPurchaseId
                                Dim found As Integer = ToolingListBS.Find("suppliermoldno", myrecord(4))
                                Dim drv As DataRowView
                                If found = -1 Then
                                    lineno = lineno + 1
                                    drv = ToolingListBS.AddNew()
                                    drv.BeginEdit()
                                    drv.Row.Item("lineno") = lineno

                                    drv.Row.Item("assetpurchaseid") = DS.Tables(0).Rows(0).Item("id")
                                    'if tooling modified then get from excel
                                    drv.Item("toolinglistid") = String.Format("{0}_{1:0000}", DS.Tables(0).Rows(0).Item("assetpurchaseid"), lineno)
                                    If DS.Tables(0).Rows(0).Item("typeofinvestment") = 2 Then
                                        Dim mytemp As String = String.Empty
                                        'mytemp = String.Format("{0}_{1:0000}", DS.Tables(0).Rows(0).Item("assetpurchaseid"), lineno)
                                        If myrecord(15) = "" Then
                                            If myrecord(14) <> "" Then
                                                'Has Original
                                                mytemp = validstrdb(myrecord(14))
                                                drv.Item("toolinglistid") = mytemp
                                                drv.Item("typeofinvestment") = 2
                                            Else
                                                'Missing Original, by pass checking
                                                If myrecord(15) = "" Then 'by pass for SETUP
                                                    If MessageBox.Show("Missing ToolingList id. The system will generate a new one. Is it ok?", "Missing Tooling List Id", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                                                        drv.Item("typeofinvestment") = 100 'no need to check for missing the original id 
                                                    Else
                                                        drv.CancelEdit()
                                                        Err.Raise(1, Description:="Import process cancelled by user.")
                                                    End If
                                                    'Else
                                                    '    mytemp = String.Format("{0}_{1:0000}", DS.Tables(0).Rows(0).Item("assetpurchaseid"), lineno)
                                                End If


                                            End If
                                        End If
                                        
                                        'drv.Item("toolinglistid") = mytemp
                                    Else
                                        'drv.Item("toolinglistid") = String.Format("{0}_{1:0000}", DS.Tables(0).Rows(0).Item("assetpurchaseid"), lineno)
                                        drv.Item("typeofinvestment") = 1
                                        If myrecord(16) <> "" Then
                                            If myrecord(14) = "" Then
                                                If MessageBox.Show("Missing ToolingList id for Common tool. The system will generate a new one. Is it ok?", "Missing Tooling List Id", MessageBoxButtons.OKCancel) = DialogResult.Cancel Then
                                                    drv.CancelEdit()
                                                    Err.Raise(1, Description:="Import process cancelled by user.")
                                                Else
                                                    drv.Item("typeofinvestment") = 0
                                                End If
                                            Else
                                                drv.Item("toolinglistid") = validstrdb(myrecord(14))

                                            End If
                                        End If

                                    End If

                                Else
                                    ToolingListBS.Position = found
                                    drv = ToolingListBS.Current
                                    If myrecord.Length > 14 Then
                                        If myrecord(14) <> "" Then
                                            drv.Item("toolinglistid") = validstrdb(myrecord(14))
                                        End If
                                    End If
                                    'drv.Item("toolingid") = String.Format("{0}_{1:0000}", DS.Tables(0).Rows(0).Item("assetpurchaseid"), drv.Item("lineno"))
                                End If

                                If myrecord.Length > 14 Then
                                    If myrecord(15) <> "" Then
                                        drv.Item("toolinglistid") = "S" + drv.Item("toolinglistid")

                                    End If
                                End If

                                'If myrecord.Length > 14 Then
                                '    If myrecord(16) <> "" And myrecord(14) = "" Then
                                '        MessageBox.Show("Tooling List Id cannot be blank for Common Tool.")
                                '        drv.CancelEdit()
                                '        Err.Raise(1, Description:="Import process rejected.")
                                '    End If
                                'End If

                                Try
                                    drv.Row.RowError = ""
                                    drv.Item("sebmodelno") = validstrdb(myrecord(0))
                                    drv.Item("suppliermodelreference") = validstrdb(myrecord(1))
                                    'drv.Item("sebasiaitemno") = myrecord(2)
                                    'drv.Item("sebasiaassetno") = myrecord(3)
                                    drv.Item("suppliermoldno") = validstrdb(myrecord(4))
                                    drv.Item("toolsdescription") = validstrdb(myrecord(5))
                                    drv.Item("material") = validstrdb(myrecord(6))
                                    drv.Item("cavities") = validstrdb(myrecord(7))
                                    drv.Item("numberoftools") = validintdb(myrecord(8))
                                    drv.Item("dailycaps") = validstrdb(myrecord(9))
                                    drv.Item("cost") = validnumdb(myrecord(10))
                                    drv.Item("balance") = validnumdb(myrecord(10))
                                    drv.Item("purchasedate") = validdatedb(myrecord(11))
                                    drv.Item("location") = validstrdb(myrecord(12))
                                    drv.Item("comments") = validstrdb(myrecord(13))
                                    drv.Item("vendorcode") = DS.Tables(0).Rows(0).Item("vendorcode")
                                    drv.Item("displaymember") = drv.Item("toolinglistid")
                                    drv.Item("commontool") = IIf(myrecord(16) <> "", True, False)
                                    drv.EndEdit()
                                Catch ex As Exception
                                    drv.CancelEdit()
                                    errormsg = ex.Message
                                    mymessage = errormsg
                                    ProgressReport(1, errormsg)
                                    Me.DialogResult = Windows.Forms.DialogResult.None
                                    ProgressReport(3, "Set Continuous Again")
                                    Exit Sub
                                End Try
                            Else
                                If MessageBox.Show("Missing Supplier Mould No, Skip the current record?", "Missing Supplier Mould", MessageBoxButtons.OKCancel) = DialogResult.Cancel Then
                                    Err.Raise(1, Description:="Import process cancelled by user.")
                                    Exit Sub
                                End If
                            End If
                        End If

                        count += 1
                    Loop
                End With
            End Using

            'update record
        'If myInsert.Length > 0 Then
        '    ProgressReport(1, "Start Add New Records")
        '    mystr.Append(String.Format("delete from doc.groupvendor where groupid = {0};", assetpurchaseid))
        '    Dim sqlstr As String = "copy doc.groupvendor(vendorcode,groupid) from stdin with null as 'Null';"
        '    Dim ra As Long = 0
        '    Dim errmessage As String = String.Empty
        '    Dim myret As Boolean = False

        '    Try
        '        If RadioButton1.Checked Then
        '            ProgressReport(1, "Replace Record Please wait!")
        '            ra = DbAdapter1.ExNonQuery(mystr.ToString)
        '        End If
        '        ProgressReport(1, "Add Record Please wait!")
        '        errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
        '        If myret Then
        '            ProgressReport(1, "Add Records Done.")
        '        Else
        '            ProgressReport(1, errmessage)
        '        End If
        '    Catch ex As Exception
        '        ProgressReport(1, ex.Message)

        '    End Try

        'End If
            ToolingListBS.EndEdit()
            ProgressReport(1, "Add Records Done.")
            ProgressReport(3, "Set Continuous Again")
            myret = True
            Me.DialogResult = Windows.Forms.DialogResult.OK
        Catch ex As Exception
            ToolingListBS.EndEdit()
            errormsg = ex.Message
            mymessage = errormsg
            ProgressReport(1, errormsg)
            Me.DialogResult = Windows.Forms.DialogResult.None
            ProgressReport(3, "Set Continuous Again")
        Finally
            'ProgressReport(3, "Releasing Memory...")
            'clear excel from memory
            oXl.Quit()
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

                Case 3
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Me.Close()
            End Select

        End If

    End Sub

  
End Class