Imports System.Text
Public MustInherit Class ClassBaseImportSIFIdentityAttachment
    Protected _siftx As BindingSource
    Protected _documentbs As BindingSource
    Protected _filename As String
    Protected myrecord() As String
    Protected myField As New Dictionary(Of Integer, Integer)
    Protected silabeldict As New Dictionary(Of Integer, String)

    Public errorMessage As String
    Private sb As StringBuilder

    Private Enum LabelId
        COMPANY_NAME = 1
        WEB_SITE = 2
        MAIN_CUSTOMER = 3
        SUPPLIER_PRODUCT_BRAND = 4
        MAIN_SHAREHOLDER = 5
        MAIN_PRODUCT_SOLD = 6
        OEM = 7
        ODM = 8
        IDENTITY_SHEET_PRODUCT = 9
        SBU1 = 10
        FAMILY_NAME_1 = 11
        SBU_2 = 12
        FAMILY_NAME_2 = 13
        OEM_ODM = 14
        MATERIAL_1 = 15
        MATERIAL_2 = 16
        ORIGIN_INFO = 17
        VISIT = 18
        _WHEN = 19
        WHO = 20
        PRODUCTION = 21
        CAPACITY_OF_PRODUCTION = 22
        NO_FACTORY = 23
        NO_WORKER = 24
        NO_ASSEMBLY_LINE = 25
        TOTAL_AREA = 26
        FACTORY_1 = 27
        FACTORY_2 = 28
        FACTORY_3 = 29
        TURNOVER_YEARS = 30
        SYNTHESIS = 31
        STATUS = 32
        OFFICE_ADDRESS = 33
        CONTACT = 34
        CONTACT_POSITION = 35
        CONTACT_MAIL = 36
    End Enum

    Public Sub New(ByRef SIFTx As BindingSource, ByRef DocumentBS As BindingSource, ByVal FileName As String)
        'Mapping Field
        Me._siftx = SIFTx 'Filtered by documentid
        Me._documentbs = DocumentBS
        Me._filename = FileName

        myField.Add(0, LabelId.COMPANY_NAME) '1
        myField.Add(1, LabelId.WEB_SITE) '2
        myField.Add(2, LabelId.MAIN_CUSTOMER) '3
        myField.Add(3, LabelId.MAIN_CUSTOMER) '3
        myField.Add(4, LabelId.MAIN_CUSTOMER) '3
        myField.Add(5, LabelId.SUPPLIER_PRODUCT_BRAND) '4
        myField.Add(6, LabelId.SUPPLIER_PRODUCT_BRAND) '4
        myField.Add(7, LabelId.SUPPLIER_PRODUCT_BRAND) '4
        myField.Add(8, LabelId.MAIN_SHAREHOLDER) '5
        myField.Add(9, LabelId.MAIN_SHAREHOLDER) '5
        myField.Add(10, LabelId.MAIN_SHAREHOLDER) '5
        myField.Add(11, 0) 'Share1
        myField.Add(12, 0) 'Share2
        myField.Add(13, 0) 'Share3
        myField.Add(14, 0) 'SIF Date
        myField.Add(15, 0) 'TO Y
        myField.Add(16, 0) 'TO Y-1
        myField.Add(17, 0) 'TO Y-2
        myField.Add(18, 0) 'TO Y-3
        myField.Add(19, 0) 'TO Y-4
        myField.Add(20, 0) '%SEB
        myField.Add(21, 0) '%SEB-1
        myField.Add(22, 0) '%SEB-2
        myField.Add(23, 0) '%SEB-3
        myField.Add(24, 0) '%SEB-4
        myField.Add(25, LabelId.OEM) '7
        myField.Add(26, LabelId.ODM) '8
        myField.Add(27, LabelId.MAIN_PRODUCT_SOLD) '6
        myField.Add(28, LabelId.MAIN_PRODUCT_SOLD) '6
        myField.Add(29, LabelId.MAIN_PRODUCT_SOLD) '6
        myField.Add(30, 0) 'Last Update
        myField.Add(31, LabelId.STATUS) '1
        myField.Add(32, LabelId.OFFICE_ADDRESS) '1
        myField.Add(33, LabelId.CONTACT) '1
        myField.Add(34, LabelId.CONTACT_POSITION) '1
        myField.Add(35, LabelId.CONTACT_MAIL) '1
        myField.Add(36, LabelId.VISIT) '1
        myField.Add(37, LabelId._WHEN) '1
        myField.Add(38, LabelId.WHO) '1
        myField.Add(39, LabelId.CAPACITY_OF_PRODUCTION) '1
        myField.Add(40, LabelId.NO_FACTORY) '1
        myField.Add(41, LabelId.NO_WORKER) '1
        myField.Add(42, LabelId.NO_ASSEMBLY_LINE) '1
        myField.Add(43, LabelId.TOTAL_AREA) '1
        myField.Add(44, LabelId.FACTORY_1) '1
        myField.Add(45, LabelId.FACTORY_1) '1
        myField.Add(46, LabelId.FACTORY_1) '1
        myField.Add(47, LabelId.SUPPLIER_PRODUCT_BRAND) '1
        myField.Add(48, LabelId.SBU1) '1
        myField.Add(49, LabelId.FAMILY_NAME_1) '1
        myField.Add(50, LabelId.SBU1) '1
        myField.Add(51, LabelId.FAMILY_NAME_1) '1
        myField.Add(52, LabelId.OEM_ODM) '1
        myField.Add(53, LabelId.MATERIAL_1) '1
        myField.Add(54, LabelId.MATERIAL_1) '1
        myField.Add(55, LabelId.ORIGIN_INFO) '1
        myField.Add(56, LabelId.TURNOVER_YEARS) '1
        myField.Add(57, LabelId.MAIN_CUSTOMER) '1
        myField.Add(58, LabelId.SYNTHESIS) '1

        silabeldict.Add(1, "Company Name")
        silabeldict.Add(2, "Website")
        silabeldict.Add(3, "Main Customer")
        silabeldict.Add(4, "Supplier Product Brand")
        silabeldict.Add(5, "Main Shareholder, % share")
        silabeldict.Add(6, "Main Product Sold")
        silabeldict.Add(7, "OEM %")
        silabeldict.Add(8, "ODM %")
        silabeldict.Add(9, "Identity Sheet Product")
        silabeldict.Add(10, "SBU")
        silabeldict.Add(11, "Family Name")
        silabeldict.Add(12, "SBU 2")
        silabeldict.Add(13, "Family Name")
        silabeldict.Add(14, "OEM ODM")
        silabeldict.Add(15, "Material")
        silabeldict.Add(16, "Material")
        silabeldict.Add(17, "Origin Info")
        silabeldict.Add(18, "Visit")
        silabeldict.Add(19, "When")
        silabeldict.Add(20, "Who")
        silabeldict.Add(21, "Production")
        silabeldict.Add(22, "Capacity of Production")
        silabeldict.Add(23, "No Factory")
        silabeldict.Add(24, "No Worker")
        silabeldict.Add(25, "No Assembly Line")
        silabeldict.Add(26, "Total Area")
        silabeldict.Add(27, "Factory")
        silabeldict.Add(28, "Factory")
        silabeldict.Add(29, "Factory")
        silabeldict.Add(30, "Turnover Years")
        silabeldict.Add(31, "Synthesis")
        silabeldict.Add(32, "Status")
        silabeldict.Add(33, "Office Address")
        silabeldict.Add(34, "Contact")
        silabeldict.Add(35, "Contact Position")
        silabeldict.Add(36, "Contact Mail")
    End Sub

    Private Function validate(ByRef myrecord() As String) As Boolean
        Dim MyRet As Boolean = True
        sb = New StringBuilder
        'if found error then assign to errormessage
        For i = 0 To 58
            If myrecord(i) <> "" Then
                myrecord(i) = myrecord(i).Replace(",", " ")
                Select Case i
                    Case 14
                        If Not IsDate(myrecord(14)) Then
                            sb.Append(String.Format("Invalid date value for SIF Date.""{0}""", myrecord(i)))
                            MyRet = False
                        End If
                    Case 15, 16, 17, 18, 19
                        If Not IsNumeric(myrecord(i)) Then
                            sb.Append(String.Format("Invalid numeric value for Turnover.""{0}""", myrecord(i)))
                            MyRet = False
                        End If
                    Case 20, 21, 22, 23, 24
                        If Not IsNumeric(myrecord(i)) Then
                            sb.Append(String.Format("Invalid numeric value for SIF% of SEB Total.""{0}""", myrecord(i)))
                            MyRet = False
                        End If
                    Case 30
                        If Not IsDate(myrecord(30)) Then
                            sb.Append(String.Format("Invalid last updating date value.""{0}""", myrecord(i)))
                            MyRet = False
                        End If
                End Select
            End If
        Next
        errorMessage = sb.ToString
        Return MyRet
    End Function
    Private Sub CleanData()
        Dim drv As DataRowView = _documentbs.Current
        drv.BeginEdit()

        drv.Row.Item("turnovery") = DBNull.Value
        drv.Row.Item("turnovery1") = DBNull.Value
        drv.Row.Item("turnovery2") = DBNull.Value
        drv.Row.Item("turnovery3") = DBNull.Value
        drv.Row.Item("turnovery4") = DBNull.Value
        drv.Row.Item("ratioy") = DBNull.Value
        drv.Row.Item("ratioy1") = DBNull.Value
        drv.Row.Item("ratioy2") = DBNull.Value
        drv.Row.Item("ratioy3") = DBNull.Value
        drv.Row.Item("ratioy4") = DBNull.Value
        drv.EndEdit()
        'Clean existing bindingSource
        If Not IsNothing(_siftx.Current) Then
            _siftx.Filter = String.Format("attachementid = {0}", drv.Item("id"))

            For Each mydrv As DataRowView In _siftx.List
                mydrv.BeginEdit()
                mydrv.Delete()
                mydrv.EndEdit()
            Next
            _siftx.Filter = ""
        End If
    End Sub

    Public Overridable Function getRecord() As Boolean
        Dim myret As Boolean = False

        CleanData()

        'Read Text file
        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        'Dim myrecord() As String
        Using objTFParser = New FileIO.TextFieldParser(_filename)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                Do Until .EndOfData
                    myrecord = .ReadFields

                    If count >= 1 Then
                        myret = validate(myrecord)
                    End If
                    count += 1
                Loop
            End With
        End Using

        Dim docdrv As DataRowView = _documentbs.Current
        For i = 0 To 58
            If myrecord(i) <> "" Then
                Try
                    If myField(i) <> 0 Then
                        Select Case i
                            Case 8, 9, 10
                                If myrecord(i + 3) <> "" Then
                                    myrecord(i + 3) = "," + myrecord(i + 3)
                                End If
                                myrecord(i) = myrecord(i) + myrecord(i + 3)
                        End Select
                        Dim sifdrv As DataRowView = _siftx.AddNew
                        sifdrv.BeginEdit()
                        sifdrv.Row.Item("attachmentid") = docdrv.Item("id")
                        sifdrv.Row.Item("labelid") = myField(i)
                        sifdrv.Row.Item("labelname") = silabeldict(myField(i))

                        sifdrv.Row.Item("value") = myrecord(i)



                        sifdrv.EndEdit()
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        Next

        myret = True
        Return myret
    End Function
End Class
