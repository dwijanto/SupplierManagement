Imports SupplierManagement.SharedClass
Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Public Class UCSupplierDashboard

    Dim FamilyBSFilter As String
    Dim BrandBSFilter As String
    Dim RangeCodeBSFilter As String

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Dim myThread As New System.Threading.Thread(AddressOf DoQuery)
    Dim myThread2 As New System.Threading.Thread(AddressOf DoQuery2)
    Dim myThread3 As New System.Threading.Thread(AddressOf DoQuery3)
    Dim myThread4 As New System.Threading.Thread(AddressOf DoQuery4)
    Dim myThread5 As New System.Threading.Thread(AddressOf DoQuery5)

    Private SupplierInfoBS As BindingSource
    Public SupplierTechnologyBS As BindingSource
    Public SupplierPanelHistoryBS As BindingSource

    Public Event LoadFinish()
    Public Event QueryShortVendorname()
    Public Event FinishQuery3()
    Public Event FinishQuery5()

    Public Event StatusVendorChanged(ByVal criteria As String)

    Private sb As New StringBuilder
    Private DS As DataSet
    Private MasterDS As DataSet
    Private mymessage As String = String.Empty


    Public BSShortNameHelper As BindingSource
    Public BSVendorHelper As BindingSource


    Public BSSBUHelper As BindingSource
    Public BSFamilyHelper As BindingSource
    Public BSBrandHelper As BindingSource

    Public BSRangeHelper As BindingSource
    Public BSCMMFHelper As BindingSource
    Public BSCommRefHelper As BindingSource

    'Public SupplierInfo As Object
    Public SupplierInfo As UCSupplierInfo
    Public Property Vendorcode As Long
    Public Property ShortName As String
    Public Property VendorName As String
    Private _SDDS As DataSet
    Private SBUBS As BindingSource

    Private ProductBS As BindingSource
    Private StatusBS As BindingSource

    Public PanelStatus As BindingSource

    Public ComboBoxFilter As String = String.Empty
    Public comboboxCriteria As String = String.Empty

    'Public FamilyFilter As String = String.Empty
    Public SBUCriteria As String = String.Empty
    Public SBUFilter As String = String.Empty
    Public FamilyCriteria As String = String.Empty
    Public BrandCriteria As String = String.Empty
    Public CMMFCriteria As String = String.Empty
    Public RangeCodeCriteria As String = String.Empty
    Public CommercialCodeCriteria As String = String.Empty
    Public ProductTypeCriteria As String = String.Empty
    Public ProductTypeFilter As String = "ALL"
    Public ProductBSFilter As String
    Public SBUBSFilter As String

    Public StatusCriteria As String = String.Empty
    Public statusFilter As String = "ALL"
    Public StatusBSFilter As String
    Private SBUDS As DataSet
    Public Criteria As String = String.Empty
    Public StringFilter As String = String.Empty

    Private _criteria As String = String.Empty

    Public Property currentDate As Date
       

    Private textbox1ori As String
    Private textbox2ori As String
    Private textbox3ori As String

    ' Private SupplierInfoBS As BindingSource

    Enum QueryTypeEnum
        shortname
        vendorcode
    End Enum

    Private QueryType As QueryTypeEnum


    'Public Sub myloadfinish() Handles Me.LoadFinish

    'End Sub

    Public Property SDDS As DataSet
        Get
            Return _SDDS
        End Get
        Set(ByVal value As DataSet)
            _SDDS = value
            SBUBS = New BindingSource
            If Not IsNothing(_SDDS) Then
                'SBUBS.DataSource = _SDDS.Tables(0)
                'ComboBox1.DataSource = SBUBS
                'ComboBox1.DisplayMember = "SBU"
                'ComboBox1.ValueMember = "SBU"
            End If
            

        End Set
    End Property


    Public Sub GetPanelStatusSupplier()
        If Not myThread5.IsAlive Then
            myThread5 = New Thread(AddressOf DoQuery5)
            myThread5.Start()
        Else
            'MessageBox.Show(String.Format("{0}::{1}->Please wait until the current process is finished.", Me.Name, System.Reflection.MethodInfo.GetCurrentMethod()))
            'MessageBox.Show("Please wait until the current process is finished.")
        End If

    End Sub

    '******DO NOT USE THREAD FOR USER CUSTOM FORM*****




    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try


                Select Case id
                    Case 1

                    Case 2

                    Case 4
                        BSShortNameHelper = New BindingSource
                        BSVendorHelper = New BindingSource
                      

                        BSShortNameHelper.DataSource = DS.Tables(0)
                        BSVendorHelper.DataSource = DS.Tables(1)
                    Case 11
                        SupplierInfo.TextBox1.Text = textbox1ori
                        SupplierInfo.textbox2.text = textbox2ori
                        SupplierInfo.textbox3.Text = textbox3ori
                    Case 10
                        'Dim bs As New BindingSource
                        'bs.DataSource = DS.Tables(0)
                        'Dim drv = bs.Current
                        'If Not IsNothing(SupplierInfo) Then
                        '    Vendorcode = 0
                        '    ShortName = TextBox1.Text
                        '    VendorName = drv.item("vendorname")
                        '    SupplierInfo.Shortname = ShortName
                        '    SupplierInfo.TextBox1.Text = ShortName
                        '    SupplierInfo.TextBox2.Text = drv.item("vendorcode")
                        '    SupplierInfo.TextBox3.Text = VendorName
                        '    'textbox1ori = ShortName
                        '    'textbox2ori = drv.item("vendorcode")
                        '    'textbox3ori = VendorName
                        'End If
                        If Not IsNothing(SupplierInfo) Then
                            Dim drv = SupplierInfoBS.Current

                            Vendorcode = 0
                            ShortName = TextBox1.Text
                            VendorName = drv.item("vendorname")
                            SupplierInfo.Shortname = ShortName
                            SupplierInfo.TextBox1.Text = ShortName
                            SupplierInfo.TextBox2.Text = drv.item("vendorcode")
                            SupplierInfo.TextBox3.Text = VendorName
                            'textbox1ori = ShortName
                            'textbox2ori = drv.item("vendorcode")
                            'textbox3ori = VendorName

                        End If
                       
                    Case 5
                        'Dim bs As New BindingSource
                        'bs.DataSource = DS.Tables(0)
                        'Dim drv = bs.Current
                        Dim drv = SupplierInfoBS.Current
                        If Not IsNothing(SupplierInfo) Then
                            Vendorcode = 0
                            ShortName = TextBox1.Text
                            VendorName = drv.item("vendorname")
                            SupplierInfo.shortname = ShortName
                            SupplierInfo.TextBox1.Text = ShortName
                            SupplierInfo.textbox2.text = drv.item("vendorcode")
                            SupplierInfo.textbox3.Text = VendorName
                            textbox1ori = ShortName
                            textbox2ori = drv.item("vendorcode")
                            textbox3ori = VendorName

                        End If

                        PopulateInfo()


                    Case (6)
                        SBUBS = New BindingSource
                        ProductBS = New BindingSource

                        If Not IsNothing(SBUDS) Then
                            'SBUBS.DataSource = SBUDS.Tables(0)
                            ProductBS.DataSource = SBUDS.Tables(1)
                            BSSBUHelper = New BindingSource
                            BSFamilyHelper = New BindingSource
                            BSBrandHelper = New BindingSource
                            BSRangeHelper = New BindingSource
                            BSCMMFHelper = New BindingSource
                            BSCommRefHelper = New BindingSource

                            BSSBUHelper.DataSource = SBUDS.Tables(0)
                            BSFamilyHelper.DataSource = SBUDS.Tables(2)
                            BSBrandHelper.DataSource = SBUDS.Tables(3)
                            BSRangeHelper.DataSource = SBUDS.Tables(4)
                            BSCMMFHelper.DataSource = SBUDS.Tables(5)
                            BSCommRefHelper.DataSource = SBUDS.Tables(6)

                            'ComboBox1.DataSource = SBUBS
                            'ComboBox1.DisplayMember = "SBU"
                            'ComboBox1.ValueMember = "SBU"


                            CheckedListBox1.DataSource = ProductBS
                            CheckedListBox1.DisplayMember = "fpcp"

                            CheckedListBox1.SetItemChecked(0, True)
                            CheckedListBox_SelectedIndexChanged(CheckedListBox1, New System.EventArgs)
                            'CheckedListBox1.ValueMember = "fpcp"

                        End If

                        'Assign ProductType


                        RaiseEvent LoadFinish()
                    Case 7
                        ErrorProvider1.SetError(CheckedListBox1, "")
                    Case 8
                        ErrorProvider1.SetError(CheckedListBox1, "Please select from the list")
                    Case 9
                        'SBUBS = New BindingSource
                        ProductBS = New BindingSource
                        StatusBS = New BindingSource
                        If Not IsNothing(MasterDS) Then

                            ApplyBSFilter()
                            'BSSBUHelper = New BindingSource
                            'BSFamilyHelper = New BindingSource
                            'BSBrandHelper = New BindingSource
                            'BSRangeHelper = New BindingSource
                            'BSCMMFHelper = New BindingSource
                            'BSCommRefHelper = New BindingSource

                            Dim MyView As DataView = MasterDS.Tables(0).AsDataView
                            'MyView.Sort = "fpcp DESC"

                            Dim ProductTable = MyView.ToTable(True, {"fpcp"})
                            ProductBS.DataSource = ProductTable
                            Dim dr = ProductTable.NewRow
                            dr.Item(0) = "ALL"
                            ProductTable.Rows.InsertAt(dr, 0)


                            MyView.Sort = "status"
                            Dim StatusTable = MyView.ToTable(True, {"status"})

                            StatusBS.DataSource = StatusTable
                            dr = StatusTable.NewRow
                            dr.Item(0) = "ALL"
                            StatusTable.Rows.InsertAt(dr, 0)


                            'MyView.Sort = "sbu"
                            'Dim SBUTable = MyView.ToTable(True, {"sbu"})
                            'BSSBUHelper.DataSource = SBUTable
                            'dr = SBUTable.NewRow
                            'dr.Item(0) = "ALL"
                            'SBUTable.Rows.InsertAt(dr, 0)


                            'MyView.Sort = "familyname"
                            'Dim FamilyTable = MyView.ToTable(True, {"familyname", "comfam"})
                            'BSFamilyHelper.DataSource = FamilyTable
                            'dr = FamilyTable.NewRow
                            'dr.Item(0) = "ALL"
                            'dr.Item(1) = 0
                            'FamilyTable.Rows.InsertAt(dr, 0)


                            'MyView.Sort = "brand"
                            'Dim BrandTable = MyView.ToTable(True, {"brand"})
                            'BSBrandHelper.DataSource = BrandTable
                            'dr = BrandTable.NewRow
                            'dr.Item(0) = "ALL"
                            'BrandTable.Rows.InsertAt(dr, 0)

                            'MyView.Sort = "rangedescription"
                            'Dim RangeTable = MyView.ToTable(True, {"rangedescription", "range"})
                            'BSRangeHelper.DataSource = RangeTable
                            'dr = RangeTable.NewRow
                            'dr.Item(0) = "ALL"
                            'dr.Item(1) = ""
                            'RangeTable.Rows.InsertAt(dr, 0)

                            'MyView.Sort = "cmmfdescription"
                            'Dim CMMFTable = MyView.ToTable(True, {"cmmfdescription", "cmmf"})
                            'BSCMMFHelper.DataSource = CMMFTable
                            'dr = CMMFTable.NewRow
                            'dr.Item(0) = "ALL"
                            'dr.Item(1) = 0
                            'CMMFTable.Rows.InsertAt(dr, 0)

                            'MyView.Sort = "commrefdescription"
                            'Dim CommRefTable = MyView.ToTable(True, {"commrefdescription", "commref"})
                            'BSCommRefHelper.DataSource = CommRefTable
                            'dr = CommRefTable.NewRow
                            'dr.Item(0) = "ALL"
                            'dr.Item(1) = ""
                            'CommRefTable.Rows.InsertAt(dr, 0)

                            ProductBS.Position = 0
                            CheckedListBox1.DataSource = ProductBS
                            CheckedListBox1.DisplayMember = "fpcp"
                           

                            CheckedListBox1.SetItemChecked(0, True)
                            CheckedListBox_SelectedIndexChanged(CheckedListBox1, New System.EventArgs)
                            'ProductTypeCriteria = "ALL"
                            'CheckedListBox1.ValueMember = "fpcp"
                            CheckedListBox2.DataSource = Nothing
                            CheckedListBox2.Items.Clear()
                            StatusBS.Position = 1

                            CheckedListBox2.DataSource = StatusBS
                            CheckedListBox2.DisplayMember = "status"

                            CheckedListBox2.SetItemChecked(0, False)
                            If CheckedListBox2.Items.Count > 1 Then
                                CheckedListBox2.SetItemChecked(1, True)
                            Else
                                CheckedListBox2.SetItemChecked(0, True)
                            End If


                            CheckedListBox_SelectedIndexChanged(CheckedListBox2, New System.EventArgs)
                            CheckedListBox2_SelectedIndexChanged(CheckedListBox2, New System.EventArgs)
                            ApplyBSFilter()
                            'RaiseEvent StatusVendorChanged(StatusBSFilter)
                        End If
                        RaiseEvent LoadFinish()
                    Case 12
                        If Not IsNothing(SupplierInfo) Then
                            SupplierInfo.BS = PanelStatus
                            SupplierInfo.RefreshataGrid()

                        End If
                End Select
            Catch ex As Exception
                MessageBox.Show(String.Format("UCSupplierDashBoard {0}::{1} {2} {3}.", Me.Name, System.Reflection.MethodInfo.GetCurrentMethod(), id, ex.Message))
            End Try
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click, Button6.Click, Button7.Click, Button9.Click, Button10.Click
        Dim myobj As Button = CType(sender, Button)

        Select Case myobj.Name

            Case "Button1"

                QueryType = QueryTypeEnum.shortname
                If Not IsNothing(BSShortNameHelper) Then
                    Dim myform = New FormHelper(BSShortNameHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "shortname"

                    If myform.ShowDialog = DialogResult.OK Then
                        RaiseEvent QueryShortVendorname()

                        Dim drv As DataRowView = BSShortNameHelper.Current
                        If Not IsNothing(drv) Then
                            TextBox1.Text = "" + drv.Item("shortname")
                            ShortName = TextBox1.Text
                            TextBox2.Text = ""
                            If TextBox1.Text <> "" Then
                                getSupplierInfo()
                                'GetPanelStatusSupplier()
                            End If
                            'Criteria = String.Format("where shortname = '{0}' and tu.year >= {1} - 4", TextBox1.Text, Year(currentDate))
                            Criteria = String.Format("where shortname = '{0}'", TextBox1.Text)
                        End If

                    End If
                End If

                'RaiseEvent LoadFinish()
            Case "Button2"
                If Not IsNothing(BSVendorHelper) Then
                    QueryType = QueryTypeEnum.vendorcode
                    Dim myform = New FormHelper(BSVendorHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "description"
                    If myform.ShowDialog = DialogResult.OK Then
                        RaiseEvent QueryShortVendorname()
                        Dim drv As DataRowView = BSVendorHelper.Current
                        If Not IsNothing(drv) Then
                            If Not IsDBNull(drv.Item("vendorcode")) Then
                                Vendorcode = drv.Item("vendorcode")
                                ShortName = Nothing
                                VendorName = drv.Item("vendorname")
                                TextBox2.Text = VendorName
                                TextBox1.Text = ""
                                SupplierInfo.TextBox1.Text = "" + drv.Item("shortname")
                                SupplierInfo.TextBox2.Text = "" + Vendorcode.ToString
                                SupplierInfo.TextBox3.Text = "" + VendorName
                                SupplierInfo.Shortname = "" + drv.Item("shortname")
                                'criteria = String.Format("where vendorcode = {0}", TextBox2.Text)
                                Criteria = String.Format("where v.vendorcode = {0} and tu.year >= {1} - 4", drv.Item("vendorcode"), Year(currentDate))
                                'RaiseEvent LoadFinish()
                                'GetPanelStatusSupplier()
                                PopulateInfo()
                                'SupplierInfo.TextBox1.Text = "" + drv.Item("shortname")
                                'SupplierInfo.textbox2.text = "" + Vendorcode.ToString
                                'SupplierInfo.textbox3.Text = "" + VendorName
                                'SupplierInfo.shortname = "" + drv.Item("shortname")
                                'textbox1ori = "" + drv.Item("shortname")
                                ' textbox2ori = "" + Vendorcode.ToString
                                'textbox3ori = "" + VendorName
                            End If
                        End If


                    End If
                End If

            Case "Button10"
                Dim myform = New FormAdvancedSearchSupplier
                If myform.ShowDialog = DialogResult.OK Then
                    If Not IsNothing(myform.DocumentBS) Then
                        Dim drv = myform.DocumentBS.Current
                        If Not IsNothing(drv) Then
                            QueryType = QueryTypeEnum.vendorcode
                            ShortName = Nothing
                            VendorName = drv.Item("vendorname")
                            Vendorcode = drv.item("vendorcode")
                            TextBox2.Text = VendorName
                            TextBox1.Text = ""
                            SupplierInfo.TextBox1.Text = "" + drv.Item("shortname")
                            SupplierInfo.TextBox2.Text = "" + Vendorcode.ToString
                            SupplierInfo.TextBox3.Text = "" + VendorName
                            SupplierInfo.Shortname = "" + drv.Item("shortname")
                            'criteria = String.Format("where vendorcode = {0}", TextBox2.Text)
                            Criteria = String.Format("where v.vendorcode = {0}", drv.Item("vendorcode"))
                            textbox1ori = SupplierInfo.Shortname 'ShortName
                            textbox2ori = drv.item("vendorcode")
                            textbox3ori = VendorName
                            PopulateInfo()
                            'RaiseEvent LoadFinish()
                            'PopulateInfo()
                            'Button2.PerformClick()
                        End If
                    End If


                End If
            Case "Button9"
                Dim sbFilter As New StringBuilder ' Textbox Text
                Dim sbCriteria As New StringBuilder ' criteria
                Dim sbcriteriaNull As New StringBuilder
                Dim sb2 As New StringBuilder
                If Not IsNothing(BSSBUHelper) Then

                    If TextBox9.Text <> "ALL" Then
                        sbFilter.Append(TextBox9.Text)
                    End If
                    Dim myform = New FormHelper(BSSBUHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "sbu"

                    If myform.ShowDialog = DialogResult.OK Then
                        For Each sel As DataGridViewRow In myform.DataGridView1.SelectedRows
                            'Check ALL
                            If sel.Cells(0).FormattedValue = "ALL" Then
                                sbCriteria.Clear()
                                sbFilter.Clear()
                                sbFilter.Append("ALL")
                                Exit For
                            End If
                            If sel.Cells(0).FormattedValue = "" Then
                                sbcriteriaNull.Append(" tu.sbu isnull")
                            ElseIf Not (TextBox9.Text.Contains(sel.Cells(0).FormattedValue)) Then
                                If sbFilter.Length > 0 Then
                                    sbFilter.Append(",")
                                    'sbCriteria.Append(",")
                                End If
                                sbFilter.Append(sel.Cells(0).FormattedValue)
                                'sbCriteria.Append(String.Format("''{0}''", sel.Cells(0).FormattedValue))
                            End If
                        Next

                        TextBox9.Text = sbFilter.ToString
                        Dim mylist = TextBox9.Text.ToString.Split(",")
                        For i = 0 To mylist.Length - 1
                            If sbCriteria.Length > 0 Then
                                sbCriteria.Append(",")
                            End If
                            If mylist(i) <> "" Then
                                sbCriteria.Append(String.Format("''{0}''", mylist(i)))
                            End If

                        Next

                        If TextBox9.Text = "ALL" Then
                            SBUFilter = "ALL"
                            SBUCriteria = ""
                            SBUBSFilter = ""
                        Else

                            SBUBSFilter = String.Format("{0} {1}", IIf(sbCriteria.Length > 0, String.Format("sbu in ({0})", sbCriteria.ToString.Replace("''", "'")), ""), IIf(sbcriteriaNull.Length > 0, IIf(sbCriteria.Length > 0, " or ", "") & " sbu is null", ""))
                            SBUCriteria = String.Format(" and ({0} {1})", IIf(sbCriteria.Length > 0, "tu.sbu in (" & sbCriteria.ToString & ")" & IIf(sbcriteriaNull.Length > 0, " or ", "") & sbcriteriaNull.ToString, ""), sbcriteriaNull.ToString)
                            SBUFilter = String.Format("SBU : {0}{1}", sbFilter.ToString, IIf(sbFilter.Length > 0 And sbcriteriaNull.Length > 0, ",", "") & IIf(sbcriteriaNull.Length > 0, "NULL", ""))

                        End If
                    End If
                End If
                ApplyBSFilter()

            Case "Button3"
                Dim sbFilter As New StringBuilder ' Textbox Text
                Dim sbFilter2 As New StringBuilder ' Id of textbox
                Dim sbCriteria As New StringBuilder ' criteria
                Dim sbcriteriaNull As New StringBuilder
                If Not IsNothing(BSFamilyHelper) Then

                    If TextBox4.Text <> "ALL" Then
                        sbFilter.Append(TextBox4.Text)
                    End If
                    Dim myform = New FormHelper(BSFamilyHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "familyname"
                    myform.DataGridView1.Columns(1).DataPropertyName = "comfam"

                    If myform.ShowDialog = DialogResult.OK Then
                        For Each sel As DataGridViewRow In myform.DataGridView1.SelectedRows
                            'Check ALL
                            If sel.Cells(0).FormattedValue = "ALL" Then
                                sbCriteria.Clear()
                                sbFilter.Clear()
                                sbFilter.Append("ALL")
                                Exit For
                            End If
                            If sel.Cells(0).FormattedValue = "" Then
                                sbcriteriaNull.Append(" and tu.comfam isnull")
                            ElseIf Not (TextBox4.Text.Contains(sel.Cells(0).FormattedValue)) Then
                                If sbFilter.Length > 0 Then
                                    sbFilter.Append(",")
                                    sbFilter2.Append(",")
                                    'sbCriteria.Append(",")
                                End If
                                sbFilter.Append(sel.Cells(0).FormattedValue)
                                sbFilter2.Append(sel.Cells(1).FormattedValue)
                                'sbCriteria.Append(String.Format("''{0}''", sel.Cells(0).FormattedValue))
                            End If
                        Next

                        TextBox4.Text = sbFilter.ToString
                        Dim mylist = TextBox4.Text.ToString.Split(",")
                        For i = 0 To mylist.Length - 1
                            If sbCriteria.Length > 0 Then
                                sbCriteria.Append(",")
                            End If
                            If mylist(i) <> "" Then
                                sbCriteria.Append(String.Format("''{0}''", mylist(i)))
                            End If

                        Next

                        If TextBox4.Text = "ALL" Then
                            FamilyCriteria = ""
                            FamilyBSFilter = ""
                        Else

                            FamilyCriteria = String.Format(" and ({0} {1})", IIf(sbCriteria.Length > 0, "familyname in (" & sbCriteria.ToString & ")" & IIf(sbcriteriaNull.Length > 0, " or ", "") & sbcriteriaNull.ToString, ""), sbcriteriaNull.ToString)
                            FamilyBSFilter = String.Format("{0} {1}", IIf(sbCriteria.Length > 0, String.Format("familyname in ({0})", sbCriteria.ToString.Replace("''", "'")), ""), IIf(sbcriteriaNull.Length > 0, IIf(sbCriteria.Length > 0, " or ", "") & " familyname is null", ""))
                        End If
                    End If
                    ApplyBSFilter()
                End If
                'If Not IsNothing(BSFamilyHelper) Then
                '    Dim myform = New FormHelper(BSFamilyHelper)
                '    myform.DataGridView1.Columns(0).DataPropertyName = "familyname"

                '    If myform.ShowDialog = DialogResult.OK Then


                '        Dim drv As DataRowView = BSFamilyHelper.Current
                '        TextBox4.Text = "" + drv.Item("familyname")

                '        If IsDBNull(drv.Item("comfam")) Then
                '            FamilyCriteria = String.Format(" and tu.comfam isnull")
                '        Else
                '            If drv.Item("familyname") = "ALL" Then
                '                FamilyCriteria = ""
                '            Else
                '                FamilyCriteria = String.Format(" and tu.comfam = {0}", drv.Item("comfam"))
                '            End If

                '        End If

                '    End If
                'End If

            Case "Button4"
                Dim sbFilter As New StringBuilder ' Textbox Text
                Dim sbFilter2 As New StringBuilder ' Id of textbox
                Dim sbCriteria As New StringBuilder ' criteria
                Dim sbcriteriaNull As New StringBuilder
                If Not IsNothing(BSBrandHelper) Then

                    If TextBox3.Text <> "ALL" Then
                        sbFilter.Append(TextBox3.Text)
                    End If
                    Dim myform = New FormHelper(BSBrandHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "brand"
                    'myform.DataGridView1.Columns(1).DataPropertyName = "comfam"

                    If myform.ShowDialog = DialogResult.OK Then
                        For Each sel As DataGridViewRow In myform.DataGridView1.SelectedRows
                            'Check ALL
                            If sel.Cells(0).FormattedValue = "ALL" Then
                                sbCriteria.Clear()
                                sbFilter.Clear()
                                sbFilter.Append("ALL")
                                Exit For
                            End If
                            If sel.Cells(0).FormattedValue = "" Then
                                sbcriteriaNull.Append(" and tu.brand isnull")
                            ElseIf Not (TextBox3.Text.Contains(sel.Cells(0).FormattedValue)) Then
                                If sbFilter.Length > 0 Then
                                    sbFilter.Append(",")
                                    sbFilter2.Append(",")
                                    'sbCriteria.Append(",")
                                End If
                                sbFilter.Append(sel.Cells(0).FormattedValue)
                                sbFilter2.Append(sel.Cells(1).FormattedValue)
                                'sbCriteria.Append(String.Format("''{0}''", sel.Cells(0).FormattedValue))
                            End If
                        Next

                        TextBox3.Text = sbFilter.ToString
                        Dim mylist = TextBox3.Text.ToString.Split(",")
                        For i = 0 To mylist.Length - 1
                            If sbCriteria.Length > 0 Then
                                sbCriteria.Append(",")
                            End If
                            If mylist(i) <> "" Then
                                sbCriteria.Append(String.Format("''{0}''", mylist(i)))
                            End If

                        Next

                        If TextBox3.Text = "ALL" Then
                            BrandCriteria = ""
                            BrandBSFilter = ""
                        Else

                            BrandCriteria = String.Format(" and ({0} {1})", IIf(sbCriteria.Length > 0, "tu.brand in (" & sbCriteria.ToString & ")" & IIf(sbcriteriaNull.Length > 0, " or ", "") & sbcriteriaNull.ToString, ""), sbcriteriaNull.ToString)
                            BrandBSFilter = String.Format("{0} {1}", IIf(sbCriteria.Length > 0, String.Format("brand in ({0})", sbCriteria.ToString.Replace("''", "'")), ""), IIf(sbcriteriaNull.Length > 0, IIf(sbCriteria.Length > 0, " or ", "") & " brand is null", ""))
                        End If
                    End If
                    ApplyBSFilter()
                End If
                'If Not IsNothing(BSBrandHelper) Then
                '    Dim myform = New FormHelper(BSBrandHelper)
                '    myform.DataGridView1.Columns(0).DataPropertyName = "brand"

                '    If myform.ShowDialog = DialogResult.OK Then


                '        Dim drv As DataRowView = BSBrandHelper.Current
                '        TextBox3.Text = "" + drv.Item("brand")

                '        If IsDBNull(drv.Item("brand")) Then
                '            BrandCriteria = String.Format(" and tu.brand isnull")
                '        Else
                '            If drv.Item("brand") = "ALL" Then
                '                BrandCriteria = ""
                '            Else
                '                BrandCriteria = String.Format("and tu.brand = ''{0}''", drv.Item("brand"))
                '            End If

                '        End If


                '    End If
                'End If

            Case "Button5"
                Dim sbFilter As New StringBuilder ' Textbox Text
                Dim sbFilter2 As New StringBuilder ' Id of textbox
                Dim sbCriteria As New StringBuilder ' criteria
                Dim sbcriteriaNull As New StringBuilder
                If Not IsNothing(BSRangeHelper) Then

                    If TextBox6.Text <> "ALL" Then
                        sbFilter2.Append(TextBox6.Text)
                    End If
                    Dim myform = New FormHelper(BSRangeHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "rangedescription"
                    myform.DataGridView1.Columns(1).DataPropertyName = "range"

                    If myform.ShowDialog = DialogResult.OK Then
                        For Each sel As DataGridViewRow In myform.DataGridView1.SelectedRows
                            'Check ALL
                            If sel.Cells(0).FormattedValue = "ALL" Then
                                sbCriteria.Clear()
                                sbFilter.Clear()
                                sbFilter2.Clear()
                                sbFilter2.Append("ALL")
                                Exit For
                            End If
                            If sel.Cells(0).FormattedValue = "" Then
                                sbcriteriaNull.Append(" and tu.range isnull")
                            ElseIf Not (TextBox6.Text.Contains(sel.Cells(0).FormattedValue)) Then
                                If sbFilter2.Length > 0 Then
                                    sbFilter.Append(",")
                                    sbFilter2.Append(",")
                                    'sbCriteria.Append(",")
                                End If
                                sbFilter.Append(sel.Cells(0).FormattedValue)
                                sbFilter2.Append(sel.Cells(1).FormattedValue)
                                'sbCriteria.Append(String.Format("''{0}''", sel.Cells(0).FormattedValue))
                            End If
                        Next

                        TextBox6.Text = sbFilter2.ToString


                        Dim mylist = TextBox6.Text.ToString.Split(",")
                        For i = 0 To mylist.Length - 1
                            If sbCriteria.Length > 0 Then
                                sbCriteria.Append(",")
                            End If
                            If mylist(i) <> "" Then
                                sbCriteria.Append(String.Format("''{0}''", mylist(i)))
                            End If

                        Next

                        If TextBox6.Text = "ALL" Then
                            RangeCodeCriteria = ""
                            RangeCodeBSFilter = ""
                            'SBUFilter = "ALL"
                        Else

                            RangeCodeCriteria = String.Format(" and ({0} {1})", IIf(sbCriteria.Length > 0, "tu.range in (" & sbCriteria.ToString & ")" & IIf(sbcriteriaNull.Length > 0, " or ", "") & sbcriteriaNull.ToString, ""), sbcriteriaNull.ToString)
                            RangeCodeBSFilter = String.Format("{0} {1}", IIf(sbCriteria.Length > 0, String.Format("range in ({0})", sbCriteria.ToString.Replace("''", "'")), ""), IIf(sbcriteriaNull.Length > 0, IIf(sbCriteria.Length > 0, " or ", "") & " range is null", ""))
                        End If
                    End If
                    ApplyBSFilter()
                End If
                'If Not IsNothing(BSRangeHelper) Then
                '    Dim myform = New FormHelper(BSRangeHelper)
                '    myform.DataGridView1.Columns(0).DataPropertyName = "description"

                '    If myform.ShowDialog = DialogResult.OK Then


                '        Dim drv As DataRowView = BSRangeHelper.Current
                '        TextBox6.Text = "" + drv.Item("range")

                '        If IsDBNull(drv.Item("range")) Then
                '            RangeCodeCriteria = String.Format(" and tu.range isnull")
                '        Else
                '            If drv.Item("range") = "ALL" Then
                '                RangeCodeCriteria = ""
                '            Else
                '                RangeCodeCriteria = String.Format(" and tu.range = ''{0}''", drv.Item("range"))
                '            End If

                '        End If


                '    End If
                'End If
            Case "Button6"
                Dim sbFilter As New StringBuilder ' Textbox Text
                Dim sbFilter2 As New StringBuilder ' Id of textbox
                Dim sbCriteria As New StringBuilder ' criteria
                Dim sbcriteriaNull As New StringBuilder
                If Not IsNothing(BSCMMFHelper) Then

                    If TextBox5.Text <> "ALL" Then
                        sbFilter2.Append(TextBox5.Text)
                    End If
                    Dim myform = New FormHelper(BSCMMFHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "cmmfdescription"
                    myform.DataGridView1.Columns(1).DataPropertyName = "cmmf"

                    If myform.ShowDialog = DialogResult.OK Then
                        For Each sel As DataGridViewRow In myform.DataGridView1.SelectedRows
                            'Check ALL
                            If sel.Cells(0).FormattedValue = "ALL" Then
                                sbCriteria.Clear()
                                sbFilter.Clear()
                                sbFilter2.Clear()
                                sbFilter2.Append("ALL")
                                Exit For
                            End If
                            If sel.Cells(0).FormattedValue = "" Then
                                sbcriteriaNull.Append(" and tu.cmmf isnull")
                            ElseIf Not (TextBox5.Text.Contains(sel.Cells(0).FormattedValue)) Then
                                If sbFilter2.Length > 0 Then
                                    sbFilter.Append(",")
                                    sbFilter2.Append(",")
                                    'sbCriteria.Append(",")
                                End If
                                sbFilter.Append(sel.Cells(0).FormattedValue)
                                sbFilter2.Append(sel.Cells(1).FormattedValue)
                                'sbCriteria.Append(String.Format("''{0}''", sel.Cells(0).FormattedValue))
                            End If
                        Next

                        TextBox5.Text = sbFilter2.ToString
                        Dim mylist = TextBox5.Text.ToString.Split(",")
                        For i = 0 To mylist.Length - 1
                            If sbCriteria.Length > 0 Then
                                sbCriteria.Append(",")
                            End If
                            If mylist(i) <> "" Then
                                sbCriteria.Append(String.Format("''{0}''", mylist(i)))
                            End If

                        Next

                        If TextBox5.Text = "ALL" Then
                            CMMFCriteria = ""
                            'SBUFilter = "ALL"
                        Else

                            CMMFCriteria = String.Format(" and ({0} {1})", IIf(sbCriteria.Length > 0, "tu.cmmf in (" & sbCriteria.ToString & ")" & IIf(sbcriteriaNull.Length > 0, " or ", "") & sbcriteriaNull.ToString, ""), sbcriteriaNull.ToString)
                            'SBUFilter = String.Format("SBU : {0}{1}", sbFilter.ToString, IIf(sbFilter.Length > 0 And sbcriteriaNull.Length > 0, ",", "") & IIf(sbcriteriaNull.Length > 0, "NULL", ""))
                        End If
                    End If
                End If
                'If Not IsNothing(BSCMMFHelper) Then
                '    Dim myform = New FormHelper(BSCMMFHelper)
                '    myform.DataGridView1.Columns(0).DataPropertyName = "description"

                '    If myform.ShowDialog = DialogResult.OK Then


                '        Dim drv As DataRowView = BSCMMFHelper.Current
                '        TextBox5.Text = "" + IIf(drv.Item("cmmf").ToString = "0", "ALL", drv.Item("cmmf").ToString)

                '        If IsDBNull(drv.Item("cmmf")) Then
                '            CMMFCriteria = String.Format(" and tu.cmmf isnull")
                '        Else
                '            If drv.Item("cmmf") = 0 Then
                '                CMMFCriteria = ""
                '            Else
                '                CMMFCriteria = String.Format(" and tu.cmmf = {0}", drv.Item("cmmf"))
                '            End If

                '        End If


                '    End If
                'End If
            Case "Button7"
                Dim sbFilter As New StringBuilder ' Textbox Text
                Dim sbFilter2 As New StringBuilder ' Id of textbox
                Dim sbCriteria As New StringBuilder ' criteria
                Dim sbcriteriaNull As New StringBuilder
                If Not IsNothing(BSCommRefHelper) Then

                    If TextBox8.Text <> "ALL" Then
                        sbFilter2.Append(TextBox8.Text)
                    End If
                    Dim myform = New FormHelper(BSCommRefHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "commrefdescription"
                    myform.DataGridView1.Columns(1).DataPropertyName = "commref"

                    If myform.ShowDialog = DialogResult.OK Then
                        For Each sel As DataGridViewRow In myform.DataGridView1.SelectedRows
                            'Check ALL
                            If sel.Cells(0).FormattedValue = "ALL" Then
                                sbCriteria.Clear()
                                sbFilter.Clear()
                                sbFilter2.Clear()
                                sbFilter2.Append("ALL")
                                Exit For
                            End If
                            If sel.Cells(0).FormattedValue = "" Then
                                sbcriteriaNull.Append(" and comref isnull")
                            ElseIf Not (TextBox8.Text.Contains(sel.Cells(0).FormattedValue)) Then
                                If sbFilter2.Length > 0 Then
                                    sbFilter.Append(",")
                                    sbFilter2.Append(",")
                                    'sbCriteria.Append(",")
                                End If
                                sbFilter.Append(sel.Cells(0).FormattedValue)
                                sbFilter2.Append(sel.Cells(1).FormattedValue)
                                'sbCriteria.Append(String.Format("''{0}''", sel.Cells(0).FormattedValue))
                            End If
                        Next

                        TextBox8.Text = sbFilter2.ToString
                        Dim mylist = TextBox8.Text.ToString.Split(",")
                        For i = 0 To mylist.Length - 1
                            If sbCriteria.Length > 0 Then
                                sbCriteria.Append(",")
                            End If
                            If mylist(i) <> "" Then
                                sbCriteria.Append(String.Format("''{0}''", mylist(i)))
                            End If

                        Next

                        If TextBox8.Text = "ALL" Then
                            CommercialCodeCriteria = ""
                            'SBUFilter = "ALL"
                        Else

                            CommercialCodeCriteria = String.Format(" and ({0} {1})", IIf(sbCriteria.Length > 0, " commref in (" & sbCriteria.ToString & ")" & IIf(sbcriteriaNull.Length > 0, " or ", "") & sbcriteriaNull.ToString, ""), sbcriteriaNull.ToString)
                            'SBUFilter = String.Format("SBU : {0}{1}", sbFilter.ToString, IIf(sbFilter.Length > 0 And sbcriteriaNull.Length > 0, ",", "") & IIf(sbcriteriaNull.Length > 0, "NULL", ""))
                        End If
                    End If
                End If
                'If Not IsNothing(BSCommRefHelper) Then
                '    Dim myform = New FormHelper(BSCommRefHelper)
                '    myform.DataGridView1.Columns(0).DataPropertyName = "description"

                '    If myform.ShowDialog = DialogResult.OK Then


                '        Dim drv As DataRowView = BSCommRefHelper.Current
                '        TextBox8.Text = "" + drv.Item("commref")

                '        If IsDBNull(drv.Item("commref")) Then
                '            CommercialCodeCriteria = String.Format(" and commref isnull")
                '        Else
                '            If drv.Item("commref") = "ALL" Then
                '                CommercialCodeCriteria = ""
                '            Else
                '                CommercialCodeCriteria = String.Format(" and commref = ''{0}''", drv.Item("commref"))
                '            End If

                '        End If


                '    End If
                'End If

        End Select
        'Populate Info
        '

        'clear
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try


    End Sub
   
    Private Sub getSupplierInfo()
        If Not myThread.IsAlive Then
            myThread = New Thread(AddressOf DoQuery)
            myThread.Start()
        Else
            'MessageBox.Show(String.Format("{0}::{1}->Please wait until the current process is finished.", Me.Name, System.Reflection.MethodInfo.GetCurrentMethod()))
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Public Sub getSupplierInfo2(ByVal _criteria As String)
        'Get supplier based on status
        If Not myThread4.IsAlive Then
            Me._criteria = _criteria
            myThread4 = New Thread(AddressOf DoQuery4)
            myThread4.Start()
        Else
            'MessageBox.Show(String.Format("{0}::{1}->Please wait until the current process is finished.", Me.Name, System.Reflection.MethodInfo.GetCurrentMethod()))
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
    Public Sub ShowOriginal()
        ProgressReport(11, "Assign TextBox")
    

    End Sub
    Private Sub DoQuery4()
        sb.Clear()
        If _criteria <> "" Then
            sb.Append(String.Format("select doc.getvendorcode('{0}','{1}') as vendorcode,doc.getvendorname('{0}','{1}') as vendorname;", TextBox1.Text.Replace("'", "''"), _criteria.Replace("'", "")))
        Else
            sb.Append(String.Format("select doc.getvendorcode('{0}') as vendorcode,doc.getvendorname('{0}') as vendorname;", TextBox1.Text.Replace("'", "''")))
        End If

        Dim Q4DS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString, Q4DS, mymessage) Then
            Try
                Q4DS.Tables(0).TableName = "ShortName"
                SupplierInfoBS.DataSource = Q4DS.Tables(0)
                ProgressReport(10, "Assign TextBox")
            Catch ex As Exception
                ProgressReport(1, "UCSupplierDashboard Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
        End If
    End Sub

    Private Sub DoQuery()
        sb.Clear()
        sb.Append(String.Format("select doc.getvendorcode('{0}') as vendorcode,doc.getvendorname('{0}') as vendorname;", TextBox1.Text.Replace("'", "''")))

        DS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                DS.Tables(0).TableName = "ShortName"
                SupplierInfoBS = New BindingSource
                SupplierInfoBS.DataSource = DS.Tables(0)
                ProgressReport(5, "Assign TextBox")
            Catch ex As Exception
                ProgressReport(1, "UCSupplierDashboard Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
        End If
    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged, TextBox2.TextAlignChanged
        ErrorProvider1.SetError(TextBox1, "")
        ErrorProvider1.SetError(TextBox2, "")
    End Sub

    Private Sub PopulateInfo()
        GetPanelStatusSupplier()
        If Not myThread2.IsAlive Then
            'myThread2 = New Thread(AddressOf DoQuery2)
            myThread2 = New Thread(AddressOf DoQuery3)
            myThread2.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
    Private Sub DoQuery3()
        'Get MasterDS based on Query producttype,sbu,family,brand,range,cmmf,comfam based on criteria(vendorcode/shortname, selection date)
        'ProductTypeSB -> distinct Producttype from MasterDS
        'SBU -> Distinct SBU from Masterds -> Applied Filter ProductTypeFilter
        'Family -> Distinct Family from MasterDS -> Applied Filter ProductTypeFilter, SBUFilter
        'Brand -> Distinct Brand from MasterDS -> Applied Filter ProductTypeFilter, SBUFilter, FamilyFilter
        'Range -> Distinct Range from MasterDS -> Applied Filter ProductTypeFilter, SBUFilter, FamilyFilter
        'CMMF -> Distinct CMMF from MasterDS -> Applied Filter ProdutTypeFilter, SBUFilter,FamilyFilter,RangeFilter
        'CommRef -> Distinct COMMRef from MasterDS -> Applied Filter ProductFilter, SBUFilter,FamilyFilter,RangeFilter,CommRef
        sb.Clear()
        

        'sb.Append(String.Format("select distinct tu.fpcp,tu.sbu,f.familyname::text,tu.brand,tu.range,tu.cmmf,comfam , tu.cmmf || ' - ' || mm.materialdesc as cmmfdescription,tu.range::text || ' - ' || r.rangedesc as rangedescription,mm.commref, mm.commref || ' - ' || mm.materialdesc as commrefdescription,sf.cmmf as sebplatform,pr.paramname as status from doc.turnover tu " &
        '                       " left join vendor v on v.vendorcode = tu.vendorcode" &
        '                       " left join family f on f.familyid = tu.comfam" &
        '                       " left join materialmaster mm on mm.cmmf = tu.cmmf" &
        '                       " left join range r on r.range = tu.range" &
        '                       " left join doc.sebplatform sf on sf.cmmf = tu.cmmf" &
        '                       " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
        '                       " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
        '                       " {0}" &
        '                       " order by sbu;", Criteria))
        'sb.Append(String.Format("(select distinct tu.fpcp,tu.sbu,f.familyname::text,tu.brand,tu.range,tu.cmmf,comfam , tu.cmmf || ' - ' || mm.materialdesc as cmmfdescription,tu.range::text || ' - ' || r.rangedesc as rangedescription,mm.commref, mm.commref || ' - ' || mm.materialdesc as commrefdescription,sf.cmmf as sebplatform,pr.paramname as status from doc.turnover tu " &
        '               " left join vendor v on v.vendorcode = tu.vendorcode" &
        '               " left join family f on f.familyid = tu.comfam" &
        '               " left join materialmaster mm on mm.cmmf = tu.cmmf" &
        '               " left join range r on r.range = tu.range" &
        '               " left join doc.sebplatform sf on sf.cmmf = tu.cmmf" &
        '               " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
        '               " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
        '               " {0}" &
        '               " order by sbu)" &
        '               " union all (select distinct null::text as fpcp,bf.sbu,f.familyname::text,bf.brand,bf.range,bf.cmmf,comfam , bf.cmmf || ' - ' || mm.materialdesc as cmmfdescription,bf.range::text || ' - ' || r.rangedesc as rangedescription,mm.commref, mm.commref || ' - ' || mm.materialdesc as commrefdescription,sf.cmmf as sebplatform,pr.paramname as status   from doc.budgetforecast bf" &
        '               " left join vendor v on v.vendorcode = bf.vendorcode " &
        '               " left join family f on f.familyid = bf.comfam " &
        '               " left join materialmaster mm on mm.cmmf = bf.cmmf " &
        '               " left join range r on r.range = bf.range " &
        '               " left join doc.sebplatform sf on sf.cmmf = bf.cmmf " &
        '               " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
        '               " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2 {1} order by bf.sbu)", Criteria, Criteria.ToString.Replace("tu.year", "date_part('Year',bf.period)")))
        sb.Append(String.Format("(select distinct tu.fpcp,tu.sbu,f.familyname::text,tu.brand,tu.range,tu.cmmf,comfam , tu.cmmf || ' - ' || mm.materialdesc as cmmfdescription,tu.range::text || ' - ' || r.rangedesc as rangedescription,mm.commref, mm.commref || ' - ' || mm.materialdesc as commrefdescription,sf.cmmf as sebplatform,pr.paramname as status from doc.turnover tu " &
                               " left join vendor v on v.vendorcode = tu.vendorcode" &
                               " left join family f on f.familyid = tu.comfam" &
                               " left join materialmaster mm on mm.cmmf = tu.cmmf" &
                               " left join range r on r.range = tu.range" &
                               " left join doc.sebplatform sf on sf.cmmf = tu.cmmf" &
                               " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                               " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                               " {0}" &
                               " order by sbu)" &
                               " union all (select distinct bf.fpcp,bf.sbu,f.familyname::text,bf.brand,bf.range,bf.cmmf,comfam , bf.cmmf || ' - ' || mm.materialdesc as cmmfdescription,bf.range::text || ' - ' || r.rangedesc as rangedescription,mm.commref, mm.commref || ' - ' || mm.materialdesc as commrefdescription,sf.cmmf as sebplatform,pr.paramname as status   from doc.budgetforecast bf" &
                               " left join vendor v on v.vendorcode = bf.vendorcode " &
                               " left join family f on f.familyid = bf.comfam " &
                               " left join materialmaster mm on mm.cmmf = bf.cmmf " &
                               " left join range r on r.range = bf.range " &
                               " left join doc.sebplatform sf on sf.cmmf = bf.cmmf " &
                               " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode " &
                               " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2 {1} order by bf.sbu)", Criteria, Criteria.ToString.Replace("tu.year", "date_part('Year',bf.period)")))



        MasterDS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString, MasterDS, mymessage) Then
            Try
                MasterDS.Tables(0).TableName = "Master"

                ProgressReport(9, "Assign BindingSource")
            Catch ex As Exception
                ProgressReport(1, "UCSupplierDashboard Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
        End If
    End Sub

    Sub DoQuery2()
        sb.Clear()
        sb.Append(String.Format("select 'ALL' as sbu union all (select distinct sbu from doc.turnover tu " &
                                " left join vendor v on v.vendorcode = tu.vendorcode" &
                                " {0}" &
                                " order by sbu);", criteria))
        sb.Append(String.Format("select 'ALL' as fpcp union all (select distinct fpcp from doc.turnover tu " &
                                " left join vendor v on v.vendorcode = tu.vendorcode" &
                                " {0}" &
                                " order by fpcp);", criteria))
        sb.Append(String.Format("select 'ALL' as familyname,0 as comfam union all (select distinct f.familyname::text,tu.comfam from doc.turnover tu " &
                                " left join vendor v on v.vendorcode = tu.vendorcode" &
                                " left join family f on f.familyid = tu.comfam" &
                                " {0}" &
                                " order by familyname);", Criteria))
        sb.Append(String.Format("select 'ALL' as brand union all (select distinct tu.brand from doc.turnover tu " &
                                " left join vendor v on v.vendorcode = tu.vendorcode" &
                                " {0}" &
                                " order by tu.brand);", Criteria))
        sb.Append(String.Format("select 'ALL' as range,'ALL' as rangedesc,'ALL' as description union all (select distinct tu.range,r.rangedesc,tu.range::text || ' - ' || r.rangedesc as rangedescription from doc.turnover tu " &
                                " left join vendor v on v.vendorcode = tu.vendorcode" &
                                " left join range r on r.range = tu.range" &
                                " {0}" &
                                " order by r.rangedesc);", Criteria))
        sb.Append(String.Format("select 0 as cmmf,'ALL' as materialdesc,'ALL' as description union all (select distinct tu.cmmf, mm.materialdesc,tu.cmmf::text || ' - ' || mm.materialdesc as cmmfdescription from doc.turnover tu " &
                                " left join vendor v on v.vendorcode = tu.vendorcode" &
                                " left join materialmaster mm on mm.cmmf = tu.cmmf" &
                                " {0}" &
                                " order by mm.materialdesc);", Criteria))
        sb.Append(String.Format("select 'ALL' as commref,'ALL' as description union all (select distinct mm.commref,mm.commref || ' - ' || mm.materialdesc as commrefdescription from doc.turnover tu " &
                                " left join vendor v on v.vendorcode = tu.vendorcode" &
                                " left join materialmaster mm on mm.cmmf = tu.cmmf" &
                                " {0}" &
                                " order by mm.commref);", Criteria))


        SBUDS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString, SBUDS, mymessage) Then
            Try
                SBUDS.Tables(0).TableName = "SBU"

                ProgressReport(6, "Assign Combobox")
            Catch ex As Exception
                ProgressReport(1, "UCSupplierDashboard Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
        End If
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        'ComboBox1.SelectedIndex = 0
        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub CheckedListBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox2.SelectedIndexChanged
        CheckedListBox_SelectedIndexChanged(sender, e)

        Dim sb As New StringBuilder
        Dim sb2 As New StringBuilder
        Dim sb3 As New StringBuilder 'ProductTypeList
        Dim mylist As New StringBuilder
        Dim dbNullSB As New StringBuilder
        Dim obj As CheckedListBox = DirectCast(sender, CheckedListBox)

        StatusBSFilter = ""
        If obj.Items.Count > 0 Then
            Dim chkstate = obj.GetItemCheckState(0)
            If chkstate = CheckState.Unchecked Then
                sb.Append(" and (status  in (")

                'sb2.Append(" ProductType Filter: ")
                'sb3.Append(" and (fpcp in (")
                For i = 1 To obj.Items.Count - 1
                    chkstate = obj.GetItemCheckState(i)
                    If chkstate Then
                        If mylist.Length > 0 Then
                            sb.Append(",")
                            sb2.Append(",")
                            sb3.Append(",")
                        End If
                        mylist.Append("list")
                        If IsDBNull(obj.Items(i).row.item("status")) Then
                            dbNullSB.Append(" or status  isnull")
                            sb3.Append(String.Format("{0}", "NULL"))
                        Else
                            sb2.Append(String.Format("'{0}'", obj.Items(i).row.item("status")))
                            sb3.Append(String.Format("{0}", obj.Items(i).row.item("status")))
                        End If
                        sb.Append(String.Format("''{0}''", obj.Items(i).row.item("status")))
                    End If
                Next
                sb.Append(String.Format(") {0}) ", dbNullSB.ToString))
                'sb3.Append(String.Format(") {0}) ", dbNullSB.ToString))
                If sb3.Length = 0 Then
                    StatusCriteria = ""
                    sb.Clear()
                    'ProgressReport(8, "Please select from the list.")                    
                Else
                    StatusCriteria = sb.ToString
                    'ProgressReport(7, "")
                End If
                statusFilter = sb3.ToString
                StatusBSFilter = String.Format("{0}{1}", IIf(sb2.Length > 0, String.Format("status  in ({0})", sb2.ToString), ""), IIf(dbNullSB.Length > 0, IIf(sb2.Length > 0, " or ", "") & " status  is null", ""))
            Else
                statusFilter = "ALL"
                StatusCriteria = ""
            End If

        End If
        'Raise Change Shortname send statusCriteria
        If QueryType = QueryTypeEnum.shortname Then
            RaiseEvent StatusVendorChanged(StatusBSFilter)
        End If

        ApplyBSFilter()
    End Sub
    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        CheckedListBox_SelectedIndexChanged(sender, e)

        Dim sb As New StringBuilder
        Dim sb2 As New StringBuilder
        Dim sb3 As New StringBuilder 'ProductTypeList
        Dim mylist As New StringBuilder
        Dim dbNullSB As New StringBuilder
        ProductBSFilter = ""
        If CheckedListBox1.Items.Count > 0 Then
            Dim chkstate = CheckedListBox1.GetItemCheckState(0)
            If chkstate = CheckState.Unchecked Then
                sb.Append(" and (fpcp in (")

                'sb2.Append(" ProductType Filter: ")
                'sb3.Append(" and (fpcp in (")
                For i = 1 To CheckedListBox1.Items.Count - 1
                    chkstate = CheckedListBox1.GetItemCheckState(i)
                    If chkstate Then
                        If mylist.Length > 0 Then
                            sb.Append(",")
                            sb2.Append(",")
                            sb3.Append(",")
                        End If
                        mylist.Append("list")
                        If IsDBNull(CheckedListBox1.Items(i).row.item("fpcp")) Then
                            dbNullSB.Append(" or fpcp isnull")
                            sb3.Append(String.Format("{0}", "NULL"))
                        Else
                            sb2.Append(String.Format("'{0}'", CheckedListBox1.Items(i).row.item("fpcp")))
                            sb3.Append(String.Format("{0}", CheckedListBox1.Items(i).row.item("fpcp")))
                        End If
                        sb.Append(String.Format("''{0}''", IIf(IsDBNull(CheckedListBox1.Items(i).row.item("fpcp")), "NULL", CheckedListBox1.Items(i).row.item("fpcp"))))
                    End If
                Next
                sb.Append(String.Format(") {0}) ", dbNullSB.ToString))
                'sb3.Append(String.Format(") {0}) ", dbNullSB.ToString))
                If sb3.Length = 0 Then
                    ProductTypeCriteria = ""
                    sb.Clear()
                    'ProgressReport(8, "Please select from the list.")                    
                Else
                    ProductTypeCriteria = sb.ToString
                    'ProgressReport(7, "")
                End If
                ProductTypeFilter = sb3.ToString
                ProductBSFilter = String.Format("{0}{1}", IIf(sb2.Length > 0, String.Format("fpcp in ({0})", sb2.ToString), ""), IIf(dbNullSB.Length > 0, IIf(sb2.Length > 0, " or ", "") & " fpcp is null", "")).Replace(",,", ",")
            Else
                ProductTypeFilter = "ALL"
                ProductTypeCriteria = ""
            End If
        End If

        ApplyBSFilter()
    End Sub

    Public Function getFilterCriteria() As String
        Dim sb As New StringBuilder
        Dim sb2 As New StringBuilder
        Dim sb3 As New StringBuilder 'ProductTypeList
        Dim mylist As New StringBuilder
        Dim dbNullSB As New StringBuilder

        'Product Type

        'If CheckedListBox1.Items.Count > 0 Then
        '    Dim chkstate = CheckedListBox1.GetItemCheckState(0)
        '    If chkstate = CheckState.Unchecked Then
        '        sb.Append(" and (fpcp in (")

        '        'sb2.Append(" ProductType Filter: ")
        '        'sb3.Append(" and (fpcp in (")
        '        For i = 1 To CheckedListBox1.Items.Count - 1
        '            chkstate = CheckedListBox1.GetItemCheckState(i)
        '            If chkstate Then
        '                If mylist.Length > 0 Then                            
        '                    sb.Append(",")
        '                    'sb2.Append(",")
        '                    sb3.Append(",")
        '                End If
        '                mylist.Append("list")
        '                If IsDBNull(CheckedListBox1.Items(i).row.item("fpcp")) Then
        '                    dbNullSB.Append(" or fpcp isnull")
        '                    'sb2.Append(String.Format("{0}", "NULL"))
        '                    sb3.Append(String.Format("{0}", "NULL"))
        '                Else                            
        '                    'sb2.Append(String.Format("{0}", CheckedListBox1.Items(i).row.item("fpcp")))
        '                    sb3.Append(String.Format("{0}", CheckedListBox1.Items(i).row.item("fpcp")))
        '                End If
        '                sb.Append(String.Format("''{0}''", CheckedListBox1.Items(i).row.item("fpcp")))
        '            End If
        '        Next
        '        sb.Append(String.Format(") {0}) ", dbNullSB.ToString))
        '        'sb3.Append(String.Format(") {0}) ", dbNullSB.ToString))
        '        If sb3.Length = 0 Then
        '            ProductTypeCriteria = ""
        '            sb.Clear()
        '            ProgressReport(8, "Please select from the list.")
        '            Return ""
        '        Else
        '            ProductTypeCriteria = sb.ToString
        '            ProgressReport(7, "")
        '        End If
        '        ProductTypeFilter = sb3.ToString
        '    Else
        '        ProductTypeFilter = "ALL"
        '    End If
        'End If

        If ProductTypeCriteria <> "" Then
            sb.Append(ProductTypeCriteria)

        End If

        If StatusCriteria <> "" Then
            sb.Append(StatusCriteria)
            If mylist.Length > 0 Then
                sb2.Append(",")
            End If
            sb2.Append(statusFilter)
        End If

        'SBU



        'If comboboxCriteria <> "" Then
        If SBUCriteria <> "" Then

            If mylist.Length > 0 Then
                sb2.Append(",")
            End If
            'sb.Append(comboboxCriteria)
            'sb2.Append(ComboBoxFilter)
            sb.Append(SBUCriteria)
            sb2.Append(SBUFilter)
        End If

        'End If

        'family
        If FamilyCriteria <> "" Then
            'Dim drv As DataRowView = BSFamilyHelper.Current
            'sb.Append(String.Format(" and comfam = {0}", drv.Item("comfam")))
            sb.Append(FamilyCriteria)
            If mylist.Length > 0 Then
                sb2.Append(",")
            End If
            If TextBox4.Text = "" Then
                sb2.Append(String.Format(" Family: {0}", "NULL"))
            Else
                sb2.Append(String.Format(" Family: {0}", TextBox4.Text))
            End If
            'Dim drv As DataRowView = BSFamilyHelper.Current
            ''sb.Append(String.Format(" and comfam = {0}", drv.Item("comfam")))
            'sb.Append(FamilyCriteria)
            'If mylist.Length > 0 Then
            '    sb2.Append(",")
            'End If
            'If IsDBNull(drv.Item("comfam")) Then
            '    sb2.Append(String.Format(" Family: {0}", "NULL"))
            'Else
            '    sb2.Append(String.Format(" Family: {0}", drv.Item("familyname")))
            'End If

        End If

        'brand()
        If BrandCriteria <> "" Then
            'Dim drv As DataRowView = BSBrandHelper.Current
            sb.Append(BrandCriteria)
            If mylist.Length > 0 Then
                sb2.Append(",")
            End If
            If TextBox3.Text = "" Then
                sb2.Append(String.Format(" Brand: {0}", "NULL"))
            Else
                sb2.Append(String.Format(" Brand: {0}", TextBox3.Text))
            End If
            'Dim drv As DataRowView = BSBrandHelper.Current
            'sb.Append(BrandCriteria)
            'If mylist.Length > 0 Then
            '    sb2.Append(",")
            'End If
            'If IsDBNull(drv.Item("brand")) Then
            '    sb2.Append(String.Format(" Brand: {0}", "NULL"))
            'Else
            '    sb2.Append(String.Format(" Brand: {0}", drv.Item("Brand")))
            'End If


        End If

        'range code 6
        If RangeCodeCriteria <> "" Then
            sb.Append(RangeCodeCriteria)
            'Dim drv As DataRowView = BSRangeHelper.Current
            If mylist.Length > 0 Then
                sb2.Append(",")
            End If
            'sb2.Append(String.Format(" Range: {0}", drv.Item("description")))
            sb2.Append(String.Format(" Range: {0}", TextBox6.Text))
        End If

        'cmmf 5
        If CMMFCriteria <> "" Then            
            'Dim drv As DataRowView = BSCMMFHelper.Current
            sb.Append(CMMFCriteria)
            If mylist.Length > 0 Then
                sb2.Append(",")
            End If
            'sb2.Append(String.Format(" CMMF: {0}", drv.Item("description")))
            sb2.Append(String.Format(" CMMF: {0}", TextBox5.Text))
        End If
        'commercialcode 8
        If CommercialCodeCriteria <> "" Then
            sb.Append(CommercialCodeCriteria)
            'Dim drv As DataRowView = BSCommRefHelper.Current
            If mylist.Length > 0 Then
                sb2.Append(",")
            End If            
            'sb2.Append(String.Format(" Commercial Code: {0}", drv.Item("description")))
            sb2.Append(String.Format(" Commercial Code: {0}", TextBox8.Text))
        End If
        'odm/oem 7
        'SEB Platform
        If CheckBox1.Checked Then
            sb2.Append(String.Format(" SEB Asia Platform"))
            sb.Append(" and (not sf.cmmf isnull)")
        End If

        StringFilter = sb2.ToString

        Return sb.ToString
    End Function

    Public Function GetStatusCriteria() As String
        Return StatusCriteria.Replace("''", "'")

    End Function
    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs)

        ''Change DropDown into DropDownList
        'comboboxCriteria = ""
        'ComboBoxFilter = ""
        'If IsDBNull(ComboBox1.SelectedValue) Then
        '    comboboxCriteria = " and tu.sbu isnull"
        '    ComboBoxFilter = "SBU : NULL"
        'Else
        '    If ComboBox1.SelectedValue <> "ALL" Then

        '        comboboxCriteria = String.Format(" and tu.sbu = ''{0}''", ComboBox1.SelectedValue)
        '        ComboBoxFilter = String.Format(" SBU : {0}", ComboBox1.SelectedValue)
        '    End If

        'End If


    End Sub


    
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        doResetFilter()
    End Sub

    Private Sub doResetFilter()
        TextBox3.Text = "ALL"
        TextBox4.Text = "ALL"
        TextBox5.Text = "ALL"
        TextBox6.Text = "ALL"
        TextBox8.Text = "ALL"
        TextBox9.Text = "ALL"
        SBUCriteria = ""
        FamilyCriteria = ""
        BrandCriteria = ""
        comboboxCriteria = ""
        'ComboBox1.SelectedIndex = 0
        RangeCodeCriteria = ""
        CMMFCriteria = ""
        CommercialCodeCriteria = ""
        SBUBSFilter = ""
        FamilyBSFilter = ""
        BrandBSFilter = ""
        RangeCodeBSFilter = ""
        CheckBox1.Checked = False
        ApplyBSFilter()

    End Sub

    Private Sub ApplyBSFilter()

        If Not IsNothing(MasterDS) Then
            Dim mysb As New StringBuilder
            mysb.Append(ProductBSFilter)
            'If mysb.Length > 0 Then
            '    If StatusBSFilter.Length > 0 Then
            '        mysb.Append(" and ")
            '        mysb.Append(StatusBSFilter)
            '    End If

            'End If
            If mysb.Length > 0 Then
                If StatusBSFilter.Length > 0 Then
                    mysb.Append(" and ")
                    mysb.Append(StatusBSFilter)
                End If
            Else
                If Not IsNothing(StatusBSFilter) Then
                    If StatusBSFilter.Length > 0 Then
                        mysb.Append(StatusBSFilter)
                    End If
                End If
               
            End If

            Dim MyView As DataView = MasterDS.Tables(0).AsDataView
            MyView.Sort = "sbu"
            MyView.RowFilter = mysb.ToString 'ProductBSFilter
            Dim SBUTable = MyView.ToTable(True, {"sbu"})

            BSSBUHelper = New BindingSource
            BSFamilyHelper = New BindingSource
            BSBrandHelper = New BindingSource
            BSRangeHelper = New BindingSource
            BSCMMFHelper = New BindingSource
            BSCommRefHelper = New BindingSource

            BSSBUHelper.DataSource = SBUTable
            Dim dr = SBUTable.NewRow
            dr.Item(0) = "ALL"
            SBUTable.Rows.InsertAt(dr, 0)
            BSSBUHelper.Position = 0

            MyView.Sort = "familyname"

            If mysb.Length > 0 Then
                If SBUBSFilter <> "" Then
                    mysb.Append(" and ")
                    mysb.Append(SBUBSFilter)
                End If
            Else
                mysb.Append(SBUBSFilter)
            End If
            MyView.RowFilter = mysb.ToString 'ProductBSFilter & SBUBSFilter
            Dim FamilyTable = MyView.ToTable(True, {"familyname", "comfam"})
            BSFamilyHelper.DataSource = FamilyTable
            dr = FamilyTable.NewRow
            dr.Item(0) = "ALL"
            dr.Item(1) = 0
            FamilyTable.Rows.InsertAt(dr, 0)
            BSFamilyHelper.Position = 0

            MyView.Sort = "brand"
            If mysb.Length > 0 Then
                If FamilyBSFilter <> "" Then
                    mysb.Append(" and ")
                    mysb.Append(FamilyBSFilter)
                End If
            Else
                mysb.Append(FamilyBSFilter)
            End If

            MyView.RowFilter = mysb.ToString
            Dim BrandTable = MyView.ToTable(True, {"brand"})
            BSBrandHelper.DataSource = BrandTable
            dr = BrandTable.NewRow
            dr.Item(0) = "ALL"
            BrandTable.Rows.InsertAt(dr, 0)
            BSBrandHelper.Position = 0

            MyView.Sort = "rangedescription"
            If mysb.Length > 0 Then
                If BrandBSFilter <> "" Then
                    mysb.Append(" and ")
                    mysb.Append(BrandBSFilter)
                End If
            Else
                mysb.Append(BrandBSFilter)
            End If

            MyView.RowFilter = mysb.ToString
            Dim RangeTable = MyView.ToTable(True, {"rangedescription", "range"})
            BSRangeHelper.DataSource = RangeTable
            dr = RangeTable.NewRow
            dr.Item(0) = "ALL"
            dr.Item(1) = ""
            RangeTable.Rows.InsertAt(dr, 0)
            BSRangeHelper.Position = 0


            MyView.Sort = "cmmfdescription"
            If mysb.Length > 0 Then
                If RangeCodeBSFilter <> "" Then
                    mysb.Append(" and ")
                    mysb.Append(RangeCodeBSFilter)
                End If
            Else
                mysb.Append(RangeCodeBSFilter)
            End If

            If CheckBox1.Checked Then
                If mysb.Length > 0 Then
                    mysb.Append(" and ")
                End If
                mysb.Append(" not (sebplatform is null)")

            End If

            MyView.RowFilter = mysb.ToString
            Dim CMMFTable = MyView.ToTable(True, {"cmmfdescription", "cmmf"})
            BSCMMFHelper.DataSource = CMMFTable
            dr = CMMFTable.NewRow
            dr.Item(0) = "ALL"
            dr.Item(1) = 0
            CMMFTable.Rows.InsertAt(dr, 0)
            BSCMMFHelper.Position = 0

            MyView.Sort = "commrefdescription"
            MyView.RowFilter = mysb.ToString
            Dim CommRefTable = MyView.ToTable(True, {"commrefdescription", "commref"})
            BSCommRefHelper.DataSource = CommRefTable
            dr = CommRefTable.NewRow
            dr.Item(0) = "ALL"
            dr.Item(1) = ""
            CommRefTable.Rows.InsertAt(dr, 0)
            BSCommRefHelper.Position = 0

        End If
    End Sub



    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        ApplyBSFilter()
    End Sub

    Sub DoQuery5()
        Dim criteria As String = String.Empty
        Select Case QueryType
            Case QueryTypeEnum.shortname
                criteria = String.Format(" shortname = '{0}' ", ShortName.Replace("'", "''"))
            Case QueryTypeEnum.vendorcode
                criteria = String.Format(" v.vendorcode = {0} ", Vendorcode)
        End Select
        Dim sb As New StringBuilder
        sb.Clear()
        'sb.Append(String.Format("with " &
        '                        " vc as (select distinct shortname from vendor v where {0})," &
        '                        " d as (select date_part('Year',validfrom) as year,ph.validfrom,v.shortname,v.vendorcode,sc.category::text,ps.paneldescription as fppanelstatus,ps2.paneldescription as cppanelstatus from doc.panelhistory ph" &
        '                        " left join vendor v on v.vendorcode = ph.vendorcode " &
        '                        " left join supplierscategory sc on sc.supplierscategoryid = ph.suppliercategoryid" &
        '                        " left join doc.panelstatus ps on ps.id = ph.fp" &
        '                        " left join doc.panelstatus ps2 on ps2.id = ph.cp" &
        '                         " inner join vc on vc.shortname = v.shortname " &
        '                        " order by validfrom desc)" &
        '                        " select distinct year, first_value(category) over (partition by year order by validfrom desc) as category," &
        '                        " first_value(fppanelstatus) over (partition by year order by validfrom desc) as fppanelstatus," &
        '                        " first_value(cppanelstatus) over (partition by year order by validfrom desc) as cppanelstatus" &
        '                        " from d " &
        '                        " where  year <= {1} " &
        '                        " order by year desc limit 5;", criteria, Year(currentDate)))
        sb.Append(String.Format("with " &
                                " s as (select generate_series({1}-4,{1}) as m order by m desc)," &
                                " vc as (select distinct shortname from vendor v where {0})," &
                                " d as (select date_part('Year',validfrom) as year,ph.validfrom,v.shortname,v.vendorcode,sc.category::text,ps.paneldescription as fppanelstatus,ps2.paneldescription as cppanelstatus from doc.panelhistory ph" &
                                " left join vendor v on v.vendorcode = ph.vendorcode " &
                                " left join supplierscategory sc on sc.supplierscategoryid = ph.suppliercategoryid" &
                                " left join doc.panelstatus ps on ps.id = ph.fp" &
                                " left join doc.panelstatus ps2 on ps2.id = ph.cp" &
                                " inner join vc on vc.shortname = v.shortname " &
                                " order by validfrom desc)," &
                                " k as (select distinct year, first_value(category) over (partition by year order by validfrom desc) as category," &
                                " first_value(fppanelstatus) over (partition by year order by validfrom desc) as fppanelstatus," &
                                " first_value(cppanelstatus) over (partition by year order by validfrom desc) as cppanelstatus" &
                                " from d " &
                                " order by year desc)" &
                                " select distinct s.m as year,first_value(k.category) over (partition by m order by year desc ) as category," &
                                " first_value(k.fppanelstatus) over (partition by m order by year desc) as fppanelstatus," &
                                " first_value(k.cppanelstatus) over (partition by m order by year desc) as cppanelstatus from s" &
                                " left join k on k.year <= s.m" &
                                " order by year desc" &
                                ";", criteria, Year(currentDate)))
        '" where not category isnull" &
        sb.Append(String.Format("with " &
                                " vc as (select distinct shortname from vendor v where {0})" &
                                " select distinct case lineno when 1 then 'CP Main Technology' else 'CP Technology' end as technology, t.technologyname ,lineno" &
                                " from doc.vendortechnology vt " &
                                " left join doc.technology t on t.id = vt.technologyid" &
                                " left join vendor v on v.vendorcode = vt.vendorcode" &
                                " inner join vc on vc.shortname = v.shortname " &
                                " order by lineno;", criteria))

        sb.Append(String.Format("with " &
                                " vc as (select distinct shortname from vendor v where {0})" &
                  "select v.vendorcode,v.vendorname::character varying,ph.validfrom::date,sc.category::text,ps.paneldescription as fppanelstatus,ps2.paneldescription as cppanelstatus from doc.panelhistory ph" &
                  " left join vendor v on v.vendorcode = ph.vendorcode" &
                  " left join supplierscategory sc on sc.supplierscategoryid = ph.suppliercategoryid" &
                  " left join doc.panelstatus ps on ps.id = ph.fp" &
                  " left join doc.panelstatus ps2 on ps2.id = ph.cp" &
                  " inner join vc on vc.shortname = v.shortname " &
                  " order by validfrom desc;", criteria))
        Dim myDS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString, myDS, mymessage) Then
            Try
                'myDS.Tables(0).TableName = "PaneStatusSupplier"
                PanelStatus = New BindingSource
                SupplierTechnologyBS = New BindingSource
                SupplierPanelHistoryBS = New BindingSource
                PanelStatus.DataSource = myDS.Tables(0)
                SupplierTechnologyBS.DataSource = myDS.Tables(1)
                SupplierPanelHistoryBS.DataSource = myDS.Tables(2)
                RaiseEvent FinishQuery5()
                ProgressReport(12, "Refresh DataGridView")
            Catch ex As Exception
                ProgressReport(1, "UCSupplierDashboard Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
        End If

    End Sub





End Class
