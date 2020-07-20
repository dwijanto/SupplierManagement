Imports System.Threading
Imports SupplierManagement.SharedClass
Imports SupplierManagement.PublicClass
Imports System.Text
Imports Microsoft.Office.Interop

Public Class FormFactoryAndContact
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myQueryThread As New System.Threading.Thread(AddressOf DoQuery)
    Public BSshortNameHelper As BindingSource       'Source coming from ClassSupplierDashBoard
    Public BSvendorhelper As BindingSource          'Source coming from ClassSupplierDashBoard



    Private vendorcode As Long
    Private shortname As String
    Private sqlstr As String
    Private sb As StringBuilder
    Private DS As DataSet
    Private MyDS As DataSet
    Private ParamDS As DataSet
    Private myCriteria As String
    Private myCriteria2 As String
    Private VBS As BindingSource 'Vendor BS
    Private VBSHelper As BindingSource 'Vendor Helper BS

    Private VFBS As BindingSource 'Vendor Factory BS
    Private VCBS As BindingSource 'Vendor Contact BS

    Private FDTBS As BindingSource 'Factory Detail BS
    Private FHDBS As BindingSource 'Factory Header BS
    Private CBS As BindingSource 'Contact BS

    Private ProvinceBS As BindingSource
    Private CountryBS As BindingSource
    Dim ClassHelper = New SupplierManagement.ClassSupplierDashBoard(Me) 'This class for fill Vendor and Shortname

    Dim myFKey As ForeignKeyConstraint
    Dim sqlstrReport As String

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        initQuery()
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub initQuery()
        If Not myQueryThread.IsAlive Then

            ToolStripStatusLabel1.Text = ""
            myQueryThread = New Thread(AddressOf DoQuery)
            myQueryThread.Start()
        Else
            MessageBox.Show(String.Format("{0} Please wait until the current process is finished.", System.Reflection.MethodInfo.GetCurrentMethod()))
        End If
    End Sub
    Private Sub InitData()
        If Not myThread.IsAlive Then

            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show(String.Format("{0} Please wait until the current process is finished.", System.Reflection.MethodInfo.GetCurrentMethod()))
        End If
    End Sub

    Sub DoQuery()
        sb = New StringBuilder
        sb.Append("select pd.paramdtid as id,pd.paramname as province from doc.paramdt pd" &
                  " left join doc.paramhd ph on ph.paramhdid = pd.paramhdid" &
                  " where ph.paramname = 'province'" &
                  " order by pd.paramname;")
        sb.Append("select pd.paramdtid as id,pd.paramname as country from doc.paramdt pd" &
                  " left join doc.paramhd ph on ph.paramhdid = pd.paramhdid" &
                  " where ph.paramname = 'country'" &
                  " order by pd.paramname;")

        Dim mymessage As String = String.Empty
        ParamDS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString, ParamDS, mymessage) Then
            Try
                ParamDS.Tables(0).TableName = "Province"
                ParamDS.Tables(1).TableName = "Country"
            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(8, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub
    Sub DoWork()
        sb = New StringBuilder
        'sb.Append(String.Format(" select v.vendorcode,v.vendorname::text,v.shortname::text,ssm.officersebname::text as ssm,v.ssmidpl as ssmid,v.ssmeffectivedate,pm.officersebname::text as pm,v.pmid as pmid,v.pmeffectivedate,o.officername::text,v.shortname2,true as status" &
        '          " from vendor v" &
        '          " left join officerseb ssm on ssm.ofsebid = v.ssmidpl" &
        '          " left join officerseb pm on pm.ofsebid = v.pmid" &
        '          " left join officer o on o.officerid = v.officerid {0}" &
        '          " order by vendorcode;", myCriteria))
        'sb.Append(String.Format(" select v.vendorcode,v.vendorname::text,v.shortname::text,mu2.username as ssm,mu.username as pm,v.shortname2,true as status" &
        '         " from vendor v" &
        '         " LEFT JOIN doc.viewvendorfamilypm vfp ON vfp.vendorcode = v.vendorcode" &
        '         " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid" &
        '         " LEFT JOIN masteruser mu ON mu.id = os.muid" &
        '         " LEFT JOIN officerseb o ON o.ofsebid = os.parent" &
        '         " LEFT JOIN masteruser mu2 ON mu2.id = o.muid" &
        '         " LEFT JOIN doc.vendorgsm gsm ON gsm.vendorcode = v.vendorcode" &
        '         " LEFT JOIN officerseb o1 ON o1.ofsebid = gsm.gsmid" &
        '         " LEFT JOIN masteruser mu1 ON mu1.id = o1.muid {0}" &
        '         " order by vendorcode;", myCriteria))
        'viewvendorpmeffectivedate replace viewvendorfamilypmeffectivedate
        sb.Append(String.Format(" select v.vendorcode,v.vendorname::text,v.shortname::text,mu2.username as ssm,vfp.spmeffectivedate,mu1.username as gsm,gsm.effectivedate as gsmeffectivedate,mu.username as pm,vfp.pmeffectivedate,v.shortname2,true as status" &
                " from vendor v" &
                " LEFT JOIN doc.viewvendorpmeffectivedate vfp ON vfp.vendorcode = v.vendorcode" &
                " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid" &
                " LEFT JOIN masteruser mu ON mu.id = os.muid" &
                " LEFT JOIN officerseb o ON o.ofsebid = os.parent" &
                " LEFT JOIN masteruser mu2 ON mu2.id = o.muid" &
                " LEFT JOIN doc.vendorgsm gsm ON gsm.vendorcode = v.vendorcode" &
                " LEFT JOIN officerseb o1 ON o1.ofsebid = gsm.gsmid" &
                " LEFT JOIN masteruser mu1 ON mu1.id = o1.muid {0}" &
                " order by vendorcode;", myCriteria))


        Dim mymessage As String = String.Empty
        DS = New DataSet

        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                DS.Tables(0).TableName = "Vendor"
            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(4, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Loading Data.Done!")
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
                        VBS = New BindingSource
                        VFBS = New BindingSource
                        FDTBS = New BindingSource
                        FHDBS = New BindingSource

                        VBS.DataSource = DS.Tables(0)





                        DataGridView1.AutoGenerateColumns = False
                        DataGridView1.DataSource = VBS

                        VFBS.DataSource = Nothing
                        FDTBS.DataSource = Nothing
                        FHDBS.DataSource = Nothing
                        DataGridView2.DataSource = FDTBS

                        Datagridview1DoubleClick()


                    Case 8
                        ProvinceBS = New BindingSource
                        CountryBS = New BindingSource
                        ProvinceBS.DataSource = ParamDS.Tables(0)
                        CountryBS.DataSource = ParamDS.Tables(1)


                    Case 7



                        FDTBS = New BindingSource
                        FHDBS = New BindingSource
                        VCBS = New BindingSource
                        CBS = New BindingSource
                        'VBSHelper = New BindingSource

                        TextBox19.Text = ""
                        TextBox19.DataBindings.Clear()

                        Dim pk1(0) As DataColumn
                        pk1(0) = MyDS.Tables(1).Columns("id")
                        MyDS.Tables(1).PrimaryKey = pk1
                        MyDS.Tables(1).Columns("id").AutoIncrement = True
                        MyDS.Tables(1).Columns("id").AutoIncrementSeed = 0
                        MyDS.Tables(1).Columns("id").AutoIncrementStep = -1
                        MyDS.Tables(1).TableName = "Factory"


                        Dim pk2(0) As DataColumn
                        pk2(0) = MyDS.Tables(2).Columns("id")
                        MyDS.Tables(2).PrimaryKey = pk2
                        MyDS.Tables(2).Columns("id").AutoIncrement = True
                        MyDS.Tables(2).Columns("id").AutoIncrementSeed = 0
                        MyDS.Tables(2).Columns("id").AutoIncrementStep = -1
                        MyDS.Tables(2).TableName = "FactoryHD"

                        Dim pk4(0) As DataColumn
                        pk4(0) = MyDS.Tables(4).Columns("id")
                        MyDS.Tables(4).PrimaryKey = pk4
                        MyDS.Tables(4).Columns("id").AutoIncrement = True
                        MyDS.Tables(4).Columns("id").AutoIncrementSeed = 0
                        MyDS.Tables(4).Columns("id").AutoIncrementStep = -1
                        MyDS.Tables(4).TableName = "Contact"

                        VFBS.DataSource = MyDS.Tables(0)
                        FDTBS.DataSource = MyDS.Tables(1)
                        FHDBS.DataSource = MyDS.Tables(2)
                        VCBS.DataSource = MyDS.Tables(3)
                        CBS.DataSource = MyDS.Tables(4)




                        Dim myrel As DataRelation

                        Dim FactoryCol As DataColumn
                        Dim VendorFactory As DataColumn
                        Dim ContactCol As DataColumn
                        Dim VendorContact As DataColumn

                        VendorFactory = MyDS.Tables(0).Columns("factoryid")
                        FactoryCol = MyDS.Tables(1).Columns("id")
                        myFKey = New ForeignKeyConstraint("VendorFactory", FactoryCol, VendorFactory)
                        myFKey.UpdateRule = Rule.Cascade
                        myFKey.DeleteRule = Rule.Cascade
                        MyDS.Tables(0).Constraints.Add(myFKey)


                        VendorContact = MyDS.Tables(3).Columns("contactid")
                        ContactCol = MyDS.Tables(4).Columns("id")
                        myFKey = New ForeignKeyConstraint("VendorContact", ContactCol, VendorContact)
                        myFKey.UpdateRule = Rule.Cascade
                        myFKey.DeleteRule = Rule.Cascade
                        MyDS.Tables(3).Constraints.Add(myFKey)


                        Dim FactoryHD As DataColumn
                        Dim FactoryDT As DataColumn
                        FactoryHD = MyDS.Tables(2).Columns("id")
                        FactoryDT = MyDS.Tables(1).Columns("factoryhdid")
                        myrel = New DataRelation("FactoryRel", FactoryHD, FactoryDT)
                        MyDS.Relations.Add(myrel)

                        DataGridView2.AutoGenerateColumns = False
                        DataGridView2.DataSource = FDTBS
                        DirectCast(DataGridView2.Columns("ColumnProvince"), DataGridViewComboBoxColumn).DataSource = ProvinceBS
                        DirectCast(DataGridView2.Columns("ColumnProvince"), DataGridViewComboBoxColumn).DisplayMember = "province"
                        DirectCast(DataGridView2.Columns("ColumnProvince"), DataGridViewComboBoxColumn).ValueMember = "id"

                        DirectCast(DataGridView2.Columns("ColumnCountry"), DataGridViewComboBoxColumn).DataSource = CountryBS
                        DirectCast(DataGridView2.Columns("ColumnCountry"), DataGridViewComboBoxColumn).DisplayMember = "country"
                        DirectCast(DataGridView2.Columns("ColumnCountry"), DataGridViewComboBoxColumn).ValueMember = "id"

                        'TextBox19.DataBindings.Clear()
                        TextBox19.DataBindings.Add(New Binding("text", FHDBS, "customname", True, DataSourceUpdateMode.OnPropertyChanged))

                        DataGridView3.AutoGenerateColumns = False
                        DataGridView3.DataSource = CBS

                        Button5.Enabled = True
                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If

    End Sub
    Public Sub clearDisplay()
        myCriteria = ""
        TextBox19.DataBindings.Clear()
        TextBox19.Text = ""
        DataGridView1.DataSource = Nothing
        DataGridView3.DataSource = Nothing
        DataGridView2.DataSource = Nothing
        DataGridView3.Invalidate()
        DataGridView1.Invalidate()
        DataGridView2.Invalidate()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button2.Click
        Dim obj As Button = DirectCast(sender, Button)



        'Check Changes
        If Not CheckChanges() Then
            Exit Sub
        End If

        Select Case obj.Name
            Case "Button1"
                Dim myform = New FormHelper(BSshortNameHelper)
                myform.DataGridView1.Columns(0).DataPropertyName = "shortname"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = BSshortNameHelper.Current
                    TextBox1.Text = "" + drv.Item("shortname")
                    TextBox2.Text = ""
                    myCriteria = String.Format("where v.shortname = '{0}'", TextBox1.Text)
                    If TextBox1.Text <> "" Then
                        RefreshDataGrid()
                    Else
                        clearDisplay()
                    End If
                End If
            Case "Button2"
                Dim myform = New FormHelper(BSvendorhelper)
                myform.DataGridView1.Columns(0).DataPropertyName = "description"
                If myform.ShowDialog = DialogResult.OK Then
                    Dim drv As DataRowView = BSvendorhelper.Current

                    If Not IsDBNull(drv.Item("vendorcode")) Then
                        vendorcode = drv.Item("vendorcode")
                        shortname = Nothing
                        Dim VendorName = drv.Item("vendorname")
                        TextBox2.Text = VendorName
                        TextBox1.Text = ""
                        myCriteria = String.Format("where v.vendorcode = {0}", vendorcode)
                        If TextBox2.Text <> "" Then
                            RefreshDataGrid()
                        End If
                    Else
                        TextBox2.Text = ""
                        clearDisplay()

                    End If
                End If

        End Select

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

    End Sub

    Private Sub RefreshDataGrid()
        InitData()

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'MessageBox.Show("Saved")
        doSave()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        DataGridView1.Enabled = False
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        'Datagridview1DoubleClick()

    End Sub

    Private Sub loaddata()
        sb = New StringBuilder
        'sb.Append(String.Format("select vf.*,v.shortname::text,v.vendorname from doc.vendorfactory vf" &
        '                        " left join vendor v on v.vendorcode = vf.vendorcode {0};", myCriteria2))
        sb.Append(String.Format("select distinct vf.*,v.shortname::text,v.vendorname from doc.vendorfactory vf" &
                                " left join vendor v on v.vendorcode = vf.vendorcode {0};", myCriteria))
        sb.Append(String.Format("with h as (select  id,max(stamp) as modifieddate from doc.factorydtl_audit group by id order by id ) select distinct f.*,h.modifieddate from doc.factorydtl f" &
                                " left join h on h.id = f.id" &
                                " left join doc.vendorfactory vf on vf.factoryid = f.id" &
                                " left join vendor v on v.vendorcode = vf.vendorcode {0} order by f.id;", myCriteria))
        sb.Append(String.Format("select distinct fh.* from doc.factoryhd fh" &
                                " left join doc.factorydtl f on f.factoryhdid = fh.id" &
                                " left join doc.vendorfactory vf on vf.factoryid = f.id" &
                                " left join vendor v on v.vendorcode = vf.vendorcode {0}; ", myCriteria))
        sb.Append(String.Format("select vc.*,v.shortname::text,v.vendorname from doc.vendorcontact vc" &
                                " left join vendor v on v.vendorcode = vc.vendorcode {0};", myCriteria))
        sb.Append(String.Format("with h as (select  id,max(stamp) as modifieddate from doc.contact_audit group by id order by id ) select distinct c.*,h.modifieddate from doc.contact c" &
                                " left join h on h.id = c.id" &
                                " left join doc.vendorcontact vc on vc.contactid = c.id" &
                                " left join vendor v on v.vendorcode = vc.vendorcode {0};", myCriteria))
        'sb.Append(String.Format("select distinct vf.*,v.shortname::text,v.vendorname from doc.vendorfactory vf" &
        '                        " left join vendor v on v.vendorcode = vf.vendorcode {0};", myCriteria))


        Dim mymessage As String = String.Empty
        MyDS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString, MyDS, mymessage) Then
            Try
                DS.Tables(0).TableName = "Vendor"

            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(7, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub
    Private Sub AddFactoryToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddFactoryToolStripMenuItem1.Click
        'Add Header
        'Add Detail
        If Not IsNothing(VBS) Then


            Dim drv As DataRowView
            If IsNothing(FHDBS.Current) Then
                drv = FHDBS.AddNew
                drv.Item("customname") = "Custom Name"
                drv.EndEdit()
            Else

            End If
            Dim drvdt As DataRowView = FDTBS.AddNew()
            drv = FHDBS.Current
            drvdt.Item("factoryhdid") = drv.Item("id")
            drvdt.Item("chinesename") = ""
            drvdt.Item("englishname") = ""
            drvdt.Item("englishaddress") = ""
            drvdt.Item("main") = False
            drvdt.EndEdit()

            Dim myform = New FormDialogFactory(FDTBS, ProvinceBS, CountryBS)
            myform.VFBS = VFBS
            myform.VBS = VBS
            If Not myform.ShowDialog = DialogResult.OK Then
                FDTBS.RemoveCurrent()
            Else
                DataGridView2.Invalidate()
            End If
        End If
    End Sub

    Private Sub AddFactoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddContactToolStripMenuItem.Click
        If Not IsNothing(VBS) Then


            'Add Contact       
            Dim drvdt As DataRowView = CBS.AddNew()
            drvdt.Item("contactname") = ""
            drvdt.Item("title") = ""
            drvdt.Item("email") = ""
            drvdt.Item("isecoqualitycontact") = False
            drvdt.EndEdit()

            Dim myform = New FormDialogContact(CBS)
            myform.VCBS = VCBS
            myform.VBS = VBS
            If Not myform.ShowDialog = DialogResult.OK Then
                CBS.RemoveCurrent()
            Else
                DataGridView3.Invalidate()
            End If
        End If
    End Sub

    Private Sub DataGridView2_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView2.CellBeginEdit
        If e.ColumnIndex = 1 Then
            Dim obj = DirectCast(sender, DataGridView)
            'obj.Invalidate()
            'If Not IsNothing(obj.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
            For i = 0 To obj.Rows.Count - 1
                obj.Rows(i).Cells(e.ColumnIndex).Value = False
            Next
            'obj.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = True
        
        End If
    End Sub

    Private Sub Datagridview1DoubleClick()
        If Not IsNothing(VBS.Current) Then
            Dim drv As DataRowView = VBS.Current
            myCriteria2 = String.Format("where v.vendorcode = {0}", drv.Row.Item("vendorcode"))
        End If
        If Not CheckChanges() Then
            Exit Sub
        End If
        loaddata()
    End Sub

    Private Function CheckChanges() As Boolean
        Dim hasChanges As Boolean = True
        If Not IsNothing(MyDS) Then
            Dim ds2 As DataSet = MyDS.GetChanges
            If Not IsNothing(ds2) Then
                Select Case MessageBox.Show(String.Format("There is unsaved data in a row.{0}Do you want to store to the database?", vbCrLf), "", MessageBoxButtons.YesNoCancel)
                    Case Windows.Forms.DialogResult.Yes
                        doSave()
                    Case Windows.Forms.DialogResult.Cancel
                        hasChanges = False
                    Case Windows.Forms.DialogResult.No
                        MyDS.AcceptChanges()
                End Select
            End If

        End If
        Return hasChanges
    End Function
    Private Sub DataGridView2_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.ColumnIndex = 1 Then
            If MessageBox.Show("Use this Factory as custom name?", "Custom Name", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.OK Then
                Try
                    Dim drv As DataRowView = FDTBS.Current
                    Dim hasError As Boolean = False
                    If IsDBNull(drv.Row.Item("chinesename")) Then
                        hasError = True
                    ElseIf drv.Row.Item("chinesename") = "" Then
                        hasError = True
                    End If
                    If hasError Then
                        MessageBox.Show("Chinese Name is blank.")

                    Else
                        Dim mydrv As DataRowView = FHDBS.Current
                        If IsNothing(mydrv) Then
                            Err.Raise(500, Description:="Change selection supplier using Short Name.")
                        Else
                            mydrv.Row.Item("customname") = drv.Row.Item("chinesename")
                        End If

                        'TextBox19.Text = drv.Row.Item("chinesename")


                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try


            End If
        End If
    End Sub

    Private Sub DataGridView2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.CellDoubleClick
        If Not IsNothing(FDTBS.Current) Then
            Dim drv As DataRowView = FDTBS.Current
            VFBS.Filter = String.Format("factoryid = {0}", drv.Row.Item("id"))
            'VFBS.Filter = ""

            Dim myform = New FormDialogFactory(FDTBS, ProvinceBS, CountryBS)
            myform.MyAction = FormDialogFactory.TxAction.Update_Record
            myform.VFBS = VFBS
            myform.VBS = VBS
            myform.ShowDialog()

            DataGridView2.Invalidate()
        End If

    End Sub


    Private Sub DataGridView3_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellDoubleClick
        If Not IsNothing(CBS.Current) Then
            Dim myform = New FormDialogContact(CBS)
            myform.VCBS = VCBS
            myform.VBS = VBS
            myform.MyAction = FormDialogContact.TxAction.Update_Record
            myform.ShowDialog()
            DataGridView3.Invalidate()
        End If
    End Sub

    Private Sub DeleteToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem1.Click
        'Show Popup Relation Factory and Vendor : vendorfactory
        If Not IsNothing(VBS) Then
            Dim drv As DataRowView = FDTBS.Current
            VFBS.Filter = String.Format("factoryid = {0}", drv.Row.Item("id"))
            Dim myform = New FormDialogDeleteFactoryContact
            myform.BS = VFBS
            myform.MstBS = FDTBS
            myform.DataBinding()
            myform.ShowDialog()
            VFBS.Filter = ""
        End If

    End Sub

    Private Sub doSave()
        If Me.Validate Then
            Try
                FDTBS.EndEdit()
                FHDBS.EndEdit()
                VFBS.EndEdit()
                VCBS.EndEdit()

                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                Dim ds2 As DataSet
                ds2 = MyDS.GetChanges

                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                    If Not DbAdapter1.VendorFactoryContact(Me, mye) Then
                        MessageBox.Show(mye.message)
                        Exit Sub
                    End If
                    'MyDS.Merge(ds2)

                    MyDS.AcceptChanges()
                    'DataGridView1.Invalidate()
                    MessageBox.Show("Saved.")
                End If
            Catch ex As Exception
                MessageBox.Show(" Error:: " & ex.Message)
            End Try
        End If
        loaddata()
        DataGridView1.Invalidate()

    End Sub


    Private Sub DataGridView3_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        'Show Popup Relation Factory and Vendor : vendorfactory
        If Not IsNothing(VBS) Then


            Dim drv As DataRowView = CBS.Current
            VCBS.Filter = String.Format("contactid = {0}", drv.Row.Item("id"))
            Dim myform = New FormDialogDeleteFactoryContact
            myform.BS = VCBS
            myform.MstBS = CBS
            myform.DataBinding()
            myform.ShowDialog()
            VCBS.Filter = ""
        End If
    End Sub


    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim sb As New StringBuilder
        ToolStripStatusLabel1.Text = ""
        ToolStripStatusLabel2.Text = ""
        If TextBox1.Text <> "" Then
            sb.Append(String.Format(" and v.shortname = '{0}'", TextBox1.Text))
            'sb.Append(String.Format(" where v.shortname = '{0}'", TextBox1.Text))
        ElseIf TextBox2.Text <> "" Then
            sb.Append(String.Format(" and v.vendorname = '{0}'", TextBox2.Text.Replace("'", "''")))
            'sb.Append(String.Format(" where v.vendorname = '{0}'", TextBox2.Text.Replace("'", "''")))
        End If

        'sqlstrReport = String.Format("with tu as (select distinct vendorcode,fpcp from doc.turnover order by vendorcode)" &
        '                                           " select distinct v.vendorcode,vendorname::text,v.shortname::text,spm.officersebname::text as spm,pm.officersebname::text as pm,tu.fpcp,pr.paramname as status, fhd.customname,fdt.chinesename,fdt.englishname,fdt.englishaddress,fdt.chineseaddress,fdt.area,fdt.city,prov.paramname as province,cty.paramname as country,c.contactname,c.title,c.email,c.officeph,c.factoryph,c.officemb,c.factorymb from vendor v " &
        '                                           " left join doc.vendorfactory vf on vf.vendorcode = v.vendorcode" &
        '                                           " left join doc.factorydtl fdt on fdt.id = vf.factoryid" &
        '                                           " left join doc.factoryhd fhd on fhd.id = fdt.factoryhdid" &
        '                                           " left join doc.vendorcontact vc on vc.vendorcode = v.vendorcode" &
        '                                           " left join doc.contact c on c.id = vc.contactid" &
        '                                           " left join doc.paramdt prov on prov.paramdtid = fdt.provinceid" &
        '                                           " left join doc.paramdt cty on cty.paramdtid = fdt.countryid" &
        '                                           " left join officerseb spm on spm.ofsebid = v.ssmidpl" &
        '                                           " left join officerseb pm on pm.ofsebid = v.pmid" &
        '                                           " left join tu on tu.vendorcode = v.vendorcode" &
        '                                           " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
        '                                           " left join doc.paramdt pr on pr.paramdtid = vs.status" &
        '                                           " where not vf.factoryid isnull and not vc.contactid isnull {0} order by v.vendorcode", sb.ToString)
        'sqlstrReport = String.Format("with tu as (select distinct vendorcode,fpcp from doc.turnover order by vendorcode)" &
        '                                           " select distinct v.vendorcode,vendorname::text,v.shortname::text,spm.officersebname::text as spm,pm.officersebname::text as pm,tu.fpcp,pr.paramname as status, fhd.customname,fdt.chinesename,fdt.englishname,fdt.englishaddress,fdt.chineseaddress,fdt.area,fdt.city,prov.paramname as province,cty.paramname as country,c.contactname,c.title,c.email,c.officeph,c.factoryph,c.officemb,c.factorymb from vendor v " &
        '                                           " left join doc.vendorfactory vf on vf.vendorcode = v.vendorcode" &
        '                                           " left join doc.factorydtl fdt on fdt.id = vf.factoryid" &
        '                                           " left join doc.factoryhd fhd on fhd.id = fdt.factoryhdid" &
        '                                           " left join doc.vendorcontact vc on vc.vendorcode = v.vendorcode" &
        '                                           " left join doc.contact c on c.id = vc.contactid" &
        '                                           " left join doc.paramdt prov on prov.paramdtid = fdt.provinceid" &
        '                                           " left join doc.paramdt cty on cty.paramdtid = fdt.countryid" &
        '                                           " left join officerseb spm on spm.ofsebid = v.ssmidpl" &
        '                                           " left join officerseb pm on pm.ofsebid = v.pmid" &
        '                                           " left join tu on tu.vendorcode = v.vendorcode" &
        '                                           " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
        '                                           " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
        '                                            " where  not (vf.factoryid isnull and  vc.contactid isnull) {0} order by v.vendorcode", sb.ToString)
        'viewvendorpmeffectivedate replace viewvendorfamilypmeffectivedate
        'sqlstrReport = String.Format("with tu as (select distinct vendorcode,fpcp from doc.turnover order by vendorcode)" &
        '                                           " select distinct v.vendorcode,vendorname::text,v.shortname::text," &
        '                                           " mu2.username as spm,vfp.spmeffectivedate, mu1.username as gsm,gsm.effectivedate as gsmeffectivedate,mu.username as pm,vfp.pmeffectivedate," &
        '                                           " tu.fpcp as producttype,pr.paramname as status, fhd.customname,fdt.chinesename,fdt.englishname,fdt.englishaddress,fdt.chineseaddress,fdt.area,fdt.city,fdt.main,fdt.modifiedby,doc.getstampfactorydtlaudit(fdt.id) as modifieddate,prov.paramname as province,cty.paramname as country,c.contactname,c.title,c.email,c.officeph,c.factoryph,c.officemb,c.factorymb,upper(c.isecoqualitycontact::text) as isecoqualitycontact from vendor v " &
        '                                           "LEFT JOIN doc.viewvendorpmeffectivedate vfp ON vfp.vendorcode = v.vendorcode " &
        '                                           " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid " &
        '                                           " LEFT JOIN masteruser mu ON mu.id = os.muid " &
        '                                           " LEFT JOIN officerseb o ON o.ofsebid = os.parent " &
        '                                           " LEFT JOIN masteruser mu2 ON mu2.id = o.muid " &
        '                                           " LEFT JOIN doc.vendorgsm gsm ON gsm.vendorcode = v.vendorcode " &
        '                                           " LEFT JOIN officerseb o1 ON o1.ofsebid = gsm.gsmid " &
        '                                           " LEFT JOIN masteruser mu1 ON mu1.id = o1.muid" &
        '                                           " left join doc.vendorfactory vf on vf.vendorcode = v.vendorcode" &
        '                                           " left join doc.factorydtl fdt on fdt.id = vf.factoryid" &
        '                                           " left join doc.factoryhd fhd on fhd.id = fdt.factoryhdid" &
        '                                           " left join doc.vendorcontact vc on vc.vendorcode = v.vendorcode" &
        '                                           " left join doc.contact c on c.id = vc.contactid" &
        '                                           " left join doc.paramdt prov on prov.paramdtid = fdt.provinceid" &
        '                                           " left join doc.paramdt cty on cty.paramdtid = fdt.countryid" &
        '                                           " left join tu on tu.vendorcode = v.vendorcode" &
        '                                           " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
        '                                           " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
        '                                           " left join doc.paramdt pt on pt.ivalue = vs.producttypeid and pt.paramhdid = 3" &
        '                                            " where  not (vf.factoryid isnull and  vc.contactid isnull) {0} order by v.vendorcode", sb.ToString)
        sqlstrReport = String.Format("select distinct v.vendorcode,vendorname::text,v.shortname::text," &
                                                   " mu2.username as spm,vfp.spmeffectivedate, mu1.username as gsm,gsm.effectivedate as gsmeffectivedate,mu.username as pm,vfp.pmeffectivedate," &
                                                   " pt.paramname as producttype,pr.paramname as status, fhd.customname,fdt.chinesename,fdt.englishname,fdt.englishaddress,fdt.chineseaddress,fdt.area,fdt.city,fdt.main,fdt.modifiedby,doc.getstampfactorydtlaudit(fdt.id) as modifieddate,prov.paramname as province,cty.paramname as country,c.contactname,c.title,c.email,c.officeph,c.factoryph,c.officemb,c.factorymb,upper(c.isecoqualitycontact::text) as isecoqualitycontact from vendor v " &
                                                   "LEFT JOIN doc.viewvendorpmeffectivedate vfp ON vfp.vendorcode = v.vendorcode " &
                                                   " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid " &
                                                   " LEFT JOIN masteruser mu ON mu.id = os.muid " &
                                                   " LEFT JOIN officerseb o ON o.ofsebid = os.parent " &
                                                   " LEFT JOIN masteruser mu2 ON mu2.id = o.muid " &
                                                   " LEFT JOIN doc.vendorgsm gsm ON gsm.vendorcode = v.vendorcode " &
                                                   " LEFT JOIN officerseb o1 ON o1.ofsebid = gsm.gsmid " &
                                                   " LEFT JOIN masteruser mu1 ON mu1.id = o1.muid" &
                                                   " left join doc.vendorfactory vf on vf.vendorcode = v.vendorcode" &
                                                   " left join doc.factorydtl fdt on fdt.id = vf.factoryid" &
                                                   " left join doc.factoryhd fhd on fhd.id = fdt.factoryhdid" &
                                                   " left join doc.vendorcontact vc on vc.vendorcode = v.vendorcode" &
                                                   " left join doc.contact c on c.id = vc.contactid" &
                                                   " left join doc.paramdt prov on prov.paramdtid = fdt.provinceid" &
                                                   " left join doc.paramdt cty on cty.paramdtid = fdt.countryid" &
                                                   " left join doc.vendorstatus vs on vs.vendorcode = v.vendorcode" &
                                                   " left join doc.paramdt pr on pr.ivalue = vs.status and pr.paramhdid = 2" &
                                                   " left join doc.paramdt pt on pt.ivalue = vs.producttypeid and pt.paramhdid = 3" &
                                                    " where  not (vf.factoryid isnull and  vc.contactid isnull) {0} order by v.vendorcode", sb.ToString)
        ' " where not vf.factoryid isnull  {0} order by v.vendorcode", sb.ToString)
        loadReport()


    End Sub

    Private Sub loadReport()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""

            myThread = New Thread(AddressOf DoReport)
            myThread.SetApartmentState(ApartmentState.STA)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoReport()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        ProgressReport(11, "InitFilter")

        sqlstr = String.Format(sqlstrReport, sb.ToString.ToLower)

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("FactoryContact{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 3

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\VendorFactoryContact.xltx")
            myreport.Run(Me, New System.EventArgs)
        End If

        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub


    Private Sub FormattingReport()

    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
        Dim oXl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        oXl = owb.Parent
        owb.Worksheets(3).select()


        'check availability data


        Dim osheet = owb.Worksheets(3)
        Dim orange = osheet.Range("A2")
        If osheet.cells(2, 2).text.ToString = "" Then
            Err.Raise(100, Description:="Data not available.")
        End If
        osheet.name = "RawData"
        owb.Names.Add("db", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C1),COUNTA('RawData'!R1))")


        owb.Worksheets(1).select()
        osheet = owb.Worksheets(1)
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        osheet.pivottables("PivotTable1").SaveData = True
        osheet = owb.Worksheets(2)
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        osheet.pivottables("PivotTable1").SaveData = True
        owb.RefreshAll()
        
    End Sub


End Class