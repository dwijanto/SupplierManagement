Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Imports Microsoft.Office.Interop

Public Class FormSearchDocument
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myThreadSearch As New System.Threading.Thread(AddressOf DoSearch)
    Dim myThreadDownload As New System.Threading.Thread(AddressOf DoDownload)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim DS As DataSet
    Dim DSSearch As DataSet
    Dim VendorBS As BindingSource
    Dim ShortnameBS As BindingSource
    Dim PMBS As BindingSource
    Dim SPMBS As BindingSource
    Dim GSMBS As BindingSource
    Dim ProductTypeBS As BindingSource
    Dim StatusNameBS As BindingSource
    Dim SBUBS As BindingSource
    Dim DocTypeBS As BindingSource
    Dim ProjectNameBS As BindingSource

    Dim docTypeId As Integer
    Dim sb As New StringBuilder
    Dim WithEvents DocumentBS As BindingSource
    Dim SelectedFolder As String = String.Empty
    Private Sub loaddata()
        'MessageBox.Show(myuser)

        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Dim mymessage As String = String.Empty
        Dim sb As New StringBuilder
        'Dim sqlstr As String = "select distinct v1.vendorcode::text || ' - ' || v1.vendorname as name, v1.vendorcode ,v1.vendorname::text from doc.vendordoc vd " &
        '                       " left join vendor v on v.vendorcode = vd.vendorcode" &
        '                       " left join vendor v1 on v1.shortname = v.shortname" &
        '                       " order by v1.vendorname::text;" &
        '                       " select distinct v1.shortname::text from doc.vendordoc vd" &
        '                       " left join vendor v on v.vendorcode = vd.vendorcode " &
        '                       " left join vendor v1 on v1.shortname = v.shortname " &
        '                       " order by v1.shortname::text; " &
        '                       " select distinct o.officersebname::text as spm from doc.vendordoc vd" &
        '                       " left join vendor v on v.vendorcode = vd.vendorcode " &
        '                       " left join officerseb o on o.ofsebid = v.ssmidpl " &
        '                       " order by o.officersebname::text;" &
        '                       " select distinct o.officersebname::text as pm from doc.vendordoc vd" &
        '                       " left join vendor v on v.vendorcode = vd.vendorcode " &
        '                       " left join officerseb o on o.ofsebid = v.pmid " &
        '                       " order by o.officersebname::text ;" &
        '                       " select dt.paramname as producttype from doc.paramdt dt" &
        '                       " where(dt.paramhdid = 3)" &
        '                       " order by ivalue;" &
        '                       " select dt.paramname as statusname from doc.paramdt dt" &
        '                       " where(dt.paramhdid = 2)" &
        '                       " order by ivalue;" &
        '                       " select sbuname from sbusap" &
        '                       " order by sbuname;" &
        '                       " select doctypename,id from doc.doctype" &
        '                       " order by doctypename;" &
        '                       " select distinct vendorcode  from doc.user u" &
        '                       " left join doc.groupuser gu on gu.userid = u.id" &
        '                       " left join doc.groupvendor gv on gv.groupid = gu.groupid" &
        '                       " where lower(u.userid) = '" & HelperClass1.UserId.ToLower & "' order by vendorcode;" &
        '                       " select distinct projectname from doc.project" &
        '                       " order by projectname;"
        'viewvendorpm replace viewvendorfamilypm
        Dim sqlstr As String = "select distinct v1.vendorcode::text || ' - ' || v1.vendorname as name, v1.vendorcode ,v1.vendorname::text from doc.vendordoc vd " &
                       " left join vendor v on v.vendorcode = vd.vendorcode" &
                       " left join vendor v1 on v1.shortname3 = v.shortname3" &
                       " order by v1.vendorname::text;" &
                       " select distinct v1.shortname3::text as shortname from doc.vendordoc vd" &
                       " left join vendor v on v.vendorcode = vd.vendorcode " &
                       " left join vendor v1 on v1.shortname3 = v.shortname3 " &
                       " order by v1.shortname3::text; " &
                       " select distinct mu2.username as spm from doc.vendordoc vd" &
                       " left join doc.viewvendorpm v on v.vendorcode = vd.vendorcode " &
                       " left join officerseb o on o.ofsebid = v.pmid " &
                       " left join officerseb o1 on o1.ofsebid = o.parent" &
                       " LEFT JOIN masteruser mu2 ON mu2.id = o1.muid" &
                       " left join doc.user u on u.userid = o1.userid" &
                       " where u.isactive" &
                       " order by mu2.username;" &
                       " select distinct mu2.username as pm from doc.vendordoc vd" &
                       " left join doc.viewvendorpm v on v.vendorcode = vd.vendorcode " &
                       " left join officerseb o on o.ofsebid = v.pmid " &
                       " LEFT JOIN masteruser mu2 ON mu2.id = o.muid" &
                       " left join doc.user u on u.userid = o.userid" &
                       " where u.isactive" &
                       " order by mu2.username;" &
                       " select dt.paramname as producttype from doc.paramdt dt" &
                       " where(dt.paramhdid = 3)" &
                       " order by ivalue;" &
                       " select dt.paramname as statusname from doc.paramdt dt" &
                       " where(dt.paramhdid = 2)" &
                       " order by ivalue;" &
                       " select sbuname from sbusap" &
                       " order by sbuname;" &
                       " select doctypename,id from doc.doctype" &
                       " order by doctypename;" &
                       " select distinct vendorcode  from doc.user u" &
                       " left join doc.groupuser gu on gu.userid = u.id" &
                       " left join doc.groupvendor gv on gv.groupid = gu.groupid" &
                       " where lower(u.userid) = '" & HelperClass1.UserId.ToLower & "' order by vendorcode;" &
                       " select distinct projectname from doc.project" &
                       " order by projectname;" &
                       " with qgsm as (select distinct vendorcode, first_value (gsmid) over (partition by vendorcode order by vendorcode,effectivedate desc) as gsmid," &
                       " first_value (effectivedate) over (partition by vendorcode order by vendorcode,effectivedate desc) as effectivedate from doc.vendorgsm)" &
                       " select distinct mu2.username as gsm from doc.vendordoc vd" &
                       " left join qgsm   on qgsm.vendorcode = vd.vendorcode" &
                       " left join officerseb o on o.ofsebid = qgsm.gsmid" &
                       " LEFT JOIN masteruser mu2 ON mu2.id = o.muid" &
                       " where not username isnull" &
                       " order by mu2.username;"
        If Not HelperClass1.UserInfo.IsAdmin Then
            'sqlstr = "select distinct v1.vendorcode::text || ' - ' || v1.vendorname as name, v1.vendorcode ,v1.vendorname::text from doc.vendordoc vd " &
            '                   " left join vendor v on v.vendorcode = vd.vendorcode" &
            '                   " left join vendor v1 on v1.shortname = v.shortname" &
            '                   " order by v1.vendorname::text;" &
            '                   " select distinct v1.shortname::text from doc.vendordoc vd" &
            '                   " left join vendor v on v.vendorcode = vd.vendorcode " &
            '                   " left join vendor v1 on v1.shortname = v.shortname " &
            '                   " order by v1.shortname::text; " &


            sb.Append(String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname3::text as shortname " &
                                  " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                  " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where lower(u.userid) ~ '{0}$'   order by vendorname) " &
                                  "    select distinct v1.vendorcode::text || ' - ' || v1.vendorname as name, v1.vendorcode ,v1.vendorname::text from doc.vendordoc vd " &
                                  " inner join va on va.vendorcode = vd.vendorcode" &
                              " left join vendor v on v.vendorcode = va.vendorcode" &
                              " left join vendor v1 on v1.shortname3 = v.shortname3" &
                              " order by v1.vendorname::text;", HelperClass1.UserInfo.userid.ToLower))

            sb.Append(String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname3::text as shortname" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where lower(u.userid) ~ '{0}$'   order by vendorname) " &
                                   "   select distinct v1.shortname3::text as shortname from doc.vendordoc vd" &
                                   " inner join va on va.vendorcode = vd.vendorcode" &
                               " left join vendor v on v.vendorcode = va.vendorcode " &
                               " left join vendor v1 on v1.shortname3 = v.shortname3 " &
                               " order by v1.shortname3::text;", HelperClass1.UserInfo.userid.ToLower))
            'sqlstr = " select distinct o.officersebname::text as spm from doc.vendordoc vd" &
            '                   " left join vendor v on v.vendorcode = vd.vendorcode " &
            '                   " left join officerseb o on o.ofsebid = v.ssmidpl " &
            '                   " order by o.officersebname::text;" &
            '                   " select distinct o.officersebname::text as pm from doc.vendordoc vd" &
            '                   " left join vendor v on v.vendorcode = vd.vendorcode " &
            '                   " left join officerseb o on o.ofsebid = v.pmid " &
            '                   " order by o.officersebname::text ;" &
            '                   " select dt.paramname as producttype from doc.paramdt dt" &
            '                   " where(dt.paramhdid = 3)" &
            '                   " order by ivalue;" &
            '                   " select dt.paramname as statusname from doc.paramdt dt" &
            '                   " where(dt.paramhdid = 2)" &
            '                   " order by ivalue;" &
            '                   " select sbuname from sbusap" &
            '                   " order by sbuname;" &
            '                   " select doctypename,id from doc.doctype" &
            '                   " order by doctypename;" &
            '                   " select distinct vendorcode  from doc.user u" &
            '                   " left join doc.groupuser gu on gu.userid = u.id" &
            '                   " left join doc.groupvendor gv on gv.groupid = gu.groupid" &
            '                   " where lower(u.userid) = '" & HelperClass1.UserId.ToLower & "' order by vendorcode;" &
            '                   " select distinct projectname from doc.project" &
            '                   " order by projectname;"
            'viewvendorpm replace viewvendorfamilypm
            sqlstr = " select distinct mu2.username as spm from doc.vendordoc vd" &
                       " left join doc.viewvendorpm v on v.vendorcode = vd.vendorcode " &
                       " left join officerseb o on o.ofsebid = v.pmid " &
                       " left join officerseb o1 on o1.ofsebid = o.parent" &
                       " LEFT JOIN masteruser mu2 ON mu2.id = o1.muid" &
                       " order by mu2.username;" &
                       " select distinct mu2.username as pm from doc.vendordoc vd" &
                       " left join doc.viewvendorpm v on v.vendorcode = vd.vendorcode " &
                       " left join officerseb o on o.ofsebid = v.pmid " &
                       " LEFT JOIN masteruser mu2 ON mu2.id = o.muid" &
                       " order by mu2.username;" &
                               " select dt.paramname as producttype from doc.paramdt dt" &
                               " where(dt.paramhdid = 3)" &
                               " order by ivalue;" &
                               " select dt.paramname as statusname from doc.paramdt dt" &
                               " where(dt.paramhdid = 2)" &
                               " order by ivalue;" &
                               " select sbuname from sbusap" &
                               " order by sbuname;" &
                               " select doctypename,id from doc.doctype" &
                               " order by doctypename;" &
                               " select distinct vendorcode  from doc.user u" &
                               " left join doc.groupuser gu on gu.userid = u.id" &
                               " left join doc.groupvendor gv on gv.groupid = gu.groupid" &
                               " where lower(u.userid) = '" & HelperClass1.UserId.ToLower & "' order by vendorcode;" &
                               " select distinct projectname from doc.project" &
                               " order by projectname;" &
                               " with qgsm as (select distinct vendorcode, first_value (gsmid) over (partition by vendorcode order by vendorcode,effectivedate desc) as gsmid," &
                               " first_value (effectivedate) over (partition by vendorcode order by vendorcode,effectivedate desc) as effectivedate from doc.vendorgsm)" &
                               " select distinct mu2.username from doc.vendordoc vd" &
                               " left join qgsm   on qgsm.vendorcode = vd.vendorcode" &
                               " left join officerseb o on o.ofsebid = qgsm.gsmid" &
                               " LEFT JOIN masteruser mu2 ON mu2.id = o.muid" &
                               " where not username isnull" &
                               " order by mu2.username;"

        End If


        DS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString & sqlstr, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "Vendor"
                DS.Tables(1).TableName = "ShortName"
                DS.Tables(2).TableName = "SPM"
                DS.Tables(3).TableName = "PM"
                DS.Tables(4).TableName = "ProductType"
                DS.Tables(5).TableName = "StatusName"
                DS.Tables(6).TableName = "SBU"
                DS.Tables(7).TableName = "DocType"
                DS.Tables(8).TableName = "VendorList"
                DS.Tables(9).TableName = "ProjectName"
                DS.Tables(10).TableName = "GSM"
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
        If Not IsNothing(DocumentBS) Then
            ProgressReport(1, String.Format("{0} {1} record(s) found.", "Loading Data.Done!", DocumentBS.Count))
        Else
            ProgressReport(1, String.Format("{0}", "Loading Data.Done!"))
        End If

        ProgressReport(5, "Continuous")
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        'Remember Filter
        TextBox9.Text = My.Settings.FSD_DocumentType
        TextBox1.Text = My.Settings.FSD_VendorCode
        TextBox2.Text = My.Settings.FSD_VendorName
        TextBox3.Text = My.Settings.FSD_ShortName
        TextBox10.Text = My.Settings.FSD_ProjectName
        CheckBox2.Checked = My.Settings.FSD_DocumentDate
        If CheckBox2.Checked Then
            DateTimePicker1.Value = My.Settings.FSD_DTP1       
            DateTimePicker2.Value = My.Settings.FSD_DTP2
        End If
        TextBox4.Text = My.Settings.FSD_SPM
        TextBox5.Text = My.Settings.FSD_PM
        TextBox6.Text = My.Settings.FSD_ProductType
        TextBox7.Text = My.Settings.FSD_VendorStatus
        TextBox8.Text = My.Settings.FSD_SBUName
        If RememberFilter() Then
            CheckBox4.Checked = True
        End If
        ' Add any initialization after the InitializeComponent() call.
        loaddata()
    End Sub
    Private Function RememberFilter() As Boolean
        Dim myreturn As Boolean = False
        If My.Settings.FSD_DocumentType <> "" Then
            myreturn = True
        ElseIf My.Settings.FSD_VendorCode <> "" Then
            myreturn = True
        ElseIf My.Settings.FSD_VendorName <> "" Then
            myreturn = True
        ElseIf My.Settings.FSD_ShortName <> "" Then
            myreturn = True
        ElseIf My.Settings.FSD_ProjectName <> "" Then
            myreturn = True
        ElseIf My.Settings.FSD_DocumentDate <> False Then
            myreturn = True
        ElseIf My.Settings.FSD_SPM <> "" Then
            myreturn = True
        ElseIf My.Settings.FSD_PM <> "" Then
            myreturn = True
        ElseIf My.Settings.FSD_ProductType <> "" Then
            myreturn = True
        ElseIf My.Settings.FSD_VendorStatus <> "" Then
            myreturn = True
        ElseIf My.Settings.FSD_SBUName <> "" Then
            myreturn = True
        End If

        Return myreturn
    End Function
    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked Then
            My.Settings.FSD_DocumentType = TextBox9.Text
            My.Settings.FSD_VendorCode = TextBox1.Text
            My.Settings.FSD_VendorName = TextBox2.Text
            My.Settings.FSD_ShortName = TextBox3.Text
            My.Settings.FSD_ProjectName = TextBox10.Text
            My.Settings.FSD_DocumentDate = CheckBox2.Checked
            If CheckBox2.Checked Then
                My.Settings.FSD_DTP1 = DateTimePicker1.Value.Date
                My.Settings.FSD_DTP2 = DateTimePicker2.Value.Date
            End If
            My.Settings.FSD_SPM = TextBox4.Text
            My.Settings.FSD_PM = TextBox5.Text
            My.Settings.FSD_ProductType = TextBox6.Text
            My.Settings.FSD_VendorStatus = TextBox7.Text
            My.Settings.FSD_SBUName = TextBox8.Text
        Else
            My.Settings.FSD_DocumentType = ""
            My.Settings.FSD_VendorCode = ""
            My.Settings.FSD_VendorName = ""
            My.Settings.FSD_ShortName = ""
            My.Settings.FSD_ProjectName = ""
            My.Settings.FSD_DocumentDate = False
            My.Settings.FSD_SPM = ""
            My.Settings.FSD_PM = ""
            My.Settings.FSD_ProductType = ""
            My.Settings.FSD_VendorStatus = ""
            My.Settings.FSD_SBUName = ""
        End If
        My.Settings.Save()

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
                            VendorBS = New BindingSource
                            ShortnameBS = New BindingSource
                            PMBS = New BindingSource
                            SPMBS = New BindingSource
                            GSMBS = New BindingSource
                            ProductTypeBS = New BindingSource
                            StatusNameBS = New BindingSource
                            SBUBS = New BindingSource
                            DocTypeBS = New BindingSource
                            ProjectNameBS = New BindingSource
                            VendorBS.DataSource = DS.Tables("Vendor")
                            ShortnameBS.DataSource = DS.Tables("ShortName")
                            SPMBS.DataSource = DS.Tables("SPM")
                            PMBS.DataSource = DS.Tables("PM")
                            ProductTypeBS.DataSource = DS.Tables("ProductType")
                            StatusNameBS.DataSource = DS.Tables("StatusName")
                            SBUBS.DataSource = DS.Tables("SBU")
                            DocTypeBS.DataSource = DS.Tables("DocType")
                            ProjectNameBS.DataSource = DS.Tables("ProjectName")
                            GSMBS.DataSource = DS.Tables("GSM")
                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                    Case 7
                        Try
                            DocumentBS = New BindingSource
                            DocumentBS.DataSource = DSSearch.Tables("Document")
                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = DocumentBS

                        Catch ex As Exception
                            message = ex.Message
                        End Try
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click, Button12.Click, Button14.Click
        If myThreadSearch.IsAlive Then
            MessageBox.Show("Please wait until the current process is finished.")
            Exit Sub
        End If
        Try
            Dim mybutton = DirectCast(sender, Button)
            Select Case mybutton.Name
                Case "Button1"
                    If Not IsNothing(VendorBS) Then
                        Dim myform = New FormHelperMulti(VendorBS)
                        myform.DataGridView1.Columns(0).DataPropertyName = "name"
                        myform.Key = 1
                        myform.ShowDialog()
                        TextBox1.Text = TextBox1.Text + IIf(TextBox1.Text = "", "", ",") + myform.SelectedResult
                    End If
                Case "Button2"
                    If Not IsNothing(VendorBS) Then
                        Dim myform = New FormHelperMulti(VendorBS)
                        myform.DataGridView1.Columns(0).DataPropertyName = "name"
                        myform.Key = 2
                        myform.ShowDialog()
                        TextBox2.Text = TextBox2.Text + IIf(TextBox2.Text = "", "", ",") + myform.SelectedResult
                    End If

                Case "Button3"
                    If Not IsNothing(ShortnameBS) Then
                        Dim myform = New FormHelperMulti(ShortnameBS)
                        myform.DataGridView1.Columns(0).DataPropertyName = "shortname"
                        myform.Key = 0
                        myform.ShowDialog()
                        TextBox3.Text = TextBox3.Text + IIf(TextBox3.Text = "", "", ",") + myform.SelectedResult
                    End If

                Case "Button4"
                    If Not IsNothing(SPMBS) Then
                        Dim myform = New FormHelperMulti(SPMBS)
                        myform.DataGridView1.Columns(0).DataPropertyName = "spm"
                        myform.Key = 0
                        myform.ShowDialog()
                        TextBox4.Text = TextBox4.Text + IIf(TextBox4.Text = "", "", ",") + myform.SelectedResult
                    End If

                Case "Button5"
                    If Not IsNothing(PMBS) Then
                        Dim myform = New FormHelperMulti(PMBS)
                        myform.DataGridView1.Columns(0).DataPropertyName = "pm"
                        myform.Key = 0
                        myform.ShowDialog()
                        TextBox5.Text = TextBox5.Text + IIf(TextBox5.Text = "", "", ",") + myform.SelectedResult
                    End If

                Case "Button6"
                    If Not IsNothing(ProductTypeBS) Then
                        Dim myform = New FormHelperMulti(ProductTypeBS)
                        myform.DataGridView1.Columns(0).DataPropertyName = "producttype"
                        myform.Key = 0
                        myform.ShowDialog()
                        TextBox6.Text = TextBox6.Text + IIf(TextBox6.Text = "", "", ",") + myform.SelectedResult
                    End If

                Case "Button7"
                    If Not IsNothing(StatusNameBS) Then
                        Dim myform = New FormHelperMulti(StatusNameBS)
                        myform.DataGridView1.Columns(0).DataPropertyName = "statusname"
                        myform.Key = 0
                        myform.ShowDialog()
                        TextBox7.Text = TextBox7.Text + IIf(TextBox7.Text = "", "", ",") + myform.SelectedResult
                    End If

                Case "Button8"
                    If Not IsNothing(SBUBS) Then
                        Dim myform = New FormHelperMulti(SBUBS)
                        myform.DataGridView1.Columns(0).DataPropertyName = "sbuname"
                        'myform.DataGridView1.MultiSelect = False
                        myform.Key = 0
                        myform.ShowDialog()
                        TextBox8.Text = TextBox8.Text + IIf(TextBox8.Text = "", "", ",") + myform.SelectedResult
                    End If

                Case "Button9"
                    If Not IsNothing(DocTypeBS) Then
                        Dim myform = New FormHelperMulti(DocTypeBS)
                        myform.DataGridView1.Columns(0).DataPropertyName = "doctypename"
                        myform.DataGridView1.MultiSelect = True
                        myform.Key = 0
                        myform.ShowDialog()
                        'TextBox9.Text = myform.SelectedResult
                        TextBox9.Text = TextBox9.Text + IIf(TextBox9.Text = "", "", ",") + myform.SelectedResult
                    End If
                Case "Button14"
                    If Not IsNothing(GSMBS) Then
                        Dim myform = New FormHelperMulti(GSMBS)
                        myform.DataGridView1.Columns(0).DataPropertyName = "gsm"
                        myform.Key = 0
                        myform.ShowDialog()
                        TextBox11.Text = TextBox11.Text + IIf(TextBox11.Text = "", "", ",") + myform.SelectedResult
                    End If
                Case "Button12"
                    If Not IsNothing(ProjectNameBS) Then
                        Dim myform = New FormHelperMulti(ProjectNameBS)
                        myform.DataGridView1.Columns(0).DataPropertyName = "projectname"
                        'myform.DataGridView1.MultiSelect = False
                        myform.Key = 0
                        myform.ShowDialog()
                        TextBox10.Text = TextBox10.Text + IIf(TextBox10.Text = "", "", ",") + myform.SelectedResult
                    End If
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try



    End Sub




    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If myThread.IsAlive Then
            MessageBox.Show("Please wait until the current process is finished.")
            Exit Sub
        End If

        Dim mymessage As String = String.Empty
        Dim sqlstr As String = String.Empty
        sb = New StringBuilder

        ToolStripStatusLabel1.Text = "Processing... Please wait!"
        ToolStripStatusLabel2.Text = ""

        'If TextBox9.Text <> "" Then
        '    Dim drv As DataRowView = DocTypeBS.Current
        '    docTypeId = drv.Row.Item("id")       
        'End If

        If Not CheckBox3.Checked Then
            If HelperClass1.UserInfo.IsAdmin Then
                sb.Append("SELECT false as download, clt.id, clt.vendorcode, v.vendorname::text AS vendorname, v.shortname3::text AS shortname, o.officersebname::text AS spm, mu1.username AS gsm, mu2.username AS spm, mu.username AS pm, d.id AS documentid, dt.doctypename, dl.levelname, de.expireddate, d.docdate, d.docname, d.docext, coalesce(d.uploaddate,vm.applicantdate) as uploaddate, coalesce(d.userid,creator) as userid, d.remarks, vr.version, (pt.payt::text || ' - '::text) || pt.details::text AS paymentterm, sc.leadtime, sc.sasl, q.nqsu, p.projectname,ps.returnrate, sa.auditby, sa.audittype, sa.auditgrade, sef.score, sif.myyear, sif.turnovery, sif.turnovery1, sif.turnovery2, sif.turnovery3, sif.turnovery4, sif.ratioy, sif.ratioy1, sif.ratioy2, sif.ratioy3, sif.ratioy4, scy.category, ps1.panelstatus AS panelstatus1, ps1.paneldescription AS paneldescription1, ps2.panelstatus AS panelstatus2, ps2.paneldescription AS paneldescription2, " &
                 " CASE WHEN dt.id = 32 THEN 1 ELSE NULL::integer END AS generalcontract, CASE WHEN dt.id = 33 THEN 1 ELSE NULL::integer END AS qualityappendix, CASE WHEN dt.id = 35 THEN 1 ELSE NULL::integer END AS supplychainappendix,null::integer  AS countcontractisfull, " &
                 " null::integer AS countcontractfull,  null::integer as countcontractnotfull, vg1.groupsbuname AS fp, vg2.groupsbuname AS cp, vg3.groupsbuname AS sp, vg4.groupsbuname AS mould, doc.groupact(vg1.groupsbuname::character varying, vg2.groupsbuname::character varying, vg3.groupsbuname::character varying, vg4.groupsbuname::character varying, si.fpcp::character varying) AS groupact, pdt.producttype, pr.paramname AS statusname, spp.rank,vsbu.sbuname ")
                sb.Append(" from doc.allvendordocument clt")
                'sb.Append(" left join doc.document d on d.id = clt.id " &
                '          " LEFT JOIN vendor v ON v.vendorcode = clt.vendorcode" &
                '          " LEFT JOIN doc.version vr ON vr.documentid = d.id" &
                '          " LEFT JOIN doc.generalcontract gt ON gt.documentid = d.id" &
                '          " LEFT JOIN paymentterm pt ON pt.paymenttermid = gt.paymentcode" &
                '          " LEFT JOIN doc.supplychain sc ON sc.documentid = d.id" &
                '          " LEFT JOIN doc.qualityappendix q ON q.documentid = d.id" &
                '          " LEFT JOIN doc.project p ON p.documentid = d.id" &
                '          " LEFT JOIN doc.socialaudit sa ON sa.documentid = d.id" &
                '          " LEFT JOIN doc.sef sef ON sef.documentid = d.id" &
                '          " LEFT JOIN doc.sif sif ON sif.documentid = d.id" &
                '          " LEFT JOIN doc.doctype dt ON dt.id = d.doctypeid" &
                '          " LEFT JOIN doc.doclevel dl ON dl.id = d.doclevelid" &
                '          " LEFT JOIN doc.docexpired de ON de.documentid = d.id" &
                '          " LEFT JOIN supplierspanel spp ON spp.vendorcode = clt.vendorcode" &
                '          " LEFT JOIN supplierscategory scy ON scy.supplierscategoryid = spp.supplierscategoryid" &
                '          " LEFT JOIN doc.panelstatus ps1 ON ps1.id = spp.fp" &
                '          " LEFT JOIN doc.panelstatus ps2 ON ps2.id = spp.cp" &
                '          " LEFT JOIN officerseb o ON o.ofsebid = v.ssmidpl" &
                '          " LEFT JOIN officerseb os ON os.ofsebid = v.pmid" &
                '          " LEFT JOIN doc.tvendorgroupsbu vg1 ON vg1.vendorcode = v.vendorcode AND vg1.groupsbuname = 'FP'::text" &
                '          " LEFT JOIN doc.tvendorgroupsbu vg2 ON vg2.vendorcode = v.vendorcode AND vg2.groupsbuname = 'CP'::text" &
                '          " LEFT JOIN doc.tvendorgroupsbu vg3 ON vg3.vendorcode = v.vendorcode AND vg3.groupsbuname = 'SP'::text" &
                '          " LEFT JOIN doc.tvendorgroupsbu vg4 ON vg4.vendorcode = v.vendorcode AND vg4.groupsbuname = 'MOULD'::text" &
                '          " LEFT JOIN doc.vendorstatus vs ON vs.vendorcode = v.vendorcode" &
                '          " LEFT JOIN doc.paramhd ph ON ph.paramname::text = 'vendorstatus'::text" &
                '          " LEFT JOIN doc.paramdt pr ON pr.ivalue = vs.status AND pr.paramhdid = ph.paramhdid" &
                '          " LEFT JOIN doc.shortnameinfo si ON si.shortname = v.shortname::text" &
                '          " LEFT JOIN ( SELECT pd.ivalue, pd.paramname AS producttype" &
                '          " FROM doc.paramdt pd" &
                '          " LEFT JOIN doc.paramhd ph ON ph.paramhdid = pd.paramhdid" &
                '          " WHERE ph.paramname::text = 'producttype'::text) pdt ON pdt.ivalue = vs.producttypeid" &
                '          " left join doc.vendorsbu vsbu on vsbu.vendorcode = clt.vendorcode" &
                '          " WHERE(Not clt.vendorcode Is NULL)")
                'viewvendorpm replace viewvendorfamilypm
                sb.Append(" left join doc.document d on d.id = clt.id " &
                          " LEFT JOIN vendor v ON v.vendorcode = clt.vendorcode" &
                          " LEFT JOIN doc.viewvendorpm vfp ON vfp.vendorcode = v.vendorcode" &
                          " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid" &
                          " LEFT JOIN masteruser mu ON mu.id = os.muid" &
                          " LEFT JOIN officerseb o ON o.ofsebid = os.parent" &
                          " LEFT JOIN masteruser mu2 ON mu2.id = o.muid" &
                          " LEFT JOIN doc.vendorgsm gsm ON gsm.vendorcode = v.vendorcode" &
                          " LEFT JOIN officerseb o1 ON o1.ofsebid = gsm.gsmid" &
                          " LEFT JOIN masteruser mu1 ON mu1.id = o1.muid" &
                          " LEFT JOIN doc.version vr ON vr.documentid = d.id" &
                          " LEFT JOIN doc.generalcontract gt ON gt.documentid = d.id" &
                          " LEFT JOIN paymentterm pt ON pt.paymenttermid = gt.paymentcode" &
                          " LEFT JOIN doc.supplychain sc ON sc.documentid = d.id" &
                          " LEFT JOIN doc.qualityappendix q ON q.documentid = d.id" &
                          " LEFT JOIN doc.project p ON p.documentid = d.id" &
                          " LEFT JOIN doc.projectspecification ps ON ps.documentid = d.id" &
                          " LEFT JOIN doc.socialaudit sa ON sa.documentid = d.id" &
                          " LEFT JOIN doc.sef sef ON sef.documentid = d.id" &
                          " LEFT JOIN doc.sif sif ON sif.documentid = d.id" &
                          " LEFT JOIN doc.doctype dt ON dt.id = d.doctypeid" &
                          " LEFT JOIN doc.doclevel dl ON dl.id = d.doclevelid" &
                          " LEFT JOIN doc.docexpired de ON de.documentid = d.id" &
                          " LEFT JOIN supplierspanel spp ON spp.vendorcode = clt.vendorcode" &
                          " LEFT JOIN supplierscategory scy ON scy.supplierscategoryid = spp.supplierscategoryid" &
                          " LEFT JOIN doc.panelstatus ps1 ON ps1.id = spp.fp" &
                          " LEFT JOIN doc.panelstatus ps2 ON ps2.id = spp.cp" &
                          " LEFT JOIN doc.tvendorgroupsbu vg1 ON vg1.vendorcode = v.vendorcode AND vg1.groupsbuname = 'FP'::text" &
                          " LEFT JOIN doc.tvendorgroupsbu vg2 ON vg2.vendorcode = v.vendorcode AND vg2.groupsbuname = 'CP'::text" &
                          " LEFT JOIN doc.tvendorgroupsbu vg3 ON vg3.vendorcode = v.vendorcode AND vg3.groupsbuname = 'SP'::text" &
                          " LEFT JOIN doc.tvendorgroupsbu vg4 ON vg4.vendorcode = v.vendorcode AND vg4.groupsbuname = 'MOULD'::text" &
                          " LEFT JOIN doc.vendorstatus vs ON vs.vendorcode = v.vendorcode" &
                          " LEFT JOIN doc.paramhd ph ON ph.paramname::text = 'vendorstatus'::text" &
                          " LEFT JOIN doc.paramdt pr ON pr.ivalue = vs.status AND pr.paramhdid = ph.paramhdid" &
                          " LEFT JOIN doc.shortnameinfo si ON si.shortname = v.shortname3::text" &
                          " left join doc.vendorinfmodiattachment vat on vat.documentid = d.id" &
                          " left join doc.vendorinfmodi vm on vm.id = vat.vendorinfmodiid" &
                          " LEFT JOIN ( SELECT pd.ivalue, pd.paramname AS producttype" &
                          " FROM doc.paramdt pd" &
                          " LEFT JOIN doc.paramhd ph ON ph.paramhdid = pd.paramhdid" &
                          " WHERE ph.paramname::text = 'producttype'::text) pdt ON pdt.ivalue = vs.producttypeid" &
                          " left join doc.vendorsbu vsbu on vsbu.vendorcode = clt.vendorcode" &
                          " WHERE(Not clt.vendorcode Is NULL)")
            Else
                sb.Append(String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname3::text as shortname" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where lower(u.userid) ~ '{0}$'   order by vendorname) ", HelperClass1.UserInfo.userid.ToLower) &
                    "SELECT false as download, clt.id, clt.vendorcode, v.vendorname::text AS vendorname, v.shortname3::text AS shortname, mu1.username AS gsm, mu2.username AS spm, mu.username AS pm, d.id AS documentid, dt.doctypename, dl.levelname, de.expireddate, d.docdate, d.docname, d.docext, coalesce(d.uploaddate,vm.applicantdate) as uploaddate, coalesce(d.userid,creator) as userid, d.remarks, vr.version, (pt.payt::text || ' - '::text) || pt.details::text AS paymentterm, sc.leadtime, sc.sasl, q.nqsu, p.projectname,ps.returnrate, sa.auditby, sa.audittype, sa.auditgrade, sef.score, sif.myyear, sif.turnovery, sif.turnovery1, sif.turnovery2, sif.turnovery3, sif.turnovery4, sif.ratioy, sif.ratioy1, sif.ratioy2, sif.ratioy3, sif.ratioy4, scy.category, ps1.panelstatus AS panelstatus1, ps1.paneldescription AS paneldescription1, ps2.panelstatus AS panelstatus2, ps2.paneldescription AS paneldescription2, " &
                 " CASE WHEN dt.id = 32 THEN 1 ELSE NULL::integer END AS generalcontract, CASE WHEN dt.id = 33 THEN 1 ELSE NULL::integer END AS qualityappendix, CASE WHEN dt.id = 35 THEN 1 ELSE NULL::integer END AS supplychainappendix,null::integer  AS countcontractisfull, " &
                 " null::integer AS countcontractfull,  null::integer as countcontractnotfull, vg1.groupsbuname AS fp, vg2.groupsbuname AS cp, vg3.groupsbuname AS sp, vg4.groupsbuname AS mould, doc.groupact(vg1.groupsbuname::character varying, vg2.groupsbuname::character varying, vg3.groupsbuname::character varying, vg4.groupsbuname::character varying, si.fpcp::character varying) AS groupact, pdt.producttype, pr.paramname AS statusname, spp.rank,vsbu.sbuname ")
                sb.Append(" from doc.allvendordocument clt")
                'sb.Append(" left join doc.document d on d.id = clt.id " &
                '          " inner join va on va.vendorcode = clt.vendorcode" &
                '          " LEFT JOIN vendor v ON v.vendorcode = va.vendorcode" &
                '          " LEFT JOIN doc.version vr ON vr.documentid = d.id" &
                '          " LEFT JOIN doc.generalcontract gt ON gt.documentid = d.id" &
                '          " LEFT JOIN paymentterm pt ON pt.paymenttermid = gt.paymentcode" &
                '          " LEFT JOIN doc.supplychain sc ON sc.documentid = d.id" &
                '          " LEFT JOIN doc.qualityappendix q ON q.documentid = d.id" &
                '          " LEFT JOIN doc.project p ON p.documentid = d.id" &
                '          " LEFT JOIN doc.socialaudit sa ON sa.documentid = d.id" &
                '          " LEFT JOIN doc.sef sef ON sef.documentid = d.id" &
                '          " LEFT JOIN doc.sif sif ON sif.documentid = d.id" &
                '          " LEFT JOIN doc.doctype dt ON dt.id = d.doctypeid" &
                '          " LEFT JOIN doc.doclevel dl ON dl.id = d.doclevelid" &
                '          " LEFT JOIN doc.docexpired de ON de.documentid = d.id" &
                '          " LEFT JOIN supplierspanel spp ON spp.vendorcode = clt.vendorcode" &
                '          " LEFT JOIN supplierscategory scy ON scy.supplierscategoryid = spp.supplierscategoryid" &
                '          " LEFT JOIN doc.panelstatus ps1 ON ps1.id = spp.fp" &
                '          " LEFT JOIN doc.panelstatus ps2 ON ps2.id = spp.cp" &
                '          " LEFT JOIN officerseb o ON o.ofsebid = v.ssmidpl" &
                '          " LEFT JOIN officerseb os ON os.ofsebid = v.pmid" &
                '          " LEFT JOIN doc.tvendorgroupsbu vg1 ON vg1.vendorcode = v.vendorcode AND vg1.groupsbuname = 'FP'::text" &
                '          " LEFT JOIN doc.tvendorgroupsbu vg2 ON vg2.vendorcode = v.vendorcode AND vg2.groupsbuname = 'CP'::text" &
                '          " LEFT JOIN doc.tvendorgroupsbu vg3 ON vg3.vendorcode = v.vendorcode AND vg3.groupsbuname = 'SP'::text" &
                '          " LEFT JOIN doc.tvendorgroupsbu vg4 ON vg4.vendorcode = v.vendorcode AND vg4.groupsbuname = 'MOULD'::text" &
                '          " LEFT JOIN doc.vendorstatus vs ON vs.vendorcode = v.vendorcode" &
                '          " LEFT JOIN doc.paramhd ph ON ph.paramname::text = 'vendorstatus'::text" &
                '          " LEFT JOIN doc.paramdt pr ON pr.ivalue = vs.status AND pr.paramhdid = ph.paramhdid" &
                '          " LEFT JOIN doc.shortnameinfo si ON si.shortname = v.shortname::text" &
                '          " LEFT JOIN ( SELECT pd.ivalue, pd.paramname AS producttype" &
                '          " FROM doc.paramdt pd" &
                '          " LEFT JOIN doc.paramhd ph ON ph.paramhdid = pd.paramhdid" &
                '          " WHERE ph.paramname::text = 'producttype'::text) pdt ON pdt.ivalue = vs.producttypeid" &
                '          " left join doc.vendorsbu vsbu on vsbu.vendorcode = clt.vendorcode" &
                '          " WHERE(Not clt.vendorcode Is NULL)")


                'doc.viewvendorpm replace doc.viewvendorfamilypm

                sb.Append(" left join doc.document d on d.id = clt.id " &
                          " inner join va on va.vendorcode = clt.vendorcode" &
                          " LEFT JOIN vendor v ON v.vendorcode = va.vendorcode" &
                          " LEFT JOIN doc.viewvendorpm vfp ON vfp.vendorcode = v.vendorcode" &
                          " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid" &
                          " LEFT JOIN masteruser mu ON mu.id = os.muid" &
                          " LEFT JOIN officerseb o ON o.ofsebid = os.parent" &
                          " LEFT JOIN masteruser mu2 ON mu2.id = o.muid" &
                          " LEFT JOIN doc.vendorgsm gsm ON gsm.vendorcode = v.vendorcode" &
                          " LEFT JOIN officerseb o1 ON o1.ofsebid = gsm.gsmid" &
                          " LEFT JOIN masteruser mu1 ON mu1.id = o1.muid" &
                          " LEFT JOIN doc.version vr ON vr.documentid = d.id" &
                          " LEFT JOIN doc.generalcontract gt ON gt.documentid = d.id" &
                          " LEFT JOIN paymentterm pt ON pt.paymenttermid = gt.paymentcode" &
                          " LEFT JOIN doc.supplychain sc ON sc.documentid = d.id" &
                          " LEFT JOIN doc.qualityappendix q ON q.documentid = d.id" &
                          " LEFT JOIN doc.project p ON p.documentid = d.id" &
                          " LEFT JOIN doc.projectspecification ps ON ps.documentid = d.id" &
                          " LEFT JOIN doc.socialaudit sa ON sa.documentid = d.id" &
                          " LEFT JOIN doc.sef sef ON sef.documentid = d.id" &
                          " LEFT JOIN doc.sif sif ON sif.documentid = d.id" &
                          " LEFT JOIN doc.doctype dt ON dt.id = d.doctypeid" &
                          " LEFT JOIN doc.doclevel dl ON dl.id = d.doclevelid" &
                          " LEFT JOIN doc.docexpired de ON de.documentid = d.id" &
                          " LEFT JOIN supplierspanel spp ON spp.vendorcode = clt.vendorcode" &
                          " LEFT JOIN supplierscategory scy ON scy.supplierscategoryid = spp.supplierscategoryid" &
                          " LEFT JOIN doc.panelstatus ps1 ON ps1.id = spp.fp" &
                          " LEFT JOIN doc.panelstatus ps2 ON ps2.id = spp.cp" &
                            " LEFT JOIN doc.tvendorgroupsbu vg1 ON vg1.vendorcode = v.vendorcode AND vg1.groupsbuname = 'FP'::text" &
                          " LEFT JOIN doc.tvendorgroupsbu vg2 ON vg2.vendorcode = v.vendorcode AND vg2.groupsbuname = 'CP'::text" &
                          " LEFT JOIN doc.tvendorgroupsbu vg3 ON vg3.vendorcode = v.vendorcode AND vg3.groupsbuname = 'SP'::text" &
                          " LEFT JOIN doc.tvendorgroupsbu vg4 ON vg4.vendorcode = v.vendorcode AND vg4.groupsbuname = 'MOULD'::text" &
                          " LEFT JOIN doc.vendorstatus vs ON vs.vendorcode = v.vendorcode" &
                          " LEFT JOIN doc.paramhd ph ON ph.paramname::text = 'vendorstatus'::text" &
                          " LEFT JOIN doc.paramdt pr ON pr.ivalue = vs.status AND pr.paramhdid = ph.paramhdid" &
                          " LEFT JOIN doc.shortnameinfo si ON si.shortname = v.shortname3::text" &
                          " left join doc.vendorinfmodiattachment vat on vat.documentid = d.id" &
                          " left join doc.vendorinfmodi vm on vm.id = vat.vendorinfmodiid" &
                          " LEFT JOIN ( SELECT pd.ivalue, pd.paramname AS producttype" &
                          " FROM doc.paramdt pd" &
                          " LEFT JOIN doc.paramhd ph ON ph.paramhdid = pd.paramhdid" &
                          " WHERE ph.paramname::text = 'producttype'::text) pdt ON pdt.ivalue = vs.producttypeid" &
                          " left join doc.vendorsbu vsbu on vsbu.vendorcode = clt.vendorcode" &
                          " WHERE(Not clt.vendorcode Is NULL)")

            End If
            
            'If TextBox9.Text <> "" Then

            '    ' sb.Append(" and '" & TextBox9.Text.Replace("'", "''") & ",' ~ (dt.doctypename || ',')::text")        
            '    'If TextBox9.Text.Contains(")") Or TextBox9.Text.Contains("(") Then
            '    '    sb.Append(" and (dt.doctypename || ',')::text ~ '^" & TextBox9.Text.Replace("'", "''").Replace("(", "\(").Replace(")", "\)") & ",'")
            '    'Else
            '    '    sb.Append(" and '" & TextBox9.Text & ",'" & " ~(dt.doctypename || ',')::text")
            '    'End If
            '    Dim myrecord() As String = TextBox9.Text.Split(",")
            '    Dim mycr As StringBuilder = New StringBuilder
            '    For i = 0 To myrecord.Length - 1
            '        If mycr.Length > 0 Then
            '            mycr.Append(",")
            '        End If
            '        mycr.Append(String.Format("'{0},'", myrecord(i)))
            '    Next

            '    sb.Append(String.Format(" and dt.doctypename || ',' in ({0})", mycr.ToString))

            'End If

            If TextBox9.Text <> "" Then
                'sb.Append(" and '" & TextBox9.Text & ",'" & " ~(dt.doctypename || ',')::text")
                sb.Append(" and '" & TextBox9.Text & "'" & " ~(replace(replace(dt.doctypename,'(','\('),')','\)'))")
                'sb.Append(" and dt.doctypename ~(replace(replace( '" & TextBox9.Text & "','(','\('),')','\)'))")
                'sb.Append(String.Format(" and dt.doctypename ~({0})", validreg(TextBox9.Text)))
            End If

            If TextBox1.Text <> "" Then
                sb.Append(" and '" & TextBox1.Text.Replace("'", "''") & "' ~ clt.vendorcode::text")
            End If

            If TextBox2.Text <> "" Then
                'sb.Append(" and '" & TextBox2.Text.Replace("'", "''") & "' ~ v.vendorname::text")
                sb.Append(" and (v.vendorname::text || ',')::text ~ '^" & TextBox2.Text.Replace("'", "''").Replace("(", "\(").Replace(")", "\)") & ",'")
                'sb.Append(String.Format(" and '{0}' ~ v.vendorname::text", TextBox2.Text.Replace("'", "''").Replace(")", "\)").Replace("(", "\(")))
            End If

            If TextBox3.Text <> "" Then
                sb.Append(" and '" & TextBox3.Text.Replace("'", "''") & "' ~ v.shortname3::text")
            End If

            'If TextBox4.Text <> "" Then
            '    sb.Append(" and '" & TextBox4.Text.Replace("'", "''") & "' ~ o.officersebname::text")
            'End If

            'If TextBox5.Text <> "" Then
            '    sb.Append(" and '" & TextBox5.Text.Replace("'", "''") & "' ~ os.officersebname::text")
            'End If
            If TextBox4.Text <> "" Then
                sb.Append(" and '" & TextBox4.Text.Replace("'", "''") & "' ~ mu2.username")
            End If

            If TextBox5.Text <> "" Then
                sb.Append(" and '" & TextBox5.Text.Replace("'", "''") & "' ~ mu.username")
            End If

            If TextBox6.Text <> "" Then
                'sb.Append(" and '" & TextBox6.Text.Replace("'", "''") & "' ~ pdt.producttype::text")
                Dim ProductTypeSB As New StringBuilder
                Dim myProdType = TextBox6.Text.ToString.Split(",")
                For i = 0 To myProdType.Length - 1
                    If ProductTypeSB.Length = 0 Then
                        ProductTypeSB.Append(" and (")
                    Else
                        ProductTypeSB.Append(" or ")
                    End If

                    ProductTypeSB.Append(" pdt.producttype::text ='" & myProdType(i).Replace("'", "''") & "'")
                Next
                ProductTypeSB.Append(")")
                sb.Append(ProductTypeSB.ToString)
            End If

            If TextBox7.Text <> "" Then
                sb.Append(" and '" & TextBox7.Text.Replace("'", "''") & "' ~ pr.paramname::text")
            End If

            If TextBox8.Text <> "" Then
                'sb.Append(" and '" & TextBox8.Text.Replace("'", "''") & "' ~ vsbu.sbuname::text")
                Dim SBUNameSB As New StringBuilder
                Dim mysbu = TextBox8.Text.ToString.Split(",")
                For i = 0 To mysbu.Length - 1
                    If SBUNameSB.Length = 0 Then
                        SBUNameSB.Append(" and (")
                    Else
                        SBUNameSB.Append(" or ")
                    End If

                    SBUNameSB.Append(" vsbu.sbuname ~'" & mysbu(i).Replace("'", "''") & "'")
                Next
                SBUNameSB.Append(")")
                sb.Append(SBUNameSB.ToString)
                'sb.Append(" and vsbu.sbuname ~'" & TextBox8.Text.Replace("'", "''") & "'")
            End If

            If TextBox10.Text <> "" Then
                'sb.Append(" and '" & TextBox10.Text.Replace("'", "''") & "' ~*(replace(replace( '^'|| projectname || '$','(','\('),')','\)'))")
                sb.Append(" and '" & TextBox10.Text.Replace("'", "''") & ",'~replace(replace(replace(projectname || ',','\','\\'),'(','\('),')','\)')")
                'sb.Append(" and projectname like '%" & TextBox10.Text.Replace("'", "''") & "%'")
            End If

            If TextBox11.Text <> "" Then
                sb.Append(" and '" & TextBox11.Text.Replace("'", "''") & "' ~ mu1.username")
            End If

            'If Not HelperClass1.UserInfo.IsAdmin Then
            '    'Generate list
            '    Dim mylist As New StringBuilder
            '    For Each drv As DataRow In DS.Tables("VendorList").Rows
            '        If mylist.Length > 0 Then
            '            mylist.Append(",")
            '        End If
            '        mylist.Append(drv.Item("vendorcode"))
            '    Next
            '    sb.Append(String.Format(" and '{0}' ~ v.vendorcode::text", mylist.ToString))
            'End If

            If CheckBox2.Checked Then
                sb.Append(String.Format(" and  d.docdate >='{0:yyyy-MM-dd}' and d.docdate <='{1:yyyy-MM-dd}'", DateTimePicker1.Value.Date, DateTimePicker2.Value.Date))
            End If

            sb.Append(" ORDER BY v.shortname3::text, dt.doctypename;")
        Else
            If HelperClass1.UserInfo.IsAdmin Then
                'sb.Append("with basedata as (select distinct shortname,doctypeid,id from (select shortname,doctypeid, d.docdate, first_value(d.id) over (partition by doctypeid,shortname order by doctypeid,shortname,docdate desc) as id from doc.allvendordocument av" &
                '      " left join doc.document d on d.id = av.id" &
                '      " order by doctypeid,shortname,docdate desc)foo order by shortname,doctypeid,id)" &
                '      " SELECT false as download, clt.id, v.vendorcode,v.vendorname::text AS vendorname, v.shortname::text AS shortname, o.officersebname::text AS spm, os.officersebname::text AS pm, d.id AS documentid, dt.doctypename, dl.levelname, de.expireddate, d.docdate, d.docname, d.docext, d.uploaddate, d.userid, d.remarks, vr.version, (pt.payt::text || ' - '::text) || pt.details::text AS paymentterm, sc.leadtime, sc.sasl, q.nqsu, p.projectname, sa.auditby, sa.audittype, sa.auditgrade, sef.score, sif.myyear, sif.turnovery, sif.turnovery1, sif.turnovery2, sif.turnovery3, sif.turnovery4, sif.ratioy, sif.ratioy1, sif.ratioy2, sif.ratioy3, sif.ratioy4," &
                '      " CASE WHEN dt.id = 32 THEN 1 ELSE NULL::integer END AS generalcontract, CASE WHEN dt.id = 33 THEN 1 ELSE NULL::integer END AS qualityappendix, CASE WHEN dt.id = 35 THEN 1 ELSE NULL::integer END AS supplychainappendix,null::integer  AS countcontractisfull,  null::integer AS countcontractfull,  null::integer as countcontractnotfull, vg1.groupsbuname AS fp, vg2.groupsbuname AS cp, vg3.groupsbuname AS sp, vg4.groupsbuname AS mould, doc.groupact(vg1.groupsbuname::character varying, vg2.groupsbuname::character varying, vg3.groupsbuname::character varying, vg4.groupsbuname::character varying, si.fpcp::character varying) AS groupact, pdt.producttype, pr.paramname AS statusname" &
                '      " from basedata clt" &
                '      " left join doc.document d on d.id = clt.id  LEFT JOIN vendor v ON v.shortname = clt.shortname LEFT JOIN doc.version vr ON vr.documentid = d.id LEFT JOIN doc.generalcontract gt ON gt.documentid = d.id LEFT JOIN paymentterm pt ON pt.paymenttermid = gt.paymentcode LEFT JOIN doc.supplychain sc ON sc.documentid = d.id LEFT JOIN doc.qualityappendix q ON q.documentid = d.id LEFT JOIN doc.project p ON p.documentid = d.id LEFT JOIN doc.socialaudit sa ON sa.documentid = d.id LEFT JOIN doc.sef sef ON sef.documentid = d.id LEFT JOIN doc.sif sif ON sif.documentid = d.id LEFT JOIN doc.doctype dt ON dt.id = d.doctypeid LEFT JOIN doc.doclevel dl ON dl.id = d.doclevelid LEFT JOIN doc.docexpired de ON de.documentid = d.id " &
                '      " LEFT JOIN officerseb o ON o.ofsebid = v.ssmidpl LEFT JOIN officerseb os ON os.ofsebid = v.pmid LEFT JOIN doc.tvendorgroupsbu vg1 ON vg1.vendorcode = v.vendorcode AND vg1.groupsbuname = 'FP'::text LEFT JOIN doc.tvendorgroupsbu vg2 ON vg2.vendorcode = v.vendorcode AND vg2.groupsbuname = 'CP'::text LEFT JOIN doc.tvendorgroupsbu vg3 ON vg3.vendorcode = v.vendorcode AND vg3.groupsbuname = 'SP'::text LEFT JOIN doc.tvendorgroupsbu vg4 ON vg4.vendorcode = v.vendorcode AND vg4.groupsbuname = 'MOULD'::text LEFT JOIN doc.vendorstatus vs ON vs.vendorcode = v.vendorcode LEFT JOIN doc.paramhd ph ON ph.paramname::text = 'vendorstatus'::text LEFT JOIN doc.paramdt pr ON pr.ivalue = vs.status AND pr.paramhdid = ph.paramhdid LEFT JOIN doc.shortnameinfo si ON si.shortname = v.shortname::text LEFT JOIN ( SELECT pd.ivalue, pd.paramname AS producttype FROM doc.paramdt pd LEFT JOIN doc.paramhd ph ON ph.paramhdid = pd.paramhdid WHERE ph.paramname::text = 'producttype'::text) pdt ON pdt.ivalue = vs.producttypeid where not clt.id isnull")

                'viewvendorpm replace viewvendorfamilypm
                sb.Append("with basedata as (select distinct shortname,doctypeid,id from (select shortname,doctypeid, d.docdate, first_value(d.id) over (partition by doctypeid,shortname order by doctypeid,shortname,docdate desc) as id from doc.allvendordocument av" &
                     " left join doc.document d on d.id = av.id" &
                     " order by doctypeid,shortname,docdate desc)foo order by shortname,doctypeid,id)" &
                     " SELECT false as download, clt.id, v.vendorcode,v.vendorname::text AS vendorname, v.shortname3::text AS shortname, mu1.username AS gsm, mu2.username AS spm, mu.username AS pm, d.id AS documentid, dt.doctypename, dl.levelname, de.expireddate, d.docdate, d.docname, d.docext, d.uploaddate, d.userid, d.remarks, vr.version, (pt.payt::text || ' - '::text) || pt.details::text AS paymentterm, sc.leadtime, sc.sasl, q.nqsu, p.projectname,ps.returnrate, sa.auditby, sa.audittype, sa.auditgrade, sef.score, sif.myyear, sif.turnovery, sif.turnovery1, sif.turnovery2, sif.turnovery3, sif.turnovery4, sif.ratioy, sif.ratioy1, sif.ratioy2, sif.ratioy3, sif.ratioy4," &
                     " CASE WHEN dt.id = 32 THEN 1 ELSE NULL::integer END AS generalcontract, CASE WHEN dt.id = 33 THEN 1 ELSE NULL::integer END AS qualityappendix, CASE WHEN dt.id = 35 THEN 1 ELSE NULL::integer END AS supplychainappendix,null::integer  AS countcontractisfull,  null::integer AS countcontractfull,  null::integer as countcontractnotfull, vg1.groupsbuname AS fp, vg2.groupsbuname AS cp, vg3.groupsbuname AS sp, vg4.groupsbuname AS mould, doc.groupact(vg1.groupsbuname::character varying, vg2.groupsbuname::character varying, vg3.groupsbuname::character varying, vg4.groupsbuname::character varying, si.fpcp::character varying) AS groupact, pdt.producttype, pr.paramname AS statusname" &
                     " from basedata clt" &
                     " left join doc.document d on d.id = clt.id  LEFT JOIN vendor v ON v.shortname3 = clt.shortname" &
                     " LEFT JOIN doc.viewvendorpm vfp ON vfp.vendorcode = v.vendorcode" &
                          " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid" &
                          " LEFT JOIN masteruser mu ON mu.id = os.muid" &
                          " LEFT JOIN officerseb o ON o.ofsebid = os.parent" &
                          " LEFT JOIN masteruser mu2 ON mu2.id = o.muid" &
                          " LEFT JOIN doc.vendorgsm gsm ON gsm.vendorcode = v.vendorcode" &
                          " LEFT JOIN officerseb o1 ON o1.ofsebid = gsm.gsmid" &
                          " LEFT JOIN masteruser mu1 ON mu1.id = o1.muid" &
                     " LEFT JOIN doc.version vr ON vr.documentid = d.id LEFT JOIN doc.generalcontract gt ON gt.documentid = d.id LEFT JOIN paymentterm pt ON pt.paymenttermid = gt.paymentcode LEFT JOIN doc.supplychain sc ON sc.documentid = d.id LEFT JOIN doc.qualityappendix q ON q.documentid = d.id LEFT JOIN doc.project p ON p.documentid = d.id LEFT JOIN doc.projectspecification ps ON p.documentid = d.id LEFT JOIN doc.socialaudit sa ON sa.documentid = d.id LEFT JOIN doc.sef sef ON sef.documentid = d.id LEFT JOIN doc.sif sif ON sif.documentid = d.id LEFT JOIN doc.doctype dt ON dt.id = d.doctypeid LEFT JOIN doc.doclevel dl ON dl.id = d.doclevelid LEFT JOIN doc.docexpired de ON de.documentid = d.id " &
                     " LEFT JOIN doc.tvendorgroupsbu vg1 ON vg1.vendorcode = v.vendorcode AND vg1.groupsbuname = 'FP'::text LEFT JOIN doc.tvendorgroupsbu vg2 ON vg2.vendorcode = v.vendorcode AND vg2.groupsbuname = 'CP'::text LEFT JOIN doc.tvendorgroupsbu vg3 ON vg3.vendorcode = v.vendorcode AND vg3.groupsbuname = 'SP'::text LEFT JOIN doc.tvendorgroupsbu vg4 ON vg4.vendorcode = v.vendorcode AND vg4.groupsbuname = 'MOULD'::text LEFT JOIN doc.vendorstatus vs ON vs.vendorcode = v.vendorcode LEFT JOIN doc.paramhd ph ON ph.paramname::text = 'vendorstatus'::text LEFT JOIN doc.paramdt pr ON pr.ivalue = vs.status AND pr.paramhdid = ph.paramhdid LEFT JOIN doc.shortnameinfo si ON si.shortname = v.shortname3::text LEFT JOIN ( SELECT pd.ivalue, pd.paramname AS producttype FROM doc.paramdt pd LEFT JOIN doc.paramhd ph ON ph.paramhdid = pd.paramhdid WHERE ph.paramname::text = 'producttype'::text) pdt ON pdt.ivalue = vs.producttypeid where not clt.id isnull")

            Else
                sb.Append(String.Format(" with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname3::text as shortname" &
                                  " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                  " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where lower(u.userid) ~ '{0}$'   order by vendorname) ", HelperClass1.UserInfo.userid.ToLower))
                'viewvendorpm replace viewvendorfamilypm
                sb.Append(", basedata as (select distinct shortname,doctypeid,id from (select shortname,doctypeid, d.docdate, first_value(d.id) over (partition by doctypeid,shortname order by doctypeid,shortname,docdate desc) as id from doc.allvendordocument av" &
                      " left join doc.document d on d.id = av.id" &
                      " order by doctypeid,shortname,docdate desc)foo order by shortname,doctypeid,id)" &
                      " SELECT distinct false as download, clt.id, v.vendorcode,v.vendorname::text AS vendorname, v.shortname3::text AS shortname, mu1.username AS gsm, mu2.username AS spm, mu.username AS pm, d.id AS documentid, dt.doctypename, dl.levelname, de.expireddate, d.docdate, d.docname, d.docext, d.uploaddate, d.userid, d.remarks, vr.version, (pt.payt::text || ' - '::text) || pt.details::text AS paymentterm, sc.leadtime, sc.sasl, q.nqsu, p.projectname,ps.returnrate, sa.auditby, sa.audittype, sa.auditgrade, sef.score, sif.myyear, sif.turnovery, sif.turnovery1, sif.turnovery2, sif.turnovery3, sif.turnovery4, sif.ratioy, sif.ratioy1, sif.ratioy2, sif.ratioy3, sif.ratioy4," &
                      " CASE WHEN dt.id = 32 THEN 1 ELSE NULL::integer END AS generalcontract, CASE WHEN dt.id = 33 THEN 1 ELSE NULL::integer END AS qualityappendix, CASE WHEN dt.id = 35 THEN 1 ELSE NULL::integer END AS supplychainappendix,null::integer  AS countcontractisfull,  null::integer AS countcontractfull,  null::integer as countcontractnotfull, vg1.groupsbuname AS fp, vg2.groupsbuname AS cp, vg3.groupsbuname AS sp, vg4.groupsbuname AS mould, doc.groupact(vg1.groupsbuname::character varying, vg2.groupsbuname::character varying, vg3.groupsbuname::character varying, vg4.groupsbuname::character varying, si.fpcp::character varying) AS groupact, pdt.producttype, pr.paramname AS statusname" &
                      " from basedata clt" &
                      " left join doc.document d on d.id = clt.id  " &
                      " inner join va on va.shortname = clt.shortname" &
                      " LEFT JOIN vendor v ON v.shortname3 = va.shortname" &
                      " LEFT JOIN doc.viewvendorpm vfp ON vfp.vendorcode = v.vendorcode" &
                          " LEFT JOIN officerseb os ON os.ofsebid = vfp.pmid" &
                          " LEFT JOIN masteruser mu ON mu.id = os.muid" &
                          " LEFT JOIN officerseb o ON o.ofsebid = os.parent" &
                          " LEFT JOIN masteruser mu2 ON mu2.id = o.muid" &
                          " LEFT JOIN doc.vendorgsm gsm ON gsm.vendorcode = v.vendorcode" &
                          " LEFT JOIN officerseb o1 ON o1.ofsebid = gsm.gsmid" &
                          " LEFT JOIN masteruser mu1 ON mu1.id = o1.muid" &
                      " LEFT JOIN doc.version vr ON vr.documentid = d.id LEFT JOIN doc.generalcontract gt ON gt.documentid = d.id LEFT JOIN paymentterm pt ON pt.paymenttermid = gt.paymentcode LEFT JOIN doc.supplychain sc ON sc.documentid = d.id LEFT JOIN doc.qualityappendix q ON q.documentid = d.id LEFT JOIN doc.project p ON p.documentid = d.id LEFT JOIN doc.projectspecification ps ON p.documentid = d.id LEFT JOIN doc.socialaudit sa ON sa.documentid = d.id LEFT JOIN doc.sef sef ON sef.documentid = d.id LEFT JOIN doc.sif sif ON sif.documentid = d.id LEFT JOIN doc.doctype dt ON dt.id = d.doctypeid LEFT JOIN doc.doclevel dl ON dl.id = d.doclevelid LEFT JOIN doc.docexpired de ON de.documentid = d.id " &
                      " LEFT JOIN doc.tvendorgroupsbu vg1 ON vg1.vendorcode = v.vendorcode AND vg1.groupsbuname = 'FP'::text LEFT JOIN doc.tvendorgroupsbu vg2 ON vg2.vendorcode = v.vendorcode AND vg2.groupsbuname = 'CP'::text LEFT JOIN doc.tvendorgroupsbu vg3 ON vg3.vendorcode = v.vendorcode AND vg3.groupsbuname = 'SP'::text LEFT JOIN doc.tvendorgroupsbu vg4 ON vg4.vendorcode = v.vendorcode AND vg4.groupsbuname = 'MOULD'::text LEFT JOIN doc.vendorstatus vs ON vs.vendorcode = v.vendorcode LEFT JOIN doc.paramhd ph ON ph.paramname::text = 'vendorstatus'::text LEFT JOIN doc.paramdt pr ON pr.ivalue = vs.status AND pr.paramhdid = ph.paramhdid LEFT JOIN doc.shortnameinfo si ON si.shortname = v.shortname3::text LEFT JOIN ( SELECT pd.ivalue, pd.paramname AS producttype FROM doc.paramdt pd LEFT JOIN doc.paramhd ph ON ph.paramhdid = pd.paramhdid WHERE ph.paramname::text = 'producttype'::text) pdt ON pdt.ivalue = vs.producttypeid where not clt.id isnull")

            End If
            If TextBox9.Text <> "" Then
                'sb.Append(" and '" & TextBox9.Text & ",'" & " ~(dt.doctypename || ',')::text")
                sb.Append(" and '" & TextBox9.Text & "'" & " ~(replace(replace(dt.doctypename,'(','\('),')','\)'))")
                'sb.Append(" and dt.doctypename ~(replace(replace( '" & TextBox9.Text & "','(','\('),')','\)'))")
            End If
            'If TextBox9.Text <> "" Then
            '    'sb.Append(" and '" & TextBox9.Text & ",'" & " ~(dt.doctypename || ',')::text")
            '    sb.Append(" and '" & TextBox9.Text & "'" & " ~(replace(replace(dt.doctypename,'(','\('),')','\)'))")
            'End If
            If TextBox3.Text <> "" Then
                sb.Append(" and '" & TextBox3.Text.Replace("'", "''") & "' ~ v.shortname3::text")
            End If
            If TextBox1.Text <> "" Then
                sb.Append(" and '" & TextBox1.Text.Replace("'", "''") & "' ~ v.vendorcode::text")
            End If
            If TextBox2.Text <> "" Then
                'sb.Append(" and '" & TextBox2.Text.Replace("'", "''") & "' ~ v.vendorname::text")
                sb.Append(" and '" & TextBox2.Text & "'" & " ~(replace(replace(v.vendorname,'(','\('),')','\)'))")
                'sb.Append(" and (v.vendorname::text || ',')::text ~ '^" & TextBox2.Text.Replace("'", "''").Replace("(", "\(").Replace(")", "\)") & ",'")
                'sb.Append(String.Format(" and '{0}' ~ v.vendorname::text", TextBox2.Text.Replace("'", "''").Replace(")", "\)").Replace("(", "\(")))
            End If
        End If

        
        Dim mysaveform As New SaveFileDialog
        'sqlstr = sb.ToString


        If Not myThreadSearch.IsAlive Then            
            myThreadSearch = New Thread(AddressOf DoSearch)
            myThreadSearch.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If

    End Sub

    Private Function validreg(ByVal value As String) As String
        Return String.Format("replace(replace('^{0}$','(','\('),')','\)')", value)
    End Function


    Private Sub DoSearch()
        DSSearch = New DataSet
        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sb.ToString, DSSearch, mymessage) Then
            Try
                DSSearch.Tables(0).TableName = "Document"
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
        'If Not IsNothing(DocumentBS) Then
        ProgressReport(1, String.Format("{0} {1:#,##0} record(s) found.", "Loading Data.Done!", DocumentBS.Count))
        'Else
        'ProgressReport(1, String.Format("{0}", "Loading Data.Done!"))
        'End If

        ProgressReport(5, "Continuous")
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If Not IsNothing(DocumentBS) Then
            For Each drv As DataRowView In DocumentBS.List
                drv.Row.Item("download") = CheckBox1.Checked
            Next
            DocumentBS.EndEdit()
        End If

    End Sub


    Private Sub DocumentBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles DocumentBS.ListChanged
        CheckBox1.Enabled = Not IsNothing(DocumentBS.Current)
        DataGridView1.Enabled = Not IsNothing(DocumentBS.Current)
        Button11.Enabled = Not IsNothing(DocumentBS.Current)
    End Sub
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Me.Validate()
        If Not myThreadDownload.IsAlive Then
            'get Folder

            Dim myfolder = New SaveFileDialog
            myfolder.FileName = "SaveFile"
            If myfolder.ShowDialog = Windows.Forms.DialogResult.OK Then
                ToolStripStatusLabel1.Text = ""
                SelectedFolder = IO.Path.GetDirectoryName(myfolder.FileName)
                myThreadDownload = New Thread(AddressOf DoDownload)
                myThreadDownload.Start()
            End If

        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
    Private Sub Button11_Click1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Validate()
        If Not myThreadDownload.IsAlive Then
            'get Folder
            Dim myfolder = New FolderBrowserDialog
            myfolder.Description = "Select the folder to download the file(s)"
            If myfolder.ShowDialog = Windows.Forms.DialogResult.OK Then
                ToolStripStatusLabel1.Text = ""
                SelectedFolder = myfolder.SelectedPath
                myThreadDownload = New Thread(AddressOf DoDownload)
                myThreadDownload.Start()
            End If

        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoDownload()
        Dim i As Integer = 0
        For Each drv As DataRowView In DocumentBS.List
            If drv.Row.Item("download") Then
                i = i + 1
                ProgressReport(1, String.Format("Downloading ::{0} of {1} {2}", i, DocumentBS.Count, drv.Row.Item("docname")))
                'Dim filesource As String = String.Format("\\172.22.10.44\SharedFolder\PriceCmmf\New\documents\{0}{1}", drv.Row.Item("id"), drv.Row.Item("docext"))
                Dim filesource As String = String.Format("{2}\{0}{1}", drv.Row.Item("id"), drv.Row.Item("docext"), HelperClass1.document)
                Dim FileTarget As String = String.Format("{0}\{1}", SelectedFolder, drv.Row.Item("docname"))
                Try
                    FileIO.FileSystem.CopyFile(filesource, FileTarget, True)
                Catch ex As Exception
                    'MessageBox.Show(String.Format("File source : {0}{1}Target: {2}{1}{3}", filesource, vbCr, FileTarget, ex.Message))
                End Try

            End If
        Next
        ProgressReport(1, "Done. Please check your folder ::" & SelectedFolder)
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick

    End Sub



    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim myform As New FormDocumentDetails(DocumentBS)
        myform.Show()
    End Sub

    Private Sub FormSearchDocument_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        CheckBox4_CheckedChanged(sender, e)
    End Sub

    Private Sub FormSearchDocument_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If myThreadDownload.IsAlive Then
            If MessageBox.Show("Download still in process. Do you want to stop this process", "Download still is in process.", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                myThreadDownload.Abort()
            Else
                e.Cancel = True
            End If
        End If
    End Sub


    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click

    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        DateTimePicker1.Visible = CheckBox2.Checked
        DateTimePicker2.Visible = CheckBox2.Checked
        Label10.Visible = CheckBox2.Checked
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        TextBox4.Enabled = Not CheckBox3.Checked
        TextBox5.Enabled = Not CheckBox3.Checked
        TextBox6.Enabled = Not CheckBox3.Checked
        TextBox7.Enabled = Not CheckBox3.Checked
        TextBox8.Enabled = Not CheckBox3.Checked
        TextBox10.Enabled = Not CheckBox3.Checked
        DateTimePicker1.Enabled = Not CheckBox3.Checked
        DateTimePicker2.Enabled = Not CheckBox3.Checked
        Button4.Enabled = Not CheckBox3.Checked
        Button5.Enabled = Not CheckBox3.Checked
        Button6.Enabled = Not CheckBox3.Checked
        Button7.Enabled = Not CheckBox3.Checked
        Button8.Enabled = Not CheckBox3.Checked
        Button12.Enabled = Not CheckBox3.Checked
        CheckBox2.Enabled = Not CheckBox3.Checked
    End Sub


    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Try
            If Not IsNothing(DocumentBS.Current) Then
                Dim drv As DataRowView = DocumentBS.Current
                Dim filesource As String = String.Format("{2}\{0}{1}", drv.Row.Item("id"), drv.Row.Item("docext"), HelperClass1.document)
                If FileIO.FileSystem.GetFileInfo(filesource).Length / 1048576 < 5 Then
                    Dim mytemp = String.Format("{1}{0}", drv.Row.Item("docname"), IO.Path.GetTempPath())

                    FileIO.FileSystem.CopyFile(filesource, mytemp, True)
                    Dim p As New System.Diagnostics.Process
                    'p.StartInfo.FileName = "\\172.22.10.44\SharedFolder\PriceCMMF\New\template\Supplier Management Task User Guide-Admin.pdf"
                    p.StartInfo.FileName = mytemp
                    p.Start()
                Else
                    MessageBox.Show("File too big.Please download.")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub




    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click

    End Sub
End Class