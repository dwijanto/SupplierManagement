Imports System.DirectoryServices
Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class FormUser
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents USERBS As BindingSource
    Dim UserGroupVendorBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim myListArray As String() = {"userid", "username", "email"}
    Dim RBS As BindingSource

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not TextBox1.ReadOnly Then
            ToolStripStatusLabel1.Text = "Checking userid. please wait..."
            Application.DoEvents()
            Dim myresult() As String = Nothing
            Try
                HelperClass1.getDataAD(TextBox1.Text)

                'If Not myresult Is Nothing Then
                TextBox1.Text = TextBox1.Text.ToLower
                TextBox2.Text = HelperClass1.UserInfo.DisplayName
                TextBox3.Text = HelperClass1.UserInfo.email
                ToolStripStatusLabel1.Text = "Checking userid. Done."
                DataGridView1.Invalidate()
                Dim drv As DataRowView = USERBS.Current
                drv.Row.RowError = ""
                'Else
                'TextBox2.Text = ""
                'TextBox3.Text = ""
                'ToolStripStatusLabel1.Text = "User not found!."
                'End If
            Catch ex As Exception
                ToolStripStatusLabel1.Text = "Checking userid. User not found."
            End Try


            Application.DoEvents()
        End If
    End Sub
    'Private Sub Button1_Click1(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If Not TextBox1.ReadOnly Then
    '        ToolStripStatusLabel1.Text = "Checking userid. please wait..."
    '        Application.DoEvents()
    '        Dim myresult() As String = Nothing
    '        myresult = getDataAD(TextBox1.Text)
    '        If Not myresult Is Nothing Then
    '            TextBox1.Text = TextBox1.Text.ToLower
    '            TextBox2.Text = If(myresult(1), TextBox2.Text)
    '            TextBox3.Text = If(myresult(0), TextBox3.Text)
    '            ToolStripStatusLabel1.Text = "Checking userid. Done."
    '            DataGridView1.Invalidate()
    '            Dim drv As DataRowView = USERBS.Current
    '            drv.Row.RowError = ""
    '        Else
    '            TextBox2.Text = ""
    '            TextBox3.Text = ""
    '            ToolStripStatusLabel1.Text = "User not found!."
    '        End If

    '        Application.DoEvents()
    '    End If
    'End Sub
    'Private Function getDataAD(ByRef userid As String) As String()
    '    Dim myresult() As String = Nothing
    '    Dim entry As DirectoryEntry = New DirectoryEntry
    '    Dim myuser() As String = userid.Split("\")
    '    Select Case myuser(0)
    '        Case "as"
    '            entry.Path = "LDAP://as/DC=as;DC=seb;DC=com"
    '        Case "eu"
    '            entry.Path = "LDAP://eu/DC=eu;DC=seb;DC=com"
    '        Case "supor"
    '            entry.Path = "LDAP://supor/DC=supor;DC=seb;DC=com"
    '        Case "sa"
    '            entry.Path = "LDAP://sa/DC=sa;DC=seb;DC=com"
    '        Case "na"
    '            entry.Path = "LDAP://na/DC=na;DC=seb;DC=com"
    '    End Select


    '    Try
    '        Dim mysearcher As DirectorySearcher = New DirectorySearcher(entry)
    '        mysearcher.Filter = "sAMAccountName=" & myuser(1)
    '        Dim result As SearchResult = mysearcher.FindOne
    '        If Not (result Is Nothing) Then
    '            ReDim myuser(2)
    '            Dim myPerson As DirectoryEntry = New DirectoryEntry
    '            myPerson.Path = result.Path
    '            myuser(0) = myPerson.Properties("mail").Value.ToString
    '            Try
    '                myuser(1) = UCase(myPerson.Properties("givenname").Value.ToString & " " & myPerson.Properties("sn").Value.ToString)
    '            Catch
    '            End Try
    '            myresult = myuser
    '        End If
    '    Catch ex As Exception
    '        myuser = Nothing
    '    End Try

    '    Return myresult
    'End Function

    Private Sub FormUser_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        USERBS.EndEdit()
        Dim abc = DS.GetChanges
        If Not IsNothing(abc) Then
            If abc.Tables(0).Rows.Count <> 0 Then
                Select Case MessageBox.Show("Save unsaved records?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                    Case Windows.Forms.DialogResult.Yes
                        If Me.validate Then
                            ToolStripButton4.PerformClick()
                        Else
                            e.Cancel = True
                        End If

                    Case Windows.Forms.DialogResult.Cancel
                        e.Cancel = True
                End Select
            End If

        End If

    End Sub



    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        sb.Append("select u.userid,username,u.email,isadmin,isfinance,allowupdatedocument,isguest,id,o.teamtitleid,o.ofsebid,m.newemail,u.isactive from doc.user u" &
                  " left join officerseb o on o.userid = u.userid " &
                  " left join doc.emailmapping m on m.oldemail = u.email order by id;")
        sb.Append("select u.username,g.groupname,v.vendorcode::text,v.vendorname::text,v.shortname::text from doc.groupuser gu" &
                  " left join doc.groupvendor gv on gv.groupid = gu.groupid" &
                  " left join doc.groupauth g on g.groupid = gu.groupid" &
                  " left join vendor v on v.vendorcode = gv.vendorcode" &
                  " left join doc.user u on u.id = gu.userid" &
                  " order by username,groupname,vendorname;")
        sb.Append("select null as userid,null as username,null as email,null as isadmin,null as isfinance,null as allowupdate,null as id,null as teamtitleid,null as ofsebid union all (select u.userid,u.username,u.email,isadmin,isfinance,allowupdatedocument,id,o.teamtitleid,o.ofsebid from doc.user u left join officerseb o on o.userid = u.userid where o.teamtitleid in (8,9)order by teamtitleid,username)")


        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "User"
                DS.Tables(1).TableName = "UserGroupVendor"
                DS.Tables(2).TableName = "Replacement"
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
                        Try
                            USERBS = New BindingSource
                            'UserGroupVendorBS = New BindingSource
                            RBS = New BindingSource
                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("id")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns("id").AutoIncrement = True
                            DS.Tables(0).Columns("id").AutoIncrementSeed = 0
                            DS.Tables(0).Columns("id").AutoIncrementStep = -1
                            DS.Tables(0).TableName = "User"

                            USERBS.DataSource = DS.Tables(0)
                            RBS.DataSource = DS.Tables(2)
                            'UserGroupVendorBS.DataSource = DS.Tables(1)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = USERBS
                            DataGridView1.RowTemplate.Height = 22

                            TextBox1.DataBindings.Clear()
                            TextBox2.DataBindings.Clear()
                            TextBox3.DataBindings.Clear()
                            TextBox4.DataBindings.Clear()
                            CheckBox1.DataBindings.Clear()
                            CheckBox2.DataBindings.Clear()
                            CheckBox3.DataBindings.Clear()
                            CheckBox4.DataBindings.Clear()
                            CheckBox5.DataBindings.Clear()

                            TextBox1.DataBindings.Add(New Binding("Text", USERBS, "userid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox2.DataBindings.Add(New Binding("Text", USERBS, "username", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox3.DataBindings.Add(New Binding("Text", USERBS, "email", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox4.DataBindings.Add(New Binding("Text", USERBS, "newemail", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            CheckBox1.DataBindings.Add(New Binding("checked", USERBS, "isadmin", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            CheckBox3.DataBindings.Add(New Binding("checked", USERBS, "isfinance", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            CheckBox2.DataBindings.Add(New Binding("checked", USERBS, "allowupdatedocument", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            CheckBox4.DataBindings.Add(New Binding("checked", USERBS, "isactive", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            CheckBox5.DataBindings.Add(New Binding("checked", USERBS, "isguest", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            ComboBox1.DataSource = RBS
                            ComboBox1.ValueMember = "userid"
                            ComboBox1.DisplayMember = "username"
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
    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles USERBS.ListChanged
        TextBox1.Enabled = Not IsNothing(USERBS.Current)
        


    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = USERBS.AddNew()
        drv.Row.Item("isadmin") = False
        drv.Row.Item("isfinance") = False
        drv.Row.Item("isguest") = False
        CheckBox1.Checked = False
        drv.Row.Item("allowupdatedocument") = False
        CheckBox2.Checked = False
    End Sub
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        USERBS.EndEdit()
        If Me.validate Then
            Try
                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                Dim ds2 As DataSet
                ds2 = DS.GetChanges

                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)

                    If Not DbAdapter1.UserTx(Me, mye) Then
                        MessageBox.Show(mye.message)
                        DS.Merge(ds2)
                        Exit Sub
                    End If
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    DataGridView1.Invalidate()
                    MessageBox.Show("Saved.")

                   
                End If

                'Replace
                If ComboBox1.Text <> "" Then
                    Dim drv As DataRowView = RBS.Current
                    Dim drv2 As DataRowView = USERBS.Current
                    Dim title As String
                    Dim title2 As String
                    Dim useridold As String = drv2.Item("userid")
                    Dim useridnew As String = drv.Item("userid")
                    Dim ofsebidold As String = drv2.Item("ofsebid")
                    Dim ofsebidnew As String = drv.Item("ofsebid")
                    Dim effectivedate As String = String.Empty
                    Dim sb As New StringBuilder
                    If drv.Item("teamtitleid") = 8 Then 'SPM
                        title = "ssmidpl"
                        title2 = "ssmid"
                        effectivedate = "ssmeffectivedate"
                    Else 'PM
                        title = "pmid"
                        title2 = "pmid"
                        effectivedate = "pmeffectivedate"
                    End If
                    sb.Append(String.Format("insert into vendorspmpm(vendorcode,{0},{1},userid) (select vendorcode,{0},{1},'{3}' from vendor where {0} = '{2}');", title, effectivedate, ofsebidold, HelperClass1.UserInfo.userid))
                    sb.Append(String.Format("Update vendor set {0} = {1},{2} = '{3:yyyy-MM-dd}' where {0} = {4};", title, ofsebidnew, effectivedate, DateTimePicker1.Value.Date, ofsebidold))
                    sb.ToString()
                    DbAdapter1.ExecuteNonQuery(sb.ToString)
                    '1. Insert VendorSPMPM with (select vendor where userid = userid) 
                    '2. Update Vendor set "SPM/PM" = new userid, effectivedate = inputdate where userid = userid
                End If


            Catch ex As Exception
                MessageBox.Show(" Error:: " & ex.Message)
            End Try
        End If
        DataGridView1.Invalidate()
    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()

        For Each drv As DataRowView In USERBS.List
            If drv.Row.RowState = DataRowState.Modified Or drv.Row.RowState = DataRowState.Added Then
                If Not validaterow(drv) Then
                    myret = False
                End If
            End If
        Next
        If Not myret Then
            If MessageBox.Show("bypass the error?", "Bypass Error", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                myret = True
                'For Each drv As DataRowView In USERBS.List
                '    'drv.Row.RowError = ""
                'Next
            End If
        End If

        Return myret
    End Function

    Private Function validaterow(ByVal drv As DataRowView) As Boolean
        Dim myret As Boolean = True
        Dim sb As New StringBuilder
        If IsDBNull(drv.Row.Item("userid")) Then
            myret = False
            sb.Append("UserId cannot be blank. ")
        End If
        If IsDBNull(drv.Row.Item("username")) Then
            myret = False
            sb.Append("Username cannot be blank. ")
        End If
        If IsDBNull(drv.Row.Item("email")) Then
            myret = False
            sb.Append("email cannot be blank. ")
        End If

        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(USERBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    USERBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        USERBS.CancelEdit()
    End Sub
    Private Sub ToolStripComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        runfilter()
    End Sub
    Private Sub ToolStripTextBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        runfilter()
    End Sub
    Private Sub runfilter()
        USERBS.Filter = ""
        If ToolStripTextBox1.Text <> "" And ToolStripComboBox1.SelectedIndex <> -1 Then
            USERBS.Filter = myListArray(ToolStripComboBox1.SelectedIndex) & " like '*" & ToolStripTextBox1.Text.Replace("'", "''") & "*'"
        End If
    End Sub



    Private Sub ToolStripButton2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim username As String = DirectCast(USERBS.Current, DataRowView).Row.Item("username").ToString
        Dim DT As DataTable = DS.Tables("UserGroupVendor").Copy()
        UserGroupVendorBS = New BindingSource
        UserGroupVendorBS.DataSource = DT
        Dim myform = New FormUserVendorGroup(UserGroupVendorBS, username)
        myform.Show()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ToolStripButton2.PerformClick()
    End Sub

  
    Private Sub USERBS_PositionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles USERBS.PositionChanged
        If Not IsNothing(USERBS.Current) Then
            Dim drv As DataRowView = USERBS.Current
            Dim check As Boolean = False

            If Not IsDBNull(drv.Item("teamtitleid")) Then
                check = (drv.Item("teamtitleid") = 8 Or drv.Item("teamtitleid") = 9)
            End If

            Label6.Visible = check
            ComboBox1.Visible = check
            DateTimePicker1.Visible = check
            If check Then
                RBS.Position = 0
            End If
            Label8.Visible = Not IsDBNull(drv.Item("isactive"))
            CheckBox4.Visible = Not IsDBNull(drv.Item("isactive"))
        End If
    End Sub
End Class