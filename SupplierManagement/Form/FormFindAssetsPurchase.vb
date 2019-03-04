Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class FormFindAssetsPurchase
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents ToolingListBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder

    Dim myarray() = {"", "lower(ap.assetpurchaseid)", "lower(projectcode)", "lower(projectname)", "applicantdate::text", "lower(applicantname)", "lower(vendorname)", "lower(shortname)", "lower(aeb)", "lower(investmentorderno)", "lower(toolingpono)", "lower(financeassetno)"}
    Dim myId As Long = -1
    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        ProgressReport(7, "InitFilter")
        
        Dim sqlstr = String.Format(" select distinct ap.id, ap.assetpurchaseid,tp.projectcode,tp.projectname,ap.applicantdate::text,v.vendorname::text,v.shortname::text,ap.aeb,ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate" &
                  " from  doc.toolingproject tp" &
                  " left join doc.assetpurchase ap on ap.projectid =  tp.id" &
                  " left join vendor v on v.vendorcode = ap.vendorcode" &
                  " where not ap.id isnull {0} order by ap.id desc;", sb.ToString.ToLower)
        If Not HelperClass1.UserInfo.IsAdmin Then
            sqlstr = String.Format(
                    " with va as (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text" &
                                   " from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid " &
                                   " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{1}$'   order by vendorname) " &
                  " select distinct ap.id, ap.assetpurchaseid,tp.projectcode,tp.projectname,ap.applicantdate::text,v.vendorname::text,v.shortname::text,ap.aeb,ap.investmentorderno,ap.toolingpono,ap.financeassetno,ap.sapcapdate" &
                  " from  doc.toolingproject tp" &
                  " left join doc.assetpurchase ap on ap.projectid =  tp.id" &
                  " inner join va on va.vendorcode = ap.vendorcode" &
                  " left join vendor v on v.vendorcode = va.vendorcode" &
                  " where not ap.id isnull {0} order by ap.id desc;", sb.ToString.ToLower, HelperClass1.UserInfo.userid)
        End If
        '" left join doc.toolinglist tl on tl.assetpurchaseid = ap.id" &
        '" left join doc.toolingpayment tpay on tpay.toolinglistid = tl.id" &
        '" left join doc.toolinginvoice ti on ti.id = tpay.invoiceid" &
        '" where not ap.id isnull {0} order by ap.id desc;", sb.ToString)


        If DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "ToolingListDT"

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
                            ToolingListBS = New BindingSource
                            DS.Tables(0).TableName = "ToolingList"

                            ToolingListBS.DataSource = DS.Tables(0)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = ToolingListBS
                            DataGridView1.RowTemplate.Height = 22

                            If myId <> -1 Then
                                Dim myposition = ToolingListBS.Find("id", myId)
                                ToolingListBS.Position = myposition
                            End If

                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                    Case 7
                        If ComboBox1.SelectedIndex > 0 And TextBox1.Text <> "" Then
                            sb.Append(String.Format("and {0} like '%{1}%'", myarray(ComboBox1.SelectedIndex), TextBox1.Text))
                        End If
                        If ComboBox2.SelectedIndex > 0 And TextBox2.Text <> "" Then
                            sb.Append(String.Format("and {0} like '%{1}%'", myarray(ComboBox2.SelectedIndex), TextBox2.Text))
                        End If
                        If ComboBox3.SelectedIndex > 0 And TextBox3.Text <> "" Then
                            sb.Append(String.Format("and {0} like '%{1}%'", myarray(ComboBox3.SelectedIndex), TextBox3.Text))
                        End If
                        If CheckBox1.Checked Then
                            sb.Append(String.Format("and ap.sapcapdate >= '{0:yyyy-MM-dd}' and ap.sapcapdate <= '{1:yyyy-MM-dd}'", DateTimePicker2.Value.Date, DateTimePicker3.Value.Date))
                        End If
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loaddata()
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Not IsNothing(ToolingListBS.Current) Then
            Dim drv = ToolingListBS.Current
            myId = drv.row.item("id")
            Dim myformshow As New FormAssetsPurchase(myid)
            myformshow.ShowDialog()
            Me.Button1.PerformClick()
           



        End If

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ComboBox1.SelectedIndex = 7
        ComboBox2.SelectedIndex = 3
        ComboBox3.SelectedIndex = 5
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        DateTimePicker2.Enabled = CheckBox1.Checked
        DateTimePicker3.Enabled = CheckBox1.Checked
    End Sub


End Class