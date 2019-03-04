Imports SupplierManagement.SharedClass
Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Public Class ClassSupplierDashBoard

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)

    Private sb As New StringBuilder
    Private DS As DataSet
    Private mymessage As String = String.Empty
    Private BSShortNameHelper As BindingSource
    Private BSVendorHelper As BindingSource

    Private shortName As String
    Private MyForm As Object
    Public Sub New(ByVal MyForm As Object)

        Me.MyForm = MyForm
        PopulateData()
  
    End Sub

    Private Sub PopulateData()
        If Not myThread.IsAlive Then

            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub



    Sub DoWork()
        If HelperClass1.UserInfo.IsAdmin Then
            sb.Append("select null::text as shortname union all (select distinct shortname::text from vendor  where not shortname isnull order by shortname);")
            sb.Append("select null as vendorcode,''::text as description,''::text as vendorname,null::text as shortname union all (select vendorcode, vendorcode::text || ' - ' || vendorname::text as description,vendorname::text,shortname::text  from vendor order by vendorname);")

        Else
            sb.Append(String.Format("select null::text as shortname union all  (select distinct shortname::text from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid" &
                 " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{0}$' and not shortname isnull order by shortname);", HelperClass1.UserInfo.userid))
            sb.Append(String.Format("select null as vendorcode,''::text as description,''::text as vendorname,null::text as shortname union all (select distinct v.vendorcode, v.vendorcode::text || ' - ' || v.vendorname::text as description,v.vendorname::text,shortname::text  from doc.groupvendor gv left join vendor  v on v.vendorcode = gv.vendorcode left join doc.groupauth g on g.groupid = gv.groupid" &
                      " left join doc.groupuser gu on gu.groupid = gv.groupid left join doc.user u on u.id = gu.userid where u.userid ~ '{0}$' and not shortname isnull  order by vendorname);", HelperClass1.UserInfo.userid))

        End If
        DS = New DataSet
        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                DS.Tables(0).TableName = "ShortName"
                DS.Tables(1).TableName = "Vendorcode"

                ProgressReport(4, "InitData")
            Catch ex As Exception
                ProgressReport(1, "UCSupplierDashboard Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
        End If
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        Try
            Select Case id
                Case 1
                Case 2
                Case 4
                    BSShortNameHelper = New BindingSource
                    BSVendorHelper = New BindingSource

                    BSShortNameHelper.DataSource = DS.Tables(0)
                    BSVendorHelper.DataSource = DS.Tables(1)

                    MyForm.BSshortNameHelper = BSShortNameHelper
                    MyForm.BSvendorhelper = BSVendorHelper

                Case 5
                    Dim bs As New BindingSource
                    bs.DataSource = DS.Tables(0)
                    Dim drv = bs.Current

                    'TextBox1.Text = drv.item("vendorcode")
                    'TextBox2.Text = drv.item("vendorname")
                Case 6
                    'ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    'Try
                    '    DocumentBS = New BindingSource
                    '    DocumentBS.DataSource = DSSearch.Tables("Document")
                    '    DataGridView1.AutoGenerateColumns = False
                    '    DataGridView1.DataSource = DocumentBS

                    'Catch ex As Exception
                    '    message = ex.Message
                    'End Try
            End Select
        Catch ex As Exception

        End Try
        'End If

    End Sub


End Class
