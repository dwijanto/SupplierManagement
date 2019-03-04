Imports System.Threading
Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class ModelPriceHistory    
    Delegate Sub myfunct(ByRef obj As Object, ByRef e As EventArgs)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Public f As myfunct
    Public myform As Object

    Public Property criteriaType As String
    Public Property criteriaCMMF As String
    Public Property sqlstr As String
    Dim DT As New DataTable
    Dim sb As New StringBuilder
    Dim drv As DataRowView

    Public Sub New(ByVal myform As Object, ByVal bs As BindingSource, ByVal vendorquery As FormSupplierDashboard.VendorQuery, ByVal vendorcode As String, ByVal shortname As String)
        drv = bs.Current
        If vendorquery = FormSupplierDashboard.VendorQuery.shortname Then
            criteriaType = String.Format(" v.shortname = '{0}'", shortname)

        Else
            criteriaType = String.Format(" v.vendorcode::text = '{0}'", vendorcode)
        End If
        Me.myform = myform
    End Sub

    Public Sub loadData()
        If Not myThread.IsAlive Then
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        DT = New DataTable
        Dim mymessage As String = String.Empty
        sb.Clear()
        sb.Append(String.Format("with std as (select cmmf, max(validfrom) as validfrom from standardcostad group by cmmf)" &
                                " select pdt.validon,pdt.price / pdt.pricingunit as unitprice, phd.negotiateddate,r.reasonname,phd.description,pdt.comment,pdt.price as price,pdt.pricingunit,pdt.purchorg,pdt.plant,phd.pricechangehdid," &
                                " (getpriceinfo(getstatusname(phd.status),pdt.cmmf,pdt.vendorcode,pdt.validon,pdt.price,pdt.pricingunit,ad.planprice1,ad.per)).*,ad.planprice1," &
                                " case when ag.enddate > current_date then ag.value else null end as amort from pricechangedtl pdt left join pricechangehd phd on phd.pricechangehdid = pdt.pricechangehdid" &
                                " left join vendor v on v.vendorcode = pdt.vendorcode left join pricechangereason r on r.id = phd.reasonid left join agreementcmmf ac on ac.material = pdt.cmmf left join agvalue ag on ag.agreement = ac.agreement and ag.vendorcode = v.vendorcode" &
                                " left join std on std.cmmf = pdt.cmmf " &
                                " left join standardcostad ad on ad.cmmf =std.cmmf and ad.validfrom = std.validfrom" &
                                " and {0}" &
                                " Where pdt.cmmf = {1} order by validon desc", criteriaType, drv.Item("cmmf")))
        Try
            If DbAdapter1.TbgetDataTable(sb.ToString, DT, mymessage) Then
                myform.invoke(f, New Object() {DT, New EventArgs})
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
