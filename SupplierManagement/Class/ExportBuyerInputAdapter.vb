Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass


Public Class ExportBuyerInputAdapter
    Private sb As StringBuilder
    Public DS As DataSet
    Public BS As BindingSource
    Public ActionPlanBS As BindingSource
    Public Property shortname
    Public Sub New(ByVal shortname)
        Me.shortname = shortname
    End Sub

    Public Function LoadData() As Boolean
        DS = New DataSet
        BS = New BindingSource
        sb = New StringBuilder
        ActionPlanBS = New BindingSource

        sb.Clear()
        'sb.Append(String.Format("with ld as (" &
        '          " select distinct first_value(id) over(partition by producttype order by lastvisit desc) as id," &
        '          " first_value(lastvisit) over(partition by producttype order by lastvisit desc) as lastvisit," &
        '          " producttype from doc.buyerinput where shortname = '{0}' order by lastvisit desc)" &
        '          " select  * from ld" &
        '          " left join doc.buyerinput bi on bi.id = ld.id;", shortname))
        sb.Append(String.Format("with ld as (select distinct by.producttype,by.documentid,d.docdate,v.shortname::text,vd.headerid from doc.buyerinput by" &
                                " left join doc.document d on d.id = by.documentid" &
                                " left join doc.vendordoc vd on vd.documentid = d.id" &
                                " left join vendor v on v.vendorcode = vd.vendorcode " &
                                " where shortname = '{0}' order by docdate desc) select * from ld left join doc.buyerinput bi on bi.documentid = ld.documentid;", shortname))
        sb.Append(String.Format("with ld as (select distinct by.producttype,by.documentid,d.docdate,v.shortname::text,vd.headerid from doc.buyerinput by" &
                                " left join doc.document d on d.id = by.documentid" &
                                " left join doc.vendordoc vd on vd.documentid = d.id" &
                                " left join vendor v on v.vendorcode = vd.vendorcode " &
                                " where shortname = '{0}' order by docdate desc) " &
                                " select ld.headerid,ac.* from ld 	left join doc.vendordoc vd on vd.headerid = ld.headerid " &
                                " inner join doc.actionplan ac on ac.documentid = vd.documentid;", shortname))

        If DbAdapter1.TbgetDataSet(sb.ToString, DS) Then
            Dim pk(0) As DataColumn
            DS.Tables(0).PrimaryKey = pk
            BS.DataSource = DS.Tables(0)
            ActionPlanBS.DataSource = DS.Tables(1)
        End If
        Return True
    End Function

    Public Function GetActionPlan(ByVal criteria As String) As Boolean
        Dim myDS As New DataSet
        Dim mysb As New StringBuilder
        mysb.Append(String.Format("with vr as (select distinct first_value(id) over (partition by shortname,actionid order by id desc) as id from doc.actionplan where trim(shortname) = '{0}')", shortname))
        mysb.Append(String.Format("(select null::text as spacer,ac.id,ac.documentid,ac.priority,ac.situation,ac.target,ac.proposal,ac.responsibleperson,ac.startdate,ac.enddate,ac.result,ac.finishdate,ac.status,ac.actionid from doc.actionplan ac inner join vr on vr.id = ac.id where status <> 'Closed' order by id asc)"))
        If criteria.Length > 0 Then
            mysb.Append(String.Format(" union all  select null::text as spacer,ac.id,ac.documentid,ac.priority,ac.situation,ac.target,ac.proposal,ac.responsibleperson,ac.startdate,ac.enddate,ac.result,ac.finishdate,ac.status,ac.actionid from doc.actionplan ac inner join vr on vr.id = ac.id where {0} ", criteria))
        End If

        mysb.Append(";")
        ActionPlanBS = New BindingSource
        If DbAdapter1.TbgetDataSet(mysb.ToString, myDS) Then
            ActionPlanBS.DataSource = myDS.Tables(0)
        End If
        Return True
    End Function

End Class
