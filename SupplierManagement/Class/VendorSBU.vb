Imports System.Text
Imports SupplierManagement.PublicClass


Public Class VendorSBU

    Private DS As DataSet
    Public Function UpdateVendorSBU01(Optional ByRef Message As String = "") As Boolean
        Dim myret As Boolean = False
        Dim myquery As New StringBuilder
        Dim BS As New BindingSource
        Dim sqlstr = "select distinct m.vendorcode, s.sbuname  from miro m" &
                     " left join pomiro pm on m.miroid = pm.miroid" &
                     " left join podtl pd on pd.podtlid = pm.podtlid" &
                     " left join materialmaster mm on mm.cmmf = pd.cmmf" &
                     " left join sbusap s on s.sbuid = mm.sbu" &
                     " where not sbu isnull" &
                     " order by vendorcode,sbuname;"
        myquery.Append(sqlstr)

        sqlstr = "select vendorcode from doc.vendorstatus;"
        myquery.Append(sqlstr)

        DS = New DataSet
        If DbAdapter1.TbgetDataSet(myquery.ToString, DS, Message) Then
            BS.DataSource = DS.Tables(0)

            Dim pk1(0) As DataColumn
            pk1(0) = DS.Tables(1).Columns("vendorcode")
            DS.Tables(1).PrimaryKey = pk1
            DS.Tables(1).TableName = "VendorStatus"

            'Vendor SBUName
            Dim mydata As New StringBuilder
            For Each drv As DataRowView In BS.List
                mydata.Append(drv.Row.Item("vendorcode").ToString & vbTab & drv.Row.Item("sbuname").ToString & vbCrLf)
            Next



            'Vendor Status

            Dim VendorGroups = From n In BS.List
                               Group By vendorcode = n.row.item("vendorcode") Into mygroup = Group


            'collect sbu
            'Using Linq to group by vendorcode
            Dim mydata2 As New StringBuilder
            For Each d In VendorGroups
                Debug.Print(d.vendorcode)
                Dim mykey(0) As DataColumn
                mykey(0) = d.vendorcode
                Dim myresult = DS.Tables(1).Rows.Find(mykey)
                If IsNothing(myresult) Then
                    mydata2.Append(d.vendorcode & vbTab & 1 & vbCrLf)
                End If

            Next


            Debug.Print(mydata.ToString)
            'delete existing,copy new one
            'copy
            If mydata.Length > 0 Then

                sqlstr = "delete from doc.vendorsbu;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, Message) Then
                    Return False
                End If

                'ProgressReport(2, String.Format("Copy Vendor"))
                sqlstr = "copy doc.vendorsbu(vendorcode,sbuname)  from stdin with null as 'Null';"
                Message = DbAdapter1.copy(sqlstr, mydata.ToString, myret)
                If Not myret Then
                    Return False
                End If

            End If
            If mydata2.Length > 0 Then

                'ProgressReport(2, String.Format("Copy Vendor"))
                sqlstr = "copy doc.vendorstatus(vendorcode,status)  from stdin with null as 'Null';"
                Message = DbAdapter1.copy(sqlstr, mydata2.ToString, myret)
                If Not myret Then
                    Return False
                End If

            End If
        End If


        Return (myret)
    End Function

    Public Function UpdateVendorSBU(Optional ByRef Message As String = "") As Boolean
        Dim myret As Boolean = False
        Dim myquery As New StringBuilder
        Dim BS As New BindingSource
        Dim BS2 As New BindingSource
        Dim BS3 As New BindingSource
        Dim sqlstr = "select distinct m.vendorcode, s.sbuname  from miro m" &
                     " left join pomiro pm on m.miroid = pm.miroid" &
                     " left join podtl pd on pd.podtlid = pm.podtlid" &
                     " left join materialmaster mm on mm.cmmf = pd.cmmf" &
                     " left join sbusap s on s.sbuid = mm.sbu" &
                     " where not sbu isnull" &
                     " order by vendorcode,sbuname;"
        myquery.Append(sqlstr)

        sqlstr = "select vendorcode from doc.vendorstatus;"
        myquery.Append(sqlstr)

        sqlstr = "WITH e AS (" &
            " SELECT DISTINCT m.vendorcode, ph.purchasinggroup" &
            " FROM miro m" &
            " LEFT JOIN pomiro pm ON pm.miroid = m.miroid" &
            " LEFT JOIN podtl pd ON pd.podtlid = pm.podtlid" &
            " LEFT JOIN pohd ph ON ph.pohd = pd.pohd" &
            " WHERE(Not ph.purchasinggroup Is NULL)" &
            ")" &
            " SELECT DISTINCT e.vendorcode, gs.groupsbuname" &
            " FROM e" &
            " LEFT JOIN purchasinggroup pg ON pg.purchasinggroup = e.purchasinggroup::bpchar" &
            " LEFT JOIN groupsbu gs ON gs.groupsbuid = pg.groupact;"
        myquery.Append(sqlstr)

        'sqlstr = "select distinct shortname::character varying from vendor" &
        '         " where not shortname isnull" &
        '         " group by shortname" &
        '         " having count(vendorcode) > 1"
        'myquery.Append(sqlstr)


        DS = New DataSet
        If DbAdapter1.TbgetDataSet(myquery.ToString, DS, Message) Then
            BS.DataSource = DS.Tables(0)
            BS2.DataSource = DS.Tables(2)
            'BS3.DataSource = DS.Tables(3)
            Dim pk1(0) As DataColumn
            pk1(0) = DS.Tables(1).Columns("vendorcode")
            DS.Tables(1).PrimaryKey = pk1
            DS.Tables(1).TableName = "VendorStatus"

            Dim VendorGroups = From n In BS.List
                               Group By vendorcode = n.row.item("vendorcode") Into mygroup = Group


            'collect sbu
            'Using Linq to group by vendorcode
            Dim mydata As New StringBuilder
            Dim mydata2 As New StringBuilder
            Dim mydata3 As New StringBuilder
            Dim mydata4 As New StringBuilder
            For Each d In VendorGroups
                'Debug.Print(d.vendorcode)
                Dim sbunamesb As New StringBuilder
                For Each e In d.mygroup
                    'Debug.Print(e.row.item("sbuname"))
                    If sbunamesb.Length > 0 Then
                        sbunamesb.Append(",")
                    End If
                    sbunamesb.Append(e.row.item("sbuname"))
                Next
                Dim mykey(0) As Object
                mykey(0) = d.vendorcode
                Dim myresult = DS.Tables(1).Rows.Find(mykey)
                If IsNothing(myresult) Then
                    mydata2.Append(d.vendorcode & vbTab & 1 & vbCrLf)
                End If

                mydata.Append(d.vendorcode & vbTab & sbunamesb.ToString & vbCrLf)
            Next

            For Each dr As DataRowView In BS2.List
                mydata3.Append(dr.Row.Item("vendorcode") & vbTab & dr.Row.Item("groupsbuname") & vbCrLf)
            Next
            'For Each dr As DataRowView In BS3.List
            '    mydata4.Append(dr.Row.Item("shortname") & vbTab & "Both" & vbCrLf)
            'Next
            'Debug.Print(mydata.ToString)

            'Vendor Status


            'collect sbu
            'Using Linq to group by vendorcode

            'For Each d In VendorGroups
            '    Debug.Print(d.vendorcode)
            '    Dim mykey(0) As DataColumn
            '    mykey(0) = d.vendorcode
            '    Dim myresult = DS.Tables(1).Rows.Find(mykey)
            '    If IsNothing(myresult) Then
            '        mydata2.Append(d.vendorcode & vbTab & 1 & vbCrLf)
            '    End If

            'Next
            'delete existing,copy new one
            'copy
            If mydata.Length > 0 Then

                sqlstr = "delete from doc.vendorsbu;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, Message) Then
                    Return False
                End If

                'ProgressReport(2, String.Format("Copy Vendor"))
                sqlstr = "copy doc.vendorsbu(vendorcode,sbuname)  from stdin with null as 'Null';"
                Message = DbAdapter1.copy(sqlstr, mydata.ToString, myret)
                If Not myret Then
                    Return False
                End If

            End If
            If mydata2.Length > 0 Then

                'ProgressReport(2, String.Format("Copy Vendor"))
                sqlstr = "copy doc.vendorstatus(vendorcode,status)  from stdin with null as 'Null';"
                Message = DbAdapter1.copy(sqlstr, mydata2.ToString, myret)
                If Not myret Then
                    Return False
                End If

            End If
            If mydata3.Length > 0 Then

                sqlstr = "delete from doc.tvendorgroupsbu;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, Message) Then
                    Return False
                End If

                'ProgressReport(2, String.Format("Copy Vendor"))
                sqlstr = "copy doc.tvendorgroupsbu(vendorcode,groupsbuname)  from stdin with null as 'Null';"
                Message = DbAdapter1.copy(sqlstr, mydata3.ToString, myret)
                If Not myret Then
                    Return False
                End If

            End If
            'If mydata4.Length > 0 Then

            '    sqlstr = "delete from doc.shortnameinfo;"
            '    Dim ra As Long
            '    If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, Message) Then
            '        Return False
            '    End If

            '    'ProgressReport(2, String.Format("Copy Vendor"))
            '    sqlstr = "copy doc.shortnameinfo(shortname,fpcp)  from stdin with null as 'Null';"
            '    Message = DbAdapter1.copy(sqlstr, mydata4.ToString, myret)
            '    If Not myret Then
            '        Return False
            '    End If

            'End If
        End If


        Return (myret)
    End Function

    Public Function UpdateShortnameInfo(Optional ByRef Message As String = "") As Boolean
        Dim myret As Boolean = False
        Dim myquery As New StringBuilder
        Dim BS As New BindingSource
        Dim sqlstr = "with e as (select distinct shortname,groupsbuname from doc.tvendorgroupsbu t" &
                     " left join vendor v on v.vendorcode = t.vendorcode" &
                     " left join doc.vendorstatus vs on vs.vendorcode = t.vendorcode" &
                     " where t.groupsbuname in ('FP','CP') and not shortname isnull and status = 1" &
                     " )" &
                     " select shortname::character varying from e" &
                     " group by shortname" &
                     " having count(groupsbuname) > 1" &
                     " order by shortname"

        myquery.Append(sqlstr)

      
        DS = New DataSet
        If DbAdapter1.TbgetDataSet(myquery.ToString, DS, Message) Then
            BS.DataSource = DS.Tables(0)

            'Shortname Info
            Dim mydata As New StringBuilder
            For Each drv As DataRowView In BS.List
                mydata.Append(drv.Row.Item("shortname").ToString & vbTab & "Both" & vbCrLf)
            Next


            If mydata.Length > 0 Then

                sqlstr = "delete from doc.shortnameinfo;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, Message) Then
                    Return False
                End If

                'ProgressReport(2, String.Format("Copy Vendor"))
                sqlstr = "copy doc.shortnameinfo(shortname,fpcp)  from stdin with null as 'Null';"
                Message = DbAdapter1.copy(sqlstr, mydata.ToString, myret)
                If Not myret Then
                    Return False
                End If

            End If

        End If


        Return (myret)
    End Function
    Public Function UpdateVendorSBU02(Optional ByRef Message As String = "") As Boolean
        Dim myret As Boolean = False
        Dim myquery As New StringBuilder
        Dim BS As New BindingSource
        Dim sqlstr = "select distinct m.vendorcode, s.sbuname  from miro m" &
                     " left join pomiro pm on m.miroid = pm.miroid" &
                     " left join podtl pd on pd.podtlid = pm.podtlid" &
                     " left join materialmaster mm on mm.cmmf = pd.cmmf" &
                     " left join sbusap s on s.sbuid = mm.sbu" &
                     " where not sbu isnull" &
                     " order by vendorcode,sbuname;"
        myquery.Append(sqlstr)

        sqlstr = "select vendorcode from doc.vendorstatus;"
        myquery.Append(sqlstr)

        DS = New DataSet
        If DbAdapter1.TbgetDataSet(myquery.ToString, DS, Message) Then
            BS.DataSource = DS.Tables(0)

            Dim pk1(0) As DataColumn
            pk1(0) = DS.Tables(1).Columns("vendorcode")
            DS.Tables(1).PrimaryKey = pk1
            DS.Tables(1).TableName = "VendorStatus"

            Dim VendorGroups = From n In BS.List
                               Group By vendorcode = n.row.item("vendorcode") Into mygroup = Group


            'collect sbu
            'Using Linq to group by vendorcode
            Dim mydata As New StringBuilder
            Dim mydata2 As New StringBuilder
            For Each d In VendorGroups
                'Debug.Print(d.vendorcode)
                Dim sbunamesb As New StringBuilder
                For Each e In d.mygroup
                    'Debug.Print(e.row.item("sbuname"))
                    If sbunamesb.Length > 0 Then
                        sbunamesb.Append(",")
                    End If
                    sbunamesb.Append(e.row.item("sbuname"))
                Next
                Dim mykey(0) As Object
                mykey(0) = d.vendorcode
                Dim myresult = DS.Tables(1).Rows.Find(mykey)
                If IsNothing(myresult) Then
                    mydata2.Append(d.vendorcode & vbTab & 1 & vbCrLf)
                End If

                mydata.Append(d.vendorcode & vbTab & sbunamesb.ToString & vbCrLf)
            Next
            'Debug.Print(mydata.ToString)

            'Vendor Status


            'collect sbu
            'Using Linq to group by vendorcode

            'For Each d In VendorGroups
            '    Debug.Print(d.vendorcode)
            '    Dim mykey(0) As DataColumn
            '    mykey(0) = d.vendorcode
            '    Dim myresult = DS.Tables(1).Rows.Find(mykey)
            '    If IsNothing(myresult) Then
            '        mydata2.Append(d.vendorcode & vbTab & 1 & vbCrLf)
            '    End If

            'Next
            'delete existing,copy new one
            'copy
            If mydata.Length > 0 Then

                sqlstr = "delete from doc.vendorsbu;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, Message) Then
                    Return False
                End If

                'ProgressReport(2, String.Format("Copy Vendor"))
                sqlstr = "copy doc.vendorsbu(vendorcode,sbuname)  from stdin with null as 'Null';"
                Message = DbAdapter1.copy(sqlstr, mydata.ToString, myret)
                If Not myret Then
                    Return False
                End If

            End If
            If mydata2.Length > 0 Then

                'ProgressReport(2, String.Format("Copy Vendor"))
                sqlstr = "copy doc.vendorstatus(vendorcode,status)  from stdin with null as 'Null';"
                Message = DbAdapter1.copy(sqlstr, mydata2.ToString, myret)
                If Not myret Then
                    Return False
                End If

            End If
        End If


        Return (myret)
    End Function
End Class
