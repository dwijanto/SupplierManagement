Imports System.Text

Public Class SupplierGSMAdapter
    Implements IAdapter
    Dim DbAdapter1 As DbAdapter = DbAdapter.getInstance
    Public DS As DataSet
    Public BS As BindingSource
    Public BS02 As BindingSource
    Public BS03 As BindingSource
    Public ReadOnly Property GetTableGSM As DataTable
        Get
            Return DS.Tables("GSM").Copy()
        End Get
    End Property

    Public ReadOnly Property GetTableVendor As DataTable
        Get
            Return DS.Tables("VENDOR").Copy()
        End Get
    End Property

    Public Function getGSMBS() As BindingSource
        Dim sqlstr As String = String.Format("select o.ofsebid as gsmid,mu.username as gsm from officerseb o" &
                  " left join masteruser mu on mu.id = o.muid" &
                  " left join teamtitle tt on tt.teamtitleid = o.teamtitleid" &
                  " where teamtitleshortname = 'GSM' and mu.isactive order by mu.username;")
        Dim DS As New DataSet
        Dim bs As New BindingSource
        If DbAdapter1.TbgetDataSet(sqlstr, DS) Then
            bs.DataSource = DS.Tables(0)
        End If
        Return bs
    End Function

    Public Function loaddata() As Boolean Implements IAdapter.loaddata
        Dim myret As Boolean = False
        Dim sb As New StringBuilder
        sb.Append("select vg.id, vg.vendorcode::text,vg.gsmid,vg.effectivedate,v.vendorname::text,v.shortname3::text as shortname,mu.username as gsm from doc.vendorgsm vg" &
                  " left join officerseb o on o.ofsebid = vg.gsmid" &
                  " left join masteruser mu on mu.id = o.muid" &
                  " left join vendor v on v.vendorcode = vg.vendorcode;")
        sb.Append("select o.ofsebid as gsmid,mu.username as gsm from officerseb o" &
                  " left join masteruser mu on mu.id = o.muid" &
                  " left join teamtitle tt on tt.teamtitleid = o.teamtitleid" &
                  " where teamtitleshortname = 'GSM' and mu.isactive;")
        sb.Append("select vendorcode,vendorname::text ,vendorcode::text || ' - ' || vendorname::text as description , shortname3::text as shortname from vendor order by vendorcode;")

        Dim sqlstr = sb.ToString
        DS = New DataSet
        If DbAdapter1.TbgetDataSet(sqlstr, DS) Then
            'set primary key
            DS.Tables(0).TableName = "VENDOR GSM"
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).PrimaryKey = pk
            DS.Tables(0).Columns("id").AutoIncrement = True
            DS.Tables(0).Columns("id").AutoIncrementSeed = -1
            DS.Tables(0).Columns("id").AutoIncrementStep = -1

           
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)

            DS.Tables(1).TableName = "GSM"
            BS02 = New BindingSource
            BS02.DataSource = DS.Tables(1)

            DS.Tables(2).TableName = "VENDOR"
            BS03 = New BindingSource
            BS03.DataSource = DS.Tables(2)
            myret = True

        End If
        Return myret  
    End Function

    Public Function GetNewRecord() As DataRowView
        Dim drv As DataRowView = BS.AddNew                
        Return drv
    End Function
    Public Function GetCurrentRecord() As DataRowView
        Return BS.Current
    End Function

    Public Sub RemoveAt(ByVal value As Integer)
        BS.RemoveAt(value)
    End Sub
    Public Function Save(ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False
        Dim myVendorGSMModel As New VendorGSMModel
        If myVendorGSMModel.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function

    Public Property ApplyFilter As String
        Get
            Return BS.Filter.ToString
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format("[vendorcode] like '*{0}*' or [vendorname] like '*{0}*' or [shortname] like '*{0}*' or [gsm] like '*{0}*'", value)
        End Set
    End Property

    Public Function save() As Boolean Implements IAdapter.save
        Dim myret As Boolean = False
        BS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If

        Return myret
    End Function

    
    

End Class
