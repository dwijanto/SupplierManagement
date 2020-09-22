Public Class VendorInformationModificationAdapter
    Implements IAdapter
    Implements IToolbarAction

    Dim DS As DataSet
    Dim BS As BindingSource
    Private DTLBS As New BindingSource
    Dim VCBS As BindingSource
    Dim CBS As BindingSource
    Dim VBS As BindingSource
    Public VendorInfoModiActionBS As BindingSource

    Dim DocumentBS As BindingSource
    Dim DocTypeBS As BindingSource

    Public myModel As VendorInformationModificationModel

    Private FileSourceFullPath As String

    Public Function loaddata(ByVal myid As Long, ByVal vendorcode As Long) As Boolean
        Dim myret As Boolean = False
        myModel = New VendorInformationModificationModel(vendorcode)
        DS = New DataSet
        If myModel.LoadData(DS, myid) Then
            'DS.Tables(0).TableName = "HEADER"
            'DS.Tables(1).TableName = "DTL"


            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).PrimaryKey = pk
            DS.Tables(0).Columns("id").AutoIncrement = True
            DS.Tables(0).Columns("id").AutoIncrementSeed = -2
            DS.Tables(0).Columns("id").AutoIncrementStep = -1
            DS.Tables(0).TableName = "Header"


            Dim pk1(0) As DataColumn
            pk1(0) = DS.Tables(1).Columns("id")
            DS.Tables(1).PrimaryKey = pk1
            DS.Tables(1).Columns("id").AutoIncrement = True
            DS.Tables(1).Columns("id").AutoIncrementSeed = -1
            DS.Tables(1).Columns("id").AutoIncrementStep = -1
            DS.Tables(1).TableName = "Detail"
            Dim fieldIdUnique As UniqueConstraint = New UniqueConstraint(New DataColumn() {DS.Tables(1).Columns("fieldid"), DS.Tables(1).Columns("hdid")})
            DS.Tables(1).Constraints.Add(fieldIdUnique)


            Dim pk3(0) As DataColumn
            pk3(0) = DS.Tables(3).Columns("vendorcode")
            pk3(0) = DS.Tables(3).Columns("contactid")
            DS.Tables(3).PrimaryKey = pk3
 

            Dim pk4(0) As DataColumn
            pk4(0) = DS.Tables(4).Columns("id")
            DS.Tables(4).PrimaryKey = pk4
            DS.Tables(4).Columns("id").AutoIncrement = True
            DS.Tables(4).Columns("id").AutoIncrementSeed = -1
            DS.Tables(4).Columns("id").AutoIncrementStep = -1

            Dim pk5(0) As DataColumn
            pk5(0) = DS.Tables(5).Columns("id")
            DS.Tables(5).PrimaryKey = pk5
            DS.Tables(5).Columns("id").AutoIncrement = True
            DS.Tables(5).Columns("id").AutoIncrementSeed = -1
            DS.Tables(5).Columns("id").AutoIncrementStep = -1

            Dim pk8(0) As DataColumn
            pk8(0) = DS.Tables(8).Columns("id")
            DS.Tables(8).PrimaryKey = pk8
            DS.Tables(8).Columns("id").AutoIncrement = True
            DS.Tables(8).Columns("id").AutoIncrementSeed = -1
            DS.Tables(8).Columns("id").AutoIncrementStep = -1
            DS.Tables(8).TableName = "VendorInfoModiAction"

            Dim HDCol As DataColumn
            Dim DTLCol As DataColumn
            Dim CHDCol As DataColumn
            Dim CDTLCol As DataColumn
            Dim DocCol As DataColumn
            Dim ActionCol As DataColumn

            HDCol = DS.Tables(0).Columns("id")
            DTLCol = DS.Tables(1).Columns("hdid")

            CHDCol = DS.Tables(4).Columns("id")
            CDTLCol = DS.Tables(3).Columns("contactid")

            DocCol = DS.Tables(5).Columns("vendorinfmodiid")
            ActionCol = DS.Tables(8).Columns("vendorinfomodiid")

            'DS.EnforceConstraints = False

            Dim rel As New DataRelation("HDRel", DS.Tables(0).Columns("id"), DS.Tables(1).Columns("hdid"))
            DS.Relations.Add(rel)

            Dim rel2 = New DataRelation("CHDDTLRel", CHDCol, CDTLCol)
            DS.Relations.Add(rel2)

            Dim rel3 = New DataRelation("DocHDRel", HDCol, DocCol)
            DS.Relations.Add(rel3)

            Dim rel4 = New DataRelation("ActionHDRel", HDCol, ActionCol)
            DS.Relations.Add(rel4)

            BS = New BindingSource
            DTLBS = New BindingSource

            CBS = New BindingSource
            VBS = New BindingSource
            VCBS = New BindingSource

            DocumentBS = New BindingSource
            DocTypeBS = New BindingSource

            VendorInfoModiActionBS = New BindingSource

            BS.DataSource = DS
            BS.DataMember = "Header"
            'DTLBS.DataSource = DS.Tables(1)
            DTLBS.DataSource = BS
            DTLBS.DataMember = "HDRel"

            CBS.DataSource = DS.Tables(4)
            'VCBS.DataSource = DS.Tables(3)
            VCBS.DataSource = CBS
            VCBS.DataMember = "CHDDTLRel"

            VBS.DataSource = DS.Tables(2)

            DocumentBS.DataSource = BS 'DS.Tables(5)
            DocumentBS.DataMember = "DocHDRel"


            DocTypeBS.DataSource = DS.Tables(6)
            FileSourceFullPath = DS.Tables(7).Rows(0).Item("cvalue").ToString


            VendorInfoModiActionBS.DataSource = BS
            VendorInfoModiActionBS.DataMember = "ActionHDRel"

            myret = True
        End If
        Return myret
    End Function

    Public Function GetChanges() As DataSet
        Dim myret As Boolean = False
        If IsNothing(BS) Then
            Return Nothing
        End If
        BS.EndEdit()
        VBS.EndEdit()
        CBS.EndEdit()
        DTLBS.EndEdit()
        DocumentBS.EndEdit()
        VendorInfoModiActionBS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges
        Return ds2
    End Function

    Public Function save() As Boolean Implements IAdapter.save
        Dim myret As Boolean = False
        BS.EndEdit()
        VBS.EndEdit()
        CBS.EndEdit()
        DTLBS.EndEdit()
        DocumentBS.EndEdit()
        VendorInfoModiActionBS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try

                If save(mye) Then
                    'If ds2.Tables(3).Rows.Count > 0 Then DS.Tables(3).Rows(0).Item("contactid") = ds2.Tables(3).Rows(0).Item("contactid")
                    DS.Merge(ds2)
                    DS.AcceptChanges()

                    'copy file
                    For Each drv As DataRowView In DocumentBS.List
                        If Not IsDBNull(drv.Row.Item("docfullname")) Then
                            If Not drv.Row.Item("docfullname") = "" Then
                                Dim mytarget = FileSourceFullPath & "\" & drv.Row.Item("id") & drv.Row.Item("docext")
                                Try
                                    FileIO.FileSystem.CopyFile(drv.Row.Item("docfullname"), mytarget, True)
                                Catch ex As Exception
                                    Logger.log(String.Format("** CopyFile error {0}**", ex.Message))
                                End Try
                            End If
                        End If
                       
                    Next

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

    Public Function Save(ByVal mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Dim myret As Boolean = False
        If myModel.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function
    '{"vendorname", "shortname", "applicantname", "vendorcode::character varying", "lateststatus", "creatorname", "doc.getmodifiedfield(q.id)"}

    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format("[vendorname] like '*{0}*' or [shortname] like '*{0}*' or [vendorcodetext] like '*{0}*' or [applicantname] like '*{0}*' or [creatorname] like '*{0}*'", value)
        End Set
    End Property

    Public Function GetCurrentRecord() As System.Data.DataRowView Implements IToolbarAction.GetCurrentRecord
        Return BS.Current
    End Function

    Public Function GetNewRecord() As System.Data.DataRowView Implements IToolbarAction.GetNewRecord
        Return BS.AddNew
    End Function

    Public Sub RemoveAt(ByVal value As Integer) Implements IToolbarAction.RemoveAt
        BS.RemoveAt(value)
    End Sub

    Public Function loaddata1() As Boolean Implements IAdapter.loaddata
        Return Nothing
    End Function

    Public Function getBindingSource() As BindingSource
        Return BS
    End Function

    Public Function getVCBS() As BindingSource
        Return VCBS
    End Function

    Public Function getCBS() As BindingSource
        Return CBS
    End Function

    Public Function getVBS() As BindingSource
        Return VBS
    End Function

    Public Function getdtlBindingSource() As BindingSource
        Return DTLBS
    End Function

    Public Function GetDocumentBS() As BindingSource
        Return DocumentBS
    End Function

    Public Function getDocTypeBS() As BindingSource
        Return DocTypeBS
    End Function

    Public Function getsuppliermodificationid() As String
        Dim myret As String = String.Empty
        Dim drv As DataRowView = BS.Current
        myret = String.Format("{0:########0}_{1:yyyyMMdd}_{2:0000}", drv.Row.Item("vendorcode"), drv.Row.Item("applicantdate"), drv.Row.Item("id"))
        Return myret
    End Function

    Function loaddata(ByVal Criteria As String) As Boolean
        Dim myret As Boolean = False
        DS = New DataSet
        myModel = New VendorInformationModificationModel
        If myModel.LoadData(DS, Criteria) Then
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret
    End Function

    Function getFileSourceFullPath() As String
        Return FileSourceFullPath
    End Function

    Function loaddataNewVendor(ByVal Criteria As String) As Boolean
        Dim myret As Boolean = False
        DS = New DataSet
        myModel = New VendorInformationModificationModel
        If myModel.LoadDataNewVendor(DS, Criteria) Then
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret
    End Function

End Class
