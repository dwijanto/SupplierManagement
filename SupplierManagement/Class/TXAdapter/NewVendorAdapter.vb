'The Main Table is doc.vendorinfomodi 
'The other table to store NewVendor is doc.vendorinfomodiext
'The relation is 1 to 1 from doc.vendorinfomodi to doc.vendorinfomodiext

Public Class NewVendorAdapter
    Implements IAdapter

    Implements IToolbarAction
    Public myModel As NewVendorModel
    Dim DS As DataSet
    Dim BS As BindingSource

    Public VendorInfoModiActionBS As BindingSource
    Dim DocumentBS As BindingSource
    Dim SITxAttachmentBS As BindingSource
    Dim ZetolAttachmentBS As BindingSource
    Dim ZetolMasterBS As BindingSource
    Dim DocTypeBS As BindingSource

    Dim LevelBS As BindingSource
    Dim PaymentTermBS As BindingSource

    Private FileSourceFullPath As String
    Private DocFileSourceFullPath As String

    Public Function getsuppliermodificationid() As String
        Dim myret As String = String.Empty
        Dim drv As DataRowView = BS.Current
        myret = String.Format("{0}_{1:yyyyMMdd}_{2:0000}", drv.Row.Item("shortname"), drv.Row.Item("applicantdate"), drv.Row.Item("id"))
        Return myret
    End Function

    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return Nothing
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Function GetCurrentRecord() As System.Data.DataRowView Implements IToolbarAction.GetCurrentRecord
        Return BS.Current
    End Function

    Public Function GetNewRecord() As System.Data.DataRowView Implements IToolbarAction.GetNewRecord
        Return BS.AddNew
    End Function

    Public Sub RemoveAt(ByVal value As Integer) Implements IToolbarAction.RemoveAt

    End Sub

    Public Function GetDocumentBS() As BindingSource
        Return DocumentBS
    End Function


    Public Function GetLevelBS() As BindingSource
        Return LevelBS
    End Function


    Public Function GetPaymentTermBS() As BindingSource
        Return PaymentTermBS
    End Function

    Public Function GetSITxAttachmentBS() As BindingSource
        Return SITxAttachmentBS
    End Function
    Public Function GetZetolAttachmentBS() As BindingSource
        Return ZetolAttachmentBS
    End Function

    Public Function GetZetolMasterBS() As BindingSource
        Return ZetolMasterBS
    End Function

    Public Function getDocTypeBS() As BindingSource
        Return DocTypeBS
    End Function

    Function getFileSourceFullPath() As String
        Return FileSourceFullPath
    End Function

    Function getDocFileSourceFullPath() As String
        Return DocFileSourceFullPath
    End Function

    Public Function loaddata() As Boolean Implements IAdapter.loaddata
        Return Nothing
    End Function

    Public Function loaddata(ByVal Id As Long) As Boolean
        Dim myret As Boolean = False
        myModel = New NewVendorModel
        DS = New DataSet
        If myModel.LoadData(DS, Id) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).PrimaryKey = pk
            DS.Tables(0).Columns("id").AutoIncrement = True
            DS.Tables(0).Columns("id").AutoIncrementSeed = -1
            DS.Tables(0).Columns("id").AutoIncrementStep = -1
            DS.Tables(0).TableName = "Master" 'doc.vendorinfomodi + doc.vendorinfomodiext

            Dim pk1(0) As DataColumn
            pk1(0) = DS.Tables(1).Columns("id")
            DS.Tables(1).PrimaryKey = pk1
            DS.Tables(1).Columns("id").AutoIncrement = True
            DS.Tables(1).Columns("id").AutoIncrementSeed = -1
            DS.Tables(1).Columns("id").AutoIncrementStep = -1
            DS.Tables(1).TableName = "VendorInfoModiAction"

            Dim pk2(0) As DataColumn
            pk2(0) = DS.Tables(2).Columns("id")
            DS.Tables(2).PrimaryKey = pk2
            DS.Tables(2).Columns("id").AutoIncrement = True
            DS.Tables(2).Columns("id").AutoIncrementSeed = -1
            DS.Tables(2).Columns("id").AutoIncrementStep = -1
            DS.Tables(2).TableName = "VendorInfoModiDocument" ' + ext

            Dim pk5(0) As DataColumn
            pk5(0) = DS.Tables(5).Columns("id")
            DS.Tables(5).PrimaryKey = pk5
            DS.Tables(5).Columns("id").AutoIncrement = True
            DS.Tables(5).Columns("id").AutoIncrementSeed = -1
            DS.Tables(5).Columns("id").AutoIncrementStep = -1
            DS.Tables(5).TableName = "SITXattachment" ' + ext

            Dim pk6(0) As DataColumn
            pk6(0) = DS.Tables(6).Columns("id")
            DS.Tables(6).PrimaryKey = pk6
            DS.Tables(6).Columns("id").AutoIncrement = True
            DS.Tables(6).Columns("id").AutoIncrementSeed = -1
            DS.Tables(6).Columns("id").AutoIncrementStep = -1
            DS.Tables(6).TableName = "ZetolAttachment"

            BS = New BindingSource
            VendorInfoModiActionBS = New BindingSource
            DocumentBS = New BindingSource
            DocTypeBS = New BindingSource
            SITxAttachmentBS = New BindingSource
            ZetolAttachmentBS = New BindingSource

            LevelBS = New BindingSource
            PaymentTermBS = New BindingSource
            ZetolMasterBS = New BindingSource

            BS.DataSource = DS.Tables(0)

            'DS.Tables(1).TableName = "VendorInfoModiAction"

            Dim HDCol As DataColumn
            Dim ActionCol As DataColumn
            Dim DocCol As DataColumn
            Dim SItxCol As DataColumn
            Dim ZetolCol As DataColumn
            Dim AttachCol As DataColumn



            'Create a relation between Vendorinfomodi with Actions and Documents

            HDCol = DS.Tables(0).Columns("id")
            ActionCol = DS.Tables(1).Columns("vendorinfomodiid")
            DocCol = DS.Tables(2).Columns("vendorinfmodiid")

            'Vendorinfmodiattachment - siftxattachment
            AttachCol = DS.Tables(2).Columns("id")
            SItxCol = DS.Tables(5).Columns("attachmentid")

            'Vendorinfmodiattachement - zetolattachment
            AttachCol = DS.Tables(2).Columns("id")
            ZetolCol = DS.Tables(6).Columns("attachmentid")

            Dim rel = New DataRelation("ActionHDRel", HDCol, ActionCol)
            DS.Relations.Add(rel)
            Dim rel2 = New DataRelation("DocHDRel", HDCol, DocCol)
            DS.Relations.Add(rel2)

            Dim rel3 = New DataRelation("AttachSIRel", AttachCol, SItxCol)
            DS.Relations.Add(rel3)

            Dim rel4 = New DataRelation("ZetolAttachRel", AttachCol, ZetolCol)
            DS.Relations.Add(rel4)

            VendorInfoModiActionBS.DataSource = BS 'Table 1
            VendorInfoModiActionBS.DataMember = "ActionHDRel"

            DocumentBS.DataSource = BS 'Tables 2
            DocumentBS.DataMember = "DocHDRel"

            'Create a relation between Documents and SITxAttachment
            SITxAttachmentBS.DataSource = DocumentBS
            SITxAttachmentBS.DataMember = "AttachSIRel"

            'Create a relation between Documents and ZetolAttachment
            ZetolAttachmentBS.DataSource = DocumentBS
            ZetolAttachmentBS.DataMember = "ZetolAttachRel"

            'Dim UC As UniqueConstraint
            'UC = New UniqueConstraint(New DataColumn() {DS.Tables("ZetolAttachment").Columns("attachmentid"),
            '                                            DS.Tables("ZetolAttachment").Columns("vdid"),
            '                                            DS.Tables("ZetolAttachment").Columns("zetolid")})
            'If Not IsNothing(UC) Then
            '    DS.Tables("ZetolAttachment").Constraints.Add(UC)
            'End If

            DocTypeBS.DataSource = DS.Tables(3)
            FileSourceFullPath = DS.Tables(4).Rows(0).Item("cvalue").ToString
            DocFileSourceFullPath = DS.Tables(11).Rows(0).Item("cvalue").ToString
            LevelBS.DataSource = DS.Tables(8)
            PaymentTermBS.DataSource = DS.Tables(9)
            ZetolMasterBS.DataSource = DS.Tables(10)
            myret = True
        End If
        Return myret

    End Function

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
                    'DS.AcceptChanges()
                    'Don't use DS.AcceptChanges. Use the statement below.
                    'Reason: Only AcceptChanges for modified Table. if unmodified table use AcceptChanges -> the position is set to first Row (not correct)
                    For Each mytable As DataTable In ds2.Tables
                        If mytable.Rows.Count > 0 Then
                            DS.Tables(mytable.TableName).AcceptChanges()
                        End If
                    Next
                    'copy file
                    For Each drv As DataRowView In DocumentBS.List
                        If Not IsDBNull(drv.Row.Item("docfullname")) Then
                            If Not drv.Row.Item("docfullname") = "" Then
                                Dim mytarget = FileSourceFullPath & "\" & drv.Row.Item("id") & drv.Row.Item("docext")
                                Dim myDocTarget = DocFileSourceFullPath & "\" & drv.Row.Item("documentid") & drv.Row.Item("docext")
                                Try
                                    FileIO.FileSystem.CopyFile(drv.Row.Item("docfullname"), mytarget, True)
                                    FileIO.FileSystem.CopyFile(drv.Row.Item("docfullname"), myDocTarget, True)
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

End Class
