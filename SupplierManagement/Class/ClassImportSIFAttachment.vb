Imports System.Text
Imports SupplierManagement.PublicClass
Imports SupplierManagement.SharedClass
Public Class ClassImportSIFAttachment
    Inherits ClassBaseImportSIFIdentityAttachment
    Implements IImportSIFIdentity
    Public Sub New(ByRef SIFTx As BindingSource, ByRef DocumentBS As BindingSource, ByVal FileName As String)
        MyBase.New(SIFTx, DocumentBS, FileName)
    End Sub

    Public Overrides Function getRecord() As Boolean Implements IImportSIFIdentity.getRecord
        Dim myret As Boolean = False
        Try

            If MyBase.getRecord Then
                myret = True
                'Assign record for SIF
                Dim docdrv As DataRowView = _documentbs.Current
                docdrv.BeginEdit()
                docdrv.Row.Item("doctype") = "SIF"
                docdrv.Row.Item("docdate") = myrecord(14)
                docdrv.Row.Item("turnovery") = DbAdapter1.validdec(myrecord(15))
                docdrv.Row.Item("turnovery1") = DbAdapter1.validdec(myrecord(16))
                docdrv.Row.Item("turnovery2") = DbAdapter1.validdec(myrecord(17))
                docdrv.Row.Item("turnovery3") = DbAdapter1.validdec(myrecord(18))
                docdrv.Row.Item("turnovery4") = DbAdapter1.validdec(myrecord(19))
                docdrv.Row.Item("ratioy") = DbAdapter1.validdec(myrecord(20))
                docdrv.Row.Item("ratioy1") = DbAdapter1.validdec(myrecord(21))
                docdrv.Row.Item("ratioy2") = DbAdapter1.validdec(myrecord(22))
                docdrv.Row.Item("ratioy3") = DbAdapter1.validdec(myrecord(23))
                docdrv.Row.Item("ratioy4") = DbAdapter1.validdec(myrecord(24))
                docdrv.EndEdit()

                ''Assign record For siltx              
                'For i = 0 To 29
                '    If myrecord(i) <> "" Then
                '        Try
                '            If myField(i) <> 0 Then
                '                Dim sifdrv As DataRowView = _siftx.AddNew
                '                sifdrv.BeginEdit()
                '                sifdrv.Row.Item("documentid") = docdrv.Item("id")
                '                sifdrv.Row.Item("labelid") = myField(i)
                '                sifdrv.Row.Item("labelname") = silabeldict(myField(i))
                '                sifdrv.Row.Item("value") = myrecord(i)
                '                sifdrv.EndEdit()
                '            End If
                '        Catch ex As Exception
                '            MessageBox.Show(ex.Message)
                '        End Try
                '    End If
                'Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return myret
    End Function

End Class
