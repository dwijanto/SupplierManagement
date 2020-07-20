Public Class ClassImportIdentityAttachment
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
                docdrv.Item("docdate") = myrecord(30)
                docdrv.Item("doctype") = "Identity sheet"
                docdrv.EndEdit()


                'Assign record For siltx
                'For i = 0 To 58
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
