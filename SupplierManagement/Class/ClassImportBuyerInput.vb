Imports System.Text

Public Class ClassImportBuyerInput
    Private _Filename As String
    Private _DRV As DataRowView
    Public errormessage As String
    Public Sub New(ByVal drv As DataRowView, ByVal filename As String)
        Me._DRV = drv
        Me._Filename = filename
    End Sub
    Public Function getRecord() As Boolean

        Dim myret As Boolean = False
        Try
            'Read Text file
            Dim mystr As New StringBuilder
            Dim myInsert As New System.Text.StringBuilder
            Dim myrecord() As String
            Using objTFParser = New FileIO.TextFieldParser(_Filename)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count >= 1 Then
                            _DRV.Row.Item("producttype") = myrecord(0)
                            '_DRV.Row.Item("shortname") = myrecord(0)
                            _DRV.Row.Item("lastvisit") = If(myrecord(1) = "", DBNull.Value, myrecord(1))
                            _DRV.Row.Item("keystakestopic") = myrecord(2)
                            _DRV.Row.Item("strength1") = myrecord(3)
                            _DRV.Row.Item("strength2") = myrecord(4)
                            _DRV.Row.Item("strength3") = myrecord(5)
                            _DRV.Row.Item("weaknessess1") = myrecord(6)
                            _DRV.Row.Item("weaknessess2") = myrecord(7)
                            _DRV.Row.Item("weaknessess3") = myrecord(8)
                            _DRV.Row.Item("opportunities1") = myrecord(9)
                            _DRV.Row.Item("opportunities2") = myrecord(10)
                            _DRV.Row.Item("opportunities3") = myrecord(11)
                            _DRV.Row.Item("threats1") = myrecord(12)
                            _DRV.Row.Item("threats2") = myrecord(13)
                            _DRV.Row.Item("threats3") = myrecord(14)
                            _DRV.Row.Item("myyear2") = myrecord(15)
                            _DRV.Row.Item("currbudget") = If(myrecord(16) = "", DBNull.Value, myrecord(16))
                            _DRV.Row.Item("totalbudget") = If(myrecord(17) = "", DBNull.Value, myrecord(17))
                            _DRV.Row.Item("productdevelopment1") = myrecord(18)
                            _DRV.Row.Item("productdevelopment2") = myrecord(19)
                            _DRV.Row.Item("negotiationresult1") = myrecord(20)
                            _DRV.Row.Item("negotiationresult2") = myrecord(21)
                            _DRV.Row.Item("negotiationresult3") = myrecord(22)
                            _DRV.Row.Item("outstandingissue1") = myrecord(23)
                            _DRV.Row.Item("outstandingissue2") = myrecord(24)
                            _DRV.Row.Item("outstandingissue3") = myrecord(25)
                            _DRV.Row.Item("outstandingissue4") = myrecord(26)
                            _DRV.Row.Item("outstandingissue5") = myrecord(27)
                            'Additional
                            Try
                                _DRV.Row.Item("keycontactperson") = myrecord(28)
                                _DRV.Row.Item("title") = myrecord(29)
                                _DRV.Row.Item("top3cust1") = myrecord(30)
                                _DRV.Row.Item("top3cust2") = myrecord(31)
                                _DRV.Row.Item("top3cust3") = myrecord(32)
                            Catch ex As Exception

                            End Try
                           


                        End If
                        count += 1
                    Loop
                End With
            End Using

            myret = True
        Catch ex As Exception
            myret = False
            errormessage = ex.Message
        End Try

        Return myret
    End Function

    Private Function validate(ByVal myrecord As String()) As Boolean
        Throw New NotImplementedException
    End Function


End Class
