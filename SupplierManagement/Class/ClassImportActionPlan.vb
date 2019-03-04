Imports System.Text
Imports SupplierManagement.PublicClass
Public Class ClassImportActionPlan
    Private _Filename As String
    Private bs As BindingSource
    Private drv As DataRowView
    Public errormessage As String
    Public shortname As String
    Private Validator As String = String.Empty
    Dim myStatus As String = "On-going,Closed,Delay,Not started"
 
    Public Sub New(ByVal bs As BindingSource, ByVal drv As DataRowView, ByVal filename As String, ByVal validator As String)
        Me.bs = bs
        Me.drv = drv
        Me._Filename = filename
        Me.Validator = validator
    End Sub
    Public Sub New(ByVal bs As BindingSource, ByVal drv As DataRowView, ByVal filename As String)
        Me.bs = bs
        Me.drv = drv
        Me._Filename = filename
    End Sub
    Sub New(ByVal BS As BindingSource, ByVal filename As String, ByVal validator As String)
        ' TODO: Complete member initialization 
        Me.bs = BS
        Me._Filename = filename
        Me.Validator = validator
    End Sub
    Sub New(ByVal BS As BindingSource, ByVal filename As String)
        ' TODO: Complete member initialization 
        Me.bs = BS
        Me._Filename = filename
    End Sub

    Public Function getRecord() As Boolean

        Dim myret As Boolean = False
        Try
            cleandata()
            'Read Text file
            Dim mystr As New StringBuilder
            Dim myInsert As New System.Text.StringBuilder
            Dim myrecord() As String
            Dim mycheck As Boolean = False
            Dim errdesc As String = String.Empty
            Using objTFParser = New FileIO.TextFieldParser(_Filename)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    Dim Lineno As Integer = 0
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count >= 3 Then
                            If myrecord(1) <> "" And myrecord(1) <> "." Then
                                'Check ValidStatus
                                If Not myStatus.Contains(myrecord(10)) Then
                                    'Err.Raise(555, String.Format("Status ""{0}"" not valid!", myrecord(10)))
                                    mycheck = True

                                    errdesc = String.Format("Status ""{0}"" not valid!", myrecord(10))

                                End If
                                'Lineno += 1
                                'Check Closed status

                                If myrecord(10) = "Closed" And Me.Validator.Length = 0 Then
                                    mycheck = True
                                    errdesc = String.Format("Status Closed, Validator is blank.")
                                End If

                                If mycheck Then
                                    For i = 0 To bs.Count - 1
                                        bs.Position = 1
                                        bs.RemoveCurrent()
                                    Next
                                    Err.Raise(513, Description:=errdesc)
                                End If

                                Dim drv As DataRowView = bs.AddNew
                                drv.BeginEdit()
                                drv.Row.Item("documentid") = Me.drv.Row.Item("documentid")
                                'drv.Row.Item("lineno") = Lineno
                                drv.Row.Item("vdid") = Me.drv.Row.Item("vdid")
                                drv.Row.Item("priority") = myrecord(1)
                                drv.Row.Item("situation") = myrecord(2)
                                drv.Row.Item("target") = myrecord(3)
                                drv.Row.Item("proposal") = myrecord(4)
                                drv.Row.Item("responsibleperson") = myrecord(5)
                                drv.Row.Item("startdate") = If(myrecord(6) = "", DBNull.Value, myrecord(6))
                                drv.Row.Item("enddate") = If(myrecord(7) = "", DBNull.Value, myrecord(7))
                                drv.Row.Item("result") = myrecord(8)
                                drv.Row.Item("finishdate") = If(myrecord(9) = "", DBNull.Value, myrecord(9))
                                If myrecord(10) = "Closed" Then
                                    drv.Row.Item("validator") = "" & Validator
                                End If

                                drv.Row.Item("status") = myrecord(10)
                                If myrecord(11).Length > 0 Then
                                    'drv.Row.Item("actionid")
                                    drv.Row.Item("actionid") = myrecord(11)
                                End If
                                drv.Row.Item("shortname") = shortname
                                drv.Row.Item("uploaddate") = Date.Today
                                drv.Row.Item("modifiedby") = HelperClass1.UserId
                                'drv.Row.Item("shortname") 
                                'drv.row.item("uploaddate")

                                drv.EndEdit()
                            End If

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

    Private Sub cleandata()
        'Clean existing bindingSource
        If Not IsNothing(bs.Current) Then
            bs.Filter = String.Format("documentid = {0}", drv.Item("documentid"))
            For Each mydrv As DataRowView In bs.List
                mydrv.BeginEdit()
                mydrv.Delete()
                mydrv.EndEdit()
            Next
            bs.Filter = ""
        End If
    End Sub
    Private Sub cleannewdata()
        'Clean existing bindingSource
        If Not IsNothing(bs.Current) Then
            For Each mydrv As DataRowView In bs.List
                mydrv.BeginEdit()
                mydrv.Delete()
                mydrv.EndEdit()
            Next           
        End If
    End Sub

    Public Function getNewRecord() As Boolean
        Dim myret As Boolean = False
        Dim mycheck As Boolean = False
        Dim errdesc As String = String.Empty
        Try
            cleannewdata()
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
                    Dim Lineno As Integer = 0
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count >= 1 Then
                            If myrecord(1) <> "" And myrecord(1) <> "." Then

                                If Not myStatus.Contains(myrecord(11)) Then
                                    'Err.Raise(555, String.Format("Status ""{0}"" not valid!", myrecord(10)))
                                    For i = 0 To bs.Count - 1
                                        bs.Position = 1
                                        bs.RemoveCurrent()
                                    Next
                                    Err.Raise(513, Description:=String.Format("Status ""{0}"" is not valid!", myrecord(11)))
                                End If
                                'Lineno += 1
                                Dim drv As DataRowView = bs.AddNew
                                drv.BeginEdit()
                                drv.Row.Item("documentid") = 0
                                'drv.Row.Item("lineno") = Lineno
                                'drv.Row.Item("vdid") = Me.drv.Row.Item("vdid")
                                drv.Row.Item("priority") = myrecord(2)
                                drv.Row.Item("situation") = myrecord(3)
                                drv.Row.Item("target") = myrecord(4)
                                drv.Row.Item("proposal") = myrecord(5)
                                drv.Row.Item("responsibleperson") = myrecord(6)
                                drv.Row.Item("startdate") = If(myrecord(7) = "", DBNull.Value, myrecord(7))
                                drv.Row.Item("enddate") = If(myrecord(8) = "", DBNull.Value, myrecord(8))
                                drv.Row.Item("result") = myrecord(9)
                                drv.Row.Item("finishdate") = If(myrecord(10) = "", DBNull.Value, myrecord(10))



                                drv.Row.Item("status") = myrecord(11)

                                If drv.Row.Item("status") = "Closed" Then
                                    drv.Row.Item("validator") = "" & Validator
                                End If
                                If drv.Row.Item("status") = "Closed" And Me.Validator.Length = 0 Then
                                    mycheck = True
                                    errdesc = String.Format("Status Closed, Validator is blank.")
                                End If

                                If mycheck Then
                                    For i = 0 To bs.Count - 1
                                        bs.Position = 1
                                        bs.RemoveCurrent()
                                    Next
                                    Err.Raise(513, Description:=errdesc)
                                End If
                                If myrecord(0).Length > 0 Then
                                    drv.Row.Item("actionid") = myrecord(0)
                                End If
                                drv.Row.Item("shortname") = myrecord(1)
                                drv.Row.Item("uploaddate") = Date.Today
                                drv.Row.Item("modifiedby") = HelperClass1.UserId
                                '

                                drv.EndEdit()
                            End If

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

End Class
