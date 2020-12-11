Imports System.Text

Public Class ImportPlant
    Inherits BaseAdapter
    Private myform As FormImportPlant
    Private FileName As String
    Public ErrorMessage As String
    Private SB As New StringBuilder
    Private SDUpdateSB As New StringBuilder
    Private SDPOSB As New StringBuilder
    Private Plant As Integer

    Public Sub New(ByVal myForm As FormImportPlant, ByVal FileName As String, ByVal Plant As Integer)
        Me.myform = myForm
        Me.FileName = FileName
        Me.Plant = Plant
    End Sub

    Public Function doImport() As Boolean
        'Read AASDHD 
        'Find SalesDoc
        'Update Shiptoparty and Plant
        'Read AASDPO
        'Find SalesDoc,SalesDocItem,PO,POItem
        'If Not
        Dim myret As Boolean = False
        SB.Append("select salesdoc,shiptoparty,plant from aasdhd;")
        SB.Append("select salesdoc,salesdocitem,pohd,poitem from aasdpo where not salesdoc isnull;")
        DS = New DataSet
        If DbAdapter1.TbgetDataSet(SB.ToString, DS, ErrorMessage) Then
            Try
                DS.Tables(0).TableName = "SalesDoc"
                Dim pk(0) As DataColumn
                pk(0) = DS.Tables(0).Columns("salesdoc")
                DS.Tables(0).PrimaryKey = pk

                DS.Tables(1).TableName = "RelSalesPO"
                Dim pk1(4) As DataColumn
                pk1(0) = DS.Tables(1).Columns("salesdoc")
                pk1(1) = DS.Tables(1).Columns("salesdocitem")
                pk1(2) = DS.Tables(1).Columns("pohd")
                pk1(3) = DS.Tables(1).Columns("poitem")
                DS.Tables(1).PrimaryKey = pk1
            Catch ex As Exception
                ErrorMessage = ex.Message
                Return myret
            End Try
        Else
            Return myret
        End If


        SDUpdateSB.Clear()
        SDPOSB.Clear()
        Try
            Dim myRecord() As String
            Using objTFParser = New FileIO.TextFieldParser(FileName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = False
                    Dim count As Long = 0
                    myform.ProgressReport(1, "Read Data")
                    myform.ProgressReport(2, "Continuous")

                    Dim result As DataRow
                    Do Until .EndOfData
                        myRecord = .ReadFields
                        If count > 1 Then
                            Dim pkey(0) As Object
                            pkey(0) = myRecord(1)
                            result = DS.Tables(0).Rows.Find(pkey)

                            If Not IsNothing(result) Then
                                'update
                                If SDUpdateSB.Length > 0 Then
                                    SDUpdateSB.Append(",")
                                End If
                                SDUpdateSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying]",
                                                                  myRecord(1), myRecord(9), Plant))

                            End If
                            Dim pkey1(3) As Object
                            pkey1(0) = myRecord(1)
                            pkey1(1) = myRecord(2)
                            pkey1(2) = myRecord(11)
                            pkey1(3) = myRecord(12)

                            result = DS.Tables(1).Rows.Find(pkey1)
                            If IsNothing(result) Then
                                SDPOSB.Append(myRecord(1) & vbTab &
                                                    myRecord(2) & vbTab &
                                                    myRecord(11) & vbTab &
                                                    myRecord(12) & vbCrLf)
                                Dim dr As DataRow = DS.Tables(1).NewRow
                                dr.Item(0) = myRecord(1)
                                dr.Item(1) = myRecord(2)
                                dr.Item(2) = myRecord(11)
                                dr.Item(3) = myRecord(12)
                                DS.Tables(1).Rows.Add(dr)
                            End If
                        End If

                        count += 1
                    Loop
                End With
            End Using
            'update record
            Dim sqlstr As String = String.Empty

            If SDPOSB.Length > 0 Then
                sqlstr = "copy aasdpo(salesdoc,salesdocitem,pohd,poitem) from stdin with null as 'Null';"
                ErrorMessage = DbAdapter1.copy(sqlstr, SDPOSB.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If SDUpdateSB.Length > 0 Then

                myform.ProgressReport(1, "Update AASDHD")
                sqlstr = "update aasdhd set plant= foo.plant::integer,shiptoparty= foo.shiptoparty::bigint" &
                " from (select * from array_to_set3(Array[" & SDUpdateSB.ToString &
                "]) as tb (id character varying,shiptoparty character varying,plant character varying))as foo " &
                " where salesdoc = foo.id::bigint;"

                myret = DbAdapter1.ExecuteNonQuery(sqlstr, message:=ErrorMessage)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        

        
        Return myret
    End Function

    Public Function getLastImportData() As Date
        Dim myret As Date
        Dim Sqlstr = "select enddate from programlocking where progname = 'FIMSO'"
        Return myret
    End Function
End Class
