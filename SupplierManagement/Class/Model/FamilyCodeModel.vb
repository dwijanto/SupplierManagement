Imports Npgsql
Imports System.Text

Public Class FamilyCodeModel
    Implements IModel

    Dim myadapter As DbAdapter = DbAdapter.getInstance
    Private _ErrorMessage As String

    Public ReadOnly Property ErrorMessage As String
        Get
            Return _ErrorMessage
        End Get
    End Property
    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "doc.ifamilycode"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "familycode"
        End Get
    End Property

    Public ReadOnly Property FilterField
        Get
            Return "[familycode] like '*{0}*' or [pop] like '*{0}*' or [popdesc] like '*{0}*' or [family] like '*{0}*'  or [familydesc] like '*{0}*' or [sbfam] like '*{0}*'  or [sbfamdesc] like '*{0}*'"
        End Get
    End Property

    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.*,p.pop,p.description as popdesc,f.family,f.description as familydesc,sb.sbfam,sb.description as sbfamdesc from {0} u " &
                                       " left join doc.ipopulation p on p.id = u.popid " &
                                       " left join doc.ifamily f on f.id = u.familyid" &
                                       " left join doc.isubfamily sb on sb.id = u.subfamid order by {1} ;", TableName, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Return Nothing
    End Function

    Sub ExportToExcel(ByVal myform As Object)
        Dim mymessage As String = String.Empty

        Dim sqlstr = String.Format("select u.familycode,p.pop,p.description as popdesc,f.family,f.description as familydesc,sb.sbfam,sb.description as sbfamdesc from {0} u " &
                                        " left join doc.ipopulation p on p.id = u.popid " &
                                        " left join doc.ifamily f on f.id = u.familyid" &
                                        " left join doc.isubfamily sb on sb.id = u.subfamid order by {1} ;", TableName, SortField)

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("IndirectFamilyCode{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1 'because hidden

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(myform, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            myreport.Run(myform, New EventArgs)
        End If
    End Sub

    Function ImportFromText(ByVal Myform As FormIndirectFamily, ByVal mySelectedPath As String)
        Dim sb As New StringBuilder
        Dim myret As Boolean = False

        Dim list As New List(Of String)
        Dim myList As New List(Of String())

        Myform.ProgressReport(1, "Open Text File...")
        Dim i As Long
        Dim DS As New DataSet
        Try
            Dim dir As New IO.DirectoryInfo(mySelectedPath)
            Dim myrecord() As String
            Dim tcount As Long = 0
            Myform.ProgressReport(6, "Set To Marque")
            Myform.ProgressReport(1, String.Format("Read Text File...{0}", mySelectedPath))
            Using objTFParser = New FileIO.TextFieldParser(mySelectedPath)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = False
                    Dim count As Long = 0

                    Do Until .EndOfData
                        'If count > 0 Then
                        myrecord = .ReadFields
                        If count > 1 Then
                            myList.Add(myrecord)
                        End If

                        tcount += 1
                        'End If
                        count += 1

                    Loop
                End With
            End Using
            
            If myList.Count = 0 Then
                _ErrorMessage = "Nothing to process."
                Myform.ProgressReport(3, _ErrorMessage)
                Return myret
            End If


            Myform.ProgressReport(1, String.Format("Prepare Data..........."))


            PrepareData(DS)

            Myform.ProgressReport(1, String.Format("Build Data row..........."))
            Myform.ProgressReport(5, "Set To Continuous")


            'Find Record
            'doc.ipopulation
            'doc.ifamily
            'doc.isubfamily
            'doc.ifamilycode

            For i = 0 To myList.Count - 1
                'Check Record
                Dim mykey3(0) As Object
                Dim mydata = myList(i)
                mykey3(0) = mydata(0)
                Dim dr As DataRow = DS.Tables(3).Rows.Find(mykey3)
                If IsNothing(dr) Then
                    'Check Population
                    dr = DS.Tables(3).NewRow
                    CheckPopulation(mydata, DS, dr)

                    'Check Family
                    CheckFamily(mydata, DS, dr)
                    'Check SubFamily
                    CheckSubFamily(mydata, DS, dr)
                    dr.Item("familycode") = mydata(0)
                    DS.Tables(3).Rows.Add(dr)
                Else
                    'don't care about existing record.
                End If

            Next
            Myform.ProgressReport(6, "Set To Marque")

            'Using save transaction
            Try
                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                Dim ds2 As DataSet
                ds2 = DS.GetChanges

                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                    If Not myadapter.FamilyCodeTx(Myform, mye) Then
                        DS.Merge(ds2)
                        MessageBox.Show(mye.message)
                        Return False
                    End If
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    'DataGridView1.Invalidate()
                    MessageBox.Show("Saved.")
                End If
            Catch ex As Exception
                MessageBox.Show(" Error:: " & ex.Message)
            End Try


            myret = True
            Myform.ProgressReport(1, String.Format("Done."))
        Catch ex As Exception
            _ErrorMessage = String.Format("Error : {0} ", i) & ex.Message
            Myform.ProgressReport(1, String.Format(_ErrorMessage))
        End Try

        Return myret


    End Function
    Private Sub CheckPopulation(ByVal mydata() As String, ByRef DS As DataSet, ByRef mydr As DataRow)
        Dim mykey(0) As Object
        mykey(0) = mydata(1)
        Dim dr As DataRow = DS.Tables(0).Rows.Find(mykey)
        If IsNothing(dr) Then
            'Check Population
            dr = DS.Tables(0).NewRow
            dr.Item("pop") = mydata(1)
            dr.Item("description") = mydata(2)
            DS.Tables(0).Rows.Add(dr)
        Else
            'mydr.Item("popid") = dr.Item("id")
        End If
        mydr.Item("popid") = dr.Item("id")
    End Sub

    Private Sub CheckFamily(ByVal mydata() As String, ByRef DS As DataSet, ByRef mydr As DataRow)
        Dim mykey(0) As Object
        mykey(0) = mydata(1) + mydata(3)
        Dim dr As DataRow = DS.Tables(1).Rows.Find(mykey)
        If IsNothing(dr) Then
            'Check Population
            dr = DS.Tables(1).NewRow

            dr.Item("key") = mykey(0)
            dr.Item("family") = mydata(3)
            dr.Item("description") = mydata(4)
            DS.Tables(1).Rows.Add(dr)
        Else
            'mydr.Item("familyid") = dr.Item("id")
        End If
        mydr.Item("familyid") = dr.Item("id")
    End Sub

    Private Sub CheckSubFamily(ByVal mydata() As String, ByRef DS As DataSet, ByRef mydr As DataRow)
        Dim mykey(0) As Object
        mykey(0) = mydata(1) + mydata(3) + mydata(5)
        Dim dr As DataRow = DS.Tables(2).Rows.Find(mykey)
        If IsNothing(dr) Then
            'Check Population
            dr = DS.Tables(2).NewRow

            dr.Item("key") = mykey(0)
            dr.Item("sbfam") = mydata(5)
            dr.Item("description") = mydata(6)
            DS.Tables(2).Rows.Add(dr)
        Else
            'mydr.Item("subfamid") = dr.Item("id")
        End If
        mydr.Item("subfamid") = dr.Item("id")
    End Sub


    Private Sub FormattingReport()
        ' Throw New NotImplementedException
    End Sub

    Private Sub PivotTable()
        'Throw New NotImplementedException
    End Sub

    Private Sub PrepareData(ByVal DS As DataSet)
        Dim sb As New StringBuilder
        sb.Append("select * from doc.ipopulation;")
        sb.Append("select * from doc.ifamily;")
        sb.Append("select * from doc.isubfamily;")
        sb.Append("select * from doc.ifamilycode;")
        Dim sqlstr = sb.ToString

        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using

        Dim pk(0) As DataColumn
        pk(0) = DS.Tables(0).Columns("pop")
        DS.Tables(0).Columns("id").AutoIncrement = True
        DS.Tables(0).Columns("id").AutoIncrementSeed = 0
        DS.Tables(0).Columns("id").AutoIncrementStep = -1
        DS.Tables(0).PrimaryKey = pk
        DS.Tables(0).TableName = "Population"

        Dim pk1(0) As DataColumn
        pk1(0) = DS.Tables(1).Columns("key")
        DS.Tables(1).Columns("id").AutoIncrement = True
        DS.Tables(1).Columns("id").AutoIncrementSeed = 0
        DS.Tables(1).Columns("id").AutoIncrementStep = -1
        DS.Tables(1).PrimaryKey = pk1
        DS.Tables(1).TableName = "Family"

        Dim pk2(0) As DataColumn
        pk2(0) = DS.Tables(2).Columns("key")
        DS.Tables(2).Columns("id").AutoIncrement = True
        DS.Tables(2).Columns("id").AutoIncrementSeed = 0
        DS.Tables(2).Columns("id").AutoIncrementStep = -1
        DS.Tables(2).PrimaryKey = pk2
        DS.Tables(2).TableName = "Sub Family"

        Dim pk3(0) As DataColumn
        pk3(0) = DS.Tables(3).Columns("familycode")
        DS.Tables(3).Columns("id").AutoIncrement = True
        DS.Tables(3).Columns("id").AutoIncrementSeed = 0
        DS.Tables(3).Columns("id").AutoIncrementStep = -1
        DS.Tables(3).PrimaryKey = pk3
        DS.Tables(3).TableName = "Family Code"

        'Create Relation Population - FamilyCode
        Dim myrel As DataRelation
        Dim PopulationHDCol As DataColumn
        Dim PopulationDTCol As DataColumn
        PopulationHDCol = DS.Tables(0).Columns("id")
        PopulationDTCol = DS.Tables(3).Columns("popid")
        myrel = New DataRelation("PopulationRel", PopulationHDCol, PopulationDTCol)
        DS.Relations.Add(myrel)

        'Create Relation Family - FamilyCode
        Dim FamilyHDCol As DataColumn
        Dim FamilyDTCol As DataColumn
        FamilyHDCol = DS.Tables(1).Columns("id")
        FamilyDTCol = DS.Tables(3).Columns("familyid")
        myrel = New DataRelation("FamilyRel", FamilyHDCol, FamilyDTCol)
        DS.Relations.Add(myrel)

        'Create Relation Family - FamilyCode
        Dim SubFamilyHDCol As DataColumn
        Dim SubFamilyDTCol As DataColumn
        SubFamilyHDCol = DS.Tables(2).Columns("id")
        SubFamilyDTCol = DS.Tables(3).Columns("subfamid")
        myrel = New DataRelation("SubFamilyRel", SubFamilyHDCol, SubFamilyDTCol)
        DS.Relations.Add(myrel)

    End Sub



End Class
