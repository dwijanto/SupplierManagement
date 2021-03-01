Imports Npgsql
Imports NpgsqlTypes
Imports System.IO

Imports SupplierManagement.PublicClass

Public Class DbAdapter
    Implements IDisposable

    Dim _ConnectionStringDict As Dictionary(Of String, String)
    Dim _connectionstring As String
    Private CopyIn1 As NpgsqlCopyIn
    Dim _userid As String
    Dim _password As String
    Dim mytransaction As NpgsqlTransaction

    Public Shared myInstance As DbAdapter

    Public Shared Function getInstance() As DbAdapter
        If myInstance Is Nothing Then
            myInstance = New DbAdapter
        End If
        Return myInstance
    End Function


    Public ReadOnly Property userid As String
        Get
            Return _userid
        End Get
    End Property
    Public ReadOnly Property password As String
        Get
            Return _password
        End Get
    End Property

    Public Property Connectionstring As String
        Get
            Return _connectionstring

        End Get
        Set(ByVal value As String)
            _connectionstring = value
        End Set
    End Property

    Public Sub New()
        InitConnectionStringDict()
        _connectionstring = getConnectionString()
    End Sub

    Public ReadOnly Property ConnectionStringDict As Dictionary(Of String, String)
        Get
            Return _ConnectionStringDict
        End Get
    End Property

    Private Sub InitConnectionStringDict()
        _ConnectionStringDict = New Dictionary(Of String, String)
        Dim connectionstring = getConnectionString()
        Dim connectionstrings() As String = connectionstring.Split(";")
        For i = 0 To (connectionstrings.Length - 1)
            Dim mystrs() As String = connectionstrings(i).Split("=")
            _ConnectionStringDict.Add(mystrs(0), mystrs(1))
        Next i

    End Sub

    Private Function getConnectionString() As String
        _userid = "admin"
        _password = "admin"
        Dim builder As New NpgsqlConnectionStringBuilder()
        builder.ConnectionString = My.Settings.Connectionstring1
        builder.Add("User Id", _userid)
        builder.Add("password", _password)
        'builder.Add("CommandTimeout", "300")
        'builder.Add("TimeOut", "300")
        Return builder.ConnectionString
    End Function

    Public Function getConnection() As NpgsqlConnection
        If IsNothing(_userid) Or IsNothing(_password) Then
            Throw New DbAdapterExeption("User Id or Password is blank.")
        End If
        Return New NpgsqlConnection(_connectionString)
    End Function

    Public Function getDbDataAdapter() As NpgsqlDataAdapter
        Return New NpgsqlDataAdapter
    End Function

    Public Function getCommandObject() As NpgsqlCommand
        Return New NpgsqlCommand
    End Function

    Public Function getCommandObject(ByVal sqlstr As String, ByVal connection As Object) As NpgsqlCommand
        Return New NpgsqlCommand(sqlstr, connection)
    End Function

#Region "GetDataSet"
    Public Overloads Function TbgetDataSet(ByVal sqlstr As String, ByRef DataSet As DataSet, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter

        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                'DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                'DataAdapter.MissingSchemaAction = MissingSchemaAction.Add
                DataAdapter.Fill(DataSet)
            End Using
            myret = True
        Catch ex As NpgsqlException
            Dim obj = TryCast(ex.Errors(0), NpgsqlError)
            Dim myerror As String = String.Empty
            If Not IsNothing(obj) Then
                myerror = obj.InternalQuery
            End If
            message = ex.Message & " " & myerror
        End Try
        Return myret
    End Function

    Public Overloads Function TbgetDataSet(ByVal sqlstr As String, ByRef DataSet As DataSet, ByVal params As NpgsqlParameter(), Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter

        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Using cmd As NpgsqlCommand = New NpgsqlCommand()
                    cmd.CommandText = sqlstr
                    cmd.Connection = conn
                    DataAdapter.SelectCommand = cmd
                    If params.Length > 0 Then
                        cmd.Parameters.AddRange(params)
                    End If
                    DataAdapter.Fill(DataSet)
                    myret = True
                End Using
            End Using
        Catch ex As NpgsqlException
            Dim obj = TryCast(ex.Errors(0), NpgsqlError)
            Dim myerror As String = String.Empty
            If Not IsNothing(obj) Then
                myerror = obj.InternalQuery
            End If
            Message = ex.Message & " " & myerror
        End Try
        Return myret
    End Function

#End Region
    Function TbgetDataTable(ByRef sqlstr As String, ByRef DT As DataTable, ByRef mymessage As String) As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter

        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.Fill(DT)
            End Using
            myret = True
        Catch ex As NpgsqlException
            Dim obj = TryCast(ex.Errors(0), NpgsqlError)
            Dim myerror As String = String.Empty
            If Not IsNothing(obj) Then
                myerror = obj.InternalQuery
            End If
            mymessage = ex.Message & " " & myerror
        End Try
        Return myret
    End Function

    Function TBScorecardDataAdapter(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "sp_getscorecard() as tb(scorecardid bigint,supplierid bigint,mydate date,deptid integer,category integer,myvalue numeric)"
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    'DataAdapter.MissingMappingAction = MissingSchemaAction.AddWithKey

                    'DataAdapter.SelectCommand.Parameters.Add("col1", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dbtools1.Region
                    DataAdapter.Fill(DataSet)

                    'Delete
                    sqlstr = "sp_deletescorecard"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "scorecardid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    'Update
                    sqlstr = "sp_updatescorecard"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "scorecardid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "supplierid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "mydate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "category").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "myvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertscorecard"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "supplierid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "mydate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "category").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "myvalue").SourceVersion = DataRowVersion.Current
                    'DataAdapter.InsertCommand.Parameters.Add("paramhdid", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = paramhdid
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    RecordAffected = DataAdapter.Update(DataSet.Tables(0))

                Catch ex As Exception
                    message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                If Not IsNothing(_ConnectionStringDict) Then
                    _ConnectionStringDict.Clear()
                    _ConnectionStringDict = Nothing
                End If
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    

    Public Function copy(ByVal sqlstr As String, ByVal InputString As String, Optional ByRef result As Boolean = False) As String
        result = False
        Dim myReturn As String = ""
        'Convert string to MemoryStream
        Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.ASCII.GetBytes(InputString))
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.Default.GetBytes(InputString))
        Dim buf(9) As Byte
        Dim CopyInStream As Stream = Nothing
        Dim i As Long
        Using conn = New NpgsqlConnection(getConnectionString())
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                CopyIn1 = New NpgsqlCopyIn(command, conn)
                Try
                    CopyIn1.Start()
                    CopyInStream = CopyIn1.CopyStream
                    i = MemoryStream1.Read(buf, 0, buf.Length)
                    While i > 0
                        CopyInStream.Write(buf, 0, i)
                        i = MemoryStream1.Read(buf, 0, buf.Length)
                        Application.DoEvents()
                    End While
                    CopyInStream.Close()
                    result = True
                Catch ex As NpgsqlException
                    Try
                        CopyIn1.Cancel("Undo Copy")
                        myReturn = ex.Message & ", " & ex.Detail
                    Catch ex2 As NpgsqlException
                        If ex2.Message.Contains("Undo Copy") Then
                            myReturn = ex2.Message & ", " & ex2.Detail
                        End If
                    End Try
                End Try

            End Using
        End Using

        Return myReturn
    End Function

   
    'Public Function validint(ByVal sinqqty As String) As Object

    '    If sinqqty = "" Then
    '        Return DBNull.Value
    '    Else
    '        Return CInt(sinqqty.Replace(",", "").Replace("""", ""))
    '    End If
    'End Function
    Public Function validint(ByVal sinqqty As String) As Object

        If sinqqty = "" Then
            Return DBNull.Value
        Else
            Return CInt(sinqqty.Replace(",", "").Replace("""", ""))
        End If
    End Function
    Public Function validbool(ByVal mybool As String) As Object
        If mybool = "Y" Then
            Return "True"
        Else
            Return "False"
        End If
    End Function
    Public Function validdec(ByVal sunitprice As String) As Object
        If sunitprice = "" Then
            Return DBNull.Value
        Else
            Return CDec(sunitprice.Replace(",", "").Replace("""", ""))
        End If
    End Function

    Public Function validlong(ByVal myvalue As String) As Object
        If myvalue = "" Then
            Return DBNull.Value
        Else
            Return CLng(myvalue)
        End If
    End Function
    Public Function validlongNull(ByVal myvalue As String) As Object
        If myvalue = "" Then
            Return "Null"
        Else
            Return CLng(myvalue)
        End If
    End Function
    Public Function validchar(ByVal updateby As String) As Object
        If updateby = "" Then
            'Return DBNull.Value
            Return ""
        Else
            Return Trim(updateby.Replace("'", "''").Replace("""", ""))
        End If
    End Function
    Public Function validcharNull(ByVal updateby As String) As Object
        If updateby = "" Then
            'Return DBNull.Value
            Return "Null"
        Else
            Return Trim(updateby.Replace("'", "''").Replace("""", ""))
        End If
    End Function
    Public Function CDateddMMyyyy(ByVal updatedate As String) As Object
        Dim mydata() As String
        If updatedate.Contains(".") Then
            mydata = updatedate.Split(".")
        Else
            mydata = updatedate.Split("/")
        End If

        If mydata.Length > 1 Then
            Return CDate(mydata(2) & "-" & mydata(1) & "-" & mydata(0))
        End If
        Return DBNull.Value
    End Function
    Public Function ddMMyyyytoyyyyMMdd(ByVal updatedate As String) As Object
        Dim mydata() As String
        If updatedate.Contains(".") Then
            mydata = updatedate.Split(".")
        Else
            mydata = updatedate.Split("/")
        End If

        If mydata.Length > 1 Then
            Return "'" & mydata(2) & "-" & mydata(1) & "-" & mydata(0) & "'"
        End If
        Return DBNull.Value
    End Function

   
    Public Function ExNonQuery(ByVal sqlstr As String) As Long
        Dim myRet As Long
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                myRet = command.ExecuteNonQuery
            End Using
        End Using
        Return myRet
    End Function

    Public Function ExecuteNonQuery(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteNonQuery
                    'recordAffected = command.ExecuteNonQuery
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteNonQueryAsync(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteNonQuery

                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function

    Public Function ExecuteScalar(ByVal sqlstr As String, ByRef result As Object, ByRef param() As NpgsqlParameter, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    If param.Length > 0 Then
                        command.Parameters.AddRange(param)
                    End If
                    result = command.ExecuteScalar
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteScalar(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteScalar
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteScalar(ByVal sqlstr As String, ByRef recordAffected As Object, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteScalar
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function

    Sub ExecuteStoreProcedure(ByVal storeprocedurename As String)
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand(storeprocedurename, conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.ExecuteScalar()
            Catch ex As Exception

            End Try

        End Using
    End Sub

    Sub ExecuteStoreProcedure(ByVal storeprocedurename As String, ByRef param() As NpgsqlParameter)
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand(storeprocedurename, conn)
                cmd.CommandType = CommandType.StoredProcedure
                If param.Length > 0 Then
                    cmd.Parameters.AddRange(param)
                End If

                cmd.ExecuteScalar()
            Catch ex As Exception

            End Try

        End Using
    End Sub

    Public Function getproglock(ByVal programname As String, ByVal userid As String, ByVal status As Integer) As Boolean
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("proglock", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = programname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = status
            result = cmd.ExecuteScalar
        End Using
       
        Return result
    End Function

    Function dateformatdot(ByVal myrecord As String) As Object
        Dim myreturn = "Null"
        If myrecord = "" Then
            Return myreturn
        End If
        Dim mysplit = Split(myrecord, ".")
        myreturn = "'" & mysplit(2) & "-" & mysplit(1) & "-" & mysplit(0) & "'"
        Return myreturn
    End Function

    Function dateformatdotdate(ByVal myrecord As String) As Date
        Dim myreturn = "Null"
        If myrecord = "" Then
            Return myreturn
        End If
        Dim mysplit = Split(myrecord, ".")
        myreturn = CDate(mysplit(2) & "-" & mysplit(1) & "-" & mysplit(0))
        Return myreturn
    End Function

    Function dateformatYYYYMMdd(ByVal myrecord As Object) As Object
        Dim myreturn = "Null"

        myreturn = "'" & CDate(myrecord).Year & "-" & CDate(myrecord).Month & "-" & CDate(myrecord).Day & "'"
        Return myreturn
    End Function

    'Function ImportTx(ByVal formImportSaving As FormImportSaving, ByVal mye As ContentBaseEventArgs) As Boolean
    '    Dim sqlstr As String = String.Empty
    '    Dim DataAdapter As New NpgsqlDataAdapter
    '    'Dim cmd As NpgsqlCommand
    '    Dim myret As Boolean = False

    '    Try
    '        Using conn As New NpgsqlConnection(Connectionstring)
    '            Try
    '                conn.Open()
    '                myTransaction = conn.BeginTransaction
    '                'Update
    '                sqlstr = "sp_updatesavinglookup"
    '                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "savinglookupid").SourceVersion = DataRowVersion.Original
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "savinglookupname").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "parentid").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

    '                'insert
    '                sqlstr = "sp_insertsavinglookup"
    '                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "savinglookupname").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "parentid").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

    '                DataAdapter.UpdateCommand.Transaction = mytransaction
    '                DataAdapter.InsertCommand.Transaction = mytransaction

    '                mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

    '                sqlstr = "sp_updatesaving"
    '                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "savingid").SourceVersion = DataRowVersion.Original
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "mytotal").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "enddate").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

    '                'insert
    '                sqlstr = "sp_insertsaving"
    '                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "actionid").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "mytotal").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "startdate").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "enddate").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

    '                DataAdapter.UpdateCommand.Transaction = mytransaction
    '                DataAdapter.InsertCommand.Transaction = mytransaction

    '                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

    '                mytransaction.Commit()
    '                myret = True
    '            Catch ex As Exception
    '                mytransaction.Rollback()
    '                mye.message = ex.Message
    '                Return False
    '            End Try
    '        End Using

    '    Catch ex As NpgsqlException
    '        mye.message = ex.Message
    '    End Try
    '    Return myret

    'End Function
    Function copyToPriceChange(ByVal creator As String, ByRef dr As DataRow, ByVal isnewrecord As Boolean, ByRef message As String) As Boolean
        Dim myret As Boolean
        Dim result As Object
        Dim myparam As NpgsqlParameter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_copypricechange", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = creator
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0).Value = isnewrecord
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dr.Item("validator1")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dr.Item("validator2")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dr.Item("pricetype")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dr.Item("description")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = dr.Item("submitdate")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = dr.Item("negotiateddate")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dr.Item("attachment")
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = dr.Item("reasonid")

                myparam = cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0)
                myparam.Value = dr.Item("pricechangehdid")
                myparam.Direction = ParameterDirection.InputOutput


                result = cmd.ExecuteScalar
                If IsDBNull(result) Then
                    myret = False
                End If
                myret = True
            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            message = ex.Message & ". " & errordetail
        End Try
        Return myret
    End Function

    'Function PriceChangeTx(ByVal formPriceChange As FormPriceChange, ByVal mye As ContentBaseEventArgs) As Boolean
    '    Dim sqlstr As String = String.Empty
    '    Dim DataAdapter As New NpgsqlDataAdapter
    '    'Dim cmd As NpgsqlCommand
    '    Dim myret As Boolean = False

    '    Try
    '        Using conn As New NpgsqlConnection(Connectionstring)
    '            Try
    '                conn.Open()
    '                mytransaction = conn.BeginTransaction
    '                'Update
    '                sqlstr = "sp_updatepricechangehd"
    '                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "creator").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator1").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator2").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pricetype").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "submitdate").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "negotiateddate").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "attachment").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "actiondate").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionby").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "reasonid").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

    '                sqlstr = "sp_insertpricechangehd"
    '                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "creator").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator1").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator2").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pricetype").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "submitdate").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "negotiateddate").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "attachment").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "actiondate").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionby").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "reasonid").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").Direction = ParameterDirection.InputOutput
    '                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


    '                DataAdapter.InsertCommand.Transaction = mytransaction
    '                DataAdapter.UpdateCommand.Transaction = mytransaction

    '                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

    '                sqlstr = "sp_updatepricechangedtl"
    '                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "purchorg").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plant").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validon").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "price").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricingunit").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

    '                'insert
    '                sqlstr = "sp_insertpricechangedtl"
    '                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "purchorg").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plant").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validon").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "price").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricingunit").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


    '                DataAdapter.UpdateCommand.Transaction = mytransaction
    '                DataAdapter.InsertCommand.Transaction = mytransaction


    '                sqlstr = "sp_deletepricechangedtl"
    '                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
    '                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


    '                mye.ra = DataAdapter.Update(mye.dataset.Tables(4))

    '                mytransaction.Commit()
    '                myret = True
    '                'Catch ex As Exception
    '                '    Dim myerr = CType(ex, NpgsqlException)
    '                '    mytransaction.Rollback()
    '                '    mye.message = ex.Message & " " & myerr.Detail
    '                '    Return False
    '                'End Try
    '            Catch ex As NpgsqlException
    '                Dim errordetail As String = String.Empty
    '                errordetail = "" & ex.Detail
    '                mye.message = ex.Message & ". " & errordetail
    '                Return False
    '            End Try
    '        End Using

    '    Catch ex As NpgsqlException
    '        mye.message = ex.Message
    '    End Try
    '    Return myret
    'End Function

    'Function PriceCommentTx(ByVal pricechangehdid As Long, ByRef message As String) As Boolean
    '    Dim sqlstr As String = String.Empty
    '    Dim DataAdapter As New NpgsqlDataAdapter
    '    Dim result As Object
    '    Dim myret As Boolean = False

    '    Try
    '        Using conn As New NpgsqlConnection(Connectionstring)
    '            conn.Open()
    '            Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_addpricecomment", conn)
    '            cmd.CommandType = CommandType.StoredProcedure
    '            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0).Value = pricechangehdid

    '            result = cmd.ExecuteScalar
    '            If IsDBNull(result) Then
    '                myret = False
    '            End If
    '            myret = True
    '        End Using

    '    Catch ex As NpgsqlException
    '        Message = ex.Message
    '    End Try
    '    Return myret
    'End Function

    'Function PriceChangeTx(ByVal formPriceChange As FormMyTask, ByVal mye As ContentBaseEventArgs) As Boolean
    '    Dim sqlstr As String = String.Empty
    '    Dim DataAdapter As New NpgsqlDataAdapter
    '    'Dim cmd As NpgsqlCommand
    '    Dim myret As Boolean = False

    '    Try
    '        Using conn As New NpgsqlConnection(Connectionstring)
    '            Try
    '                conn.Open()
    '                mytransaction = conn.BeginTransaction
    '                'Update
    '                sqlstr = "sp_updatepricechangehd"
    '                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "creator").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator1").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator2").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pricetype").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "description").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "submitdate").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "negotiateddate").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "attachment").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "actiondate").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionby").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "reasonid").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


    '                DataAdapter.UpdateCommand.Transaction = mytransaction


    '                mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

    '                'sqlstr = "sp_updatepricechangedtl"
    '                'DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
    '                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
    '                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
    '                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
    '                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
    '                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "purchorg").SourceVersion = DataRowVersion.Current
    '                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plant").SourceVersion = DataRowVersion.Current
    '                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validon").SourceVersion = DataRowVersion.Current
    '                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "price").SourceVersion = DataRowVersion.Current
    '                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricingunit").SourceVersion = DataRowVersion.Current
    '                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
    '                'DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure



    '                'DataAdapter.UpdateCommand.Transaction = mytransaction


    '                'mye.ra = DataAdapter.Update(mye.dataset.Tables(4))

    '                mytransaction.Commit()
    '                myret = True
    '            Catch ex As Exception
    '                mytransaction.Rollback()
    '                mye.message = ex.Message
    '                Return False
    '            End Try
    '        End Using

    '    Catch ex As NpgsqlException
    '        mye.message = ex.Message
    '    End Try
    '    Return myret
    'End Function
    'Function PriceChangeReasonTx(ByVal formMasterProduct As FormPriceChangeReasonMaster, ByVal mye As ContentBaseEventArgs) As Boolean
    '    Dim sqlstr As String = String.Empty
    '    Dim DataAdapter As New NpgsqlDataAdapter
    '    Dim myret As Boolean = False

    '    Try
    '        Using conn As New NpgsqlConnection(Connectionstring)
    '            Try
    '                conn.Open()

    '                'Family
    '                'Insert
    '                sqlstr = "sp_insertpricechangereason"
    '                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "reasonname").SourceVersion = DataRowVersion.Current
    '                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
    '                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


    '                'Update
    '                sqlstr = "sp_updatepricechangereason"
    '                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
    '                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "reasonname").SourceVersion = DataRowVersion.Current
    '                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


    '                sqlstr = "sp_deletepricechangereason"
    '                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
    '                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
    '                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

    '                mytransaction = conn.BeginTransaction
    '                DataAdapter.DeleteCommand.Transaction = mytransaction
    '                DataAdapter.UpdateCommand.Transaction = mytransaction
    '                DataAdapter.InsertCommand.Transaction = mytransaction

    '                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

    '                mytransaction.Commit()
    '            Catch ex As Exception
    '                mytransaction.Rollback()
    '                mye.message = ex.Message
    '                Return False
    '            End Try
    '        End Using
    '        myret = True
    '    Catch ex As NpgsqlException
    '        mye.message = ex.Message
    '    End Try
    '    Return myret
    'End Function

    Function PriceChangeSendEmail(ByVal roleTasks As RoleTasks, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    'Update
                    sqlstr = "sp_updatepricechangehdemailsend"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sendstdvalidatedtocreator").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sendcompletedtocreator").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sendtocc").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function PriceChangeExportFile(ByVal exportSAPClass As ExportSAPClass, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    'Update
                    sqlstr = "sp_updatepricechangehdexportfile"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "exportfileid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "exportfiledate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function PriceChangedtlsap(ByVal validateSAPPrice As ValidateSAPPrice, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    'Update
                    sqlstr = "sp_updatepricechangedtlsap"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sap").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function PriceChangehdcompleted(ByVal validateSAPPrice As ValidateSAPPrice, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    'Update
                    sqlstr = "sp_updatepricechangehdcompleted"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pricechangehdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function PriceChangeDTLTx(ByVal formHistoryDetail As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()


                    sqlstr = "sp_deletepricechangedtl"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pricechangedtlid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.DeleteCommand.Transaction = mytransaction

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function DocumentVendorTxOld(ByVal formDocumentVendor As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatedocumenthd"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc4").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "otheremail").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "creationdate").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertdocumenthd"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc4").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "otheremail").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "creationdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletedocumenthd"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                sqlstr = "doc.sp_updatedocumentdtl"

                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                'vendordoc
                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vdid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vdid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "headerid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
                'document
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "docdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docext").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "uploaddate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "remarks").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doctypeid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doclevelid").SourceVersion = DataRowVersion.Current
                'version
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "version").SourceVersion = DataRowVersion.Current
                'generalcontract
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paymentcode").SourceVersion = DataRowVersion.Current
                'supplychain
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "leadtime").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "sasl").SourceVersion = DataRowVersion.Current
                'quality
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nqsu").SourceVersion = DataRowVersion.Current
                'project
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectname").SourceVersion = DataRowVersion.Current
                'socialaudit
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "auditby").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "audittype").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "auditgrade").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "overallauditresult").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "auditscore").SourceVersion = DataRowVersion.Current
                'sef
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "score").SourceVersion = DataRowVersion.Current
                'sif
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myyear").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery4").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy4").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "expireddate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                'insert
                sqlstr = "doc.sp_insertdocumentdtl"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                'vendordoc

                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "headerid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                'DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
                'document
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "docdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docext").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "uploaddate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "remarks").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doctypeid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doclevelid").SourceVersion = DataRowVersion.Current
                'version
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "version").SourceVersion = DataRowVersion.Current
                'generalcontract
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paymentcode").SourceVersion = DataRowVersion.Current
                'supplychain
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "leadtime").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "sasl").SourceVersion = DataRowVersion.Current
                'quality
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nqsu").SourceVersion = DataRowVersion.Current
                'project
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectname").SourceVersion = DataRowVersion.Current
                'socialaudit
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "auditby").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "audittype").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "auditgrade").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "overallauditresult").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "auditscore").SourceVersion = DataRowVersion.Current

                'sef
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "score").SourceVersion = DataRowVersion.Current
                'sif
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myyear").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery4").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy4").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "expireddate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").Direction = ParameterDirection.InputOutput
                'DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vdid").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletedocumentdtl"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                'DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vdid").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(1))


                sqlstr = "doc.sp_deletesitx"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertsitx"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "labelid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "value").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updatesitx"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "labelid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "value").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction



                mye.ra = DataAdapter.Update(mye.dataset.Tables(10))


                'Zetol
                sqlstr = "doc.sp_deletezetol"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertzetol"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "zetolid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updatezetol"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "zetolid").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(12))

                'GeneralOther
                sqlstr = "doc.sp_deletegeneralother"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertgeneralother"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paymentcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "otherdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updategeneralother"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paymentcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "otherdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(13))

                'SupplyOther
                sqlstr = "doc.sp_deletesupplychainother"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertsupplychainother"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "leadtime").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "sasl").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "otherdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updatesupplychainother"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "leadtime").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "sasl").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "otherdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(14))

                'QualityOther
                sqlstr = "doc.sp_deletequalityother"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertqualityother"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nqsu").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "otherdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updatequalityother"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nqsu").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "otherdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(15))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function DocumentVendorTx(ByVal formDocumentVendor As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatedocumenthd"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc4").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "otheremail").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "creationdate").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertdocumenthd"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "validator").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cc4").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "otheremail").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "creationdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletedocumenthd"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                sqlstr = "doc.sp_updatedocumentdtl"

                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                'vendordoc
                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vdid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vdid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "headerid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
                'document
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "docdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docext").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "uploaddate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "remarks").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doctypeid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doclevelid").SourceVersion = DataRowVersion.Current
                'version
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "version").SourceVersion = DataRowVersion.Current
                'generalcontract
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paymentcode").SourceVersion = DataRowVersion.Current
                'supplychain
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "leadtime").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "sasl").SourceVersion = DataRowVersion.Current
                'quality
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nqsu").SourceVersion = DataRowVersion.Current
                'project
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectname").SourceVersion = DataRowVersion.Current
                'socialaudit
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "auditby").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "audittype").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "auditgrade").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "overallauditresult").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "auditscore").SourceVersion = DataRowVersion.Current
                'sef
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "score").SourceVersion = DataRowVersion.Current
                'sif
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myyear").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery4").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy4").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "expireddate").SourceVersion = DataRowVersion.Current

                'BuyerInput
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "producttype").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "lastvisit").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "keystakestopic").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "strength1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "strength2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "strength3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "weaknessess1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "weaknessess2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "weaknessess3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "opportunities1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "opportunities2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "opportunities3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "threats1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "threats2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "threats3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myyear2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "currbudget").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "totalbudget").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "productdevelopment1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "productdevelopment2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "negotiationresult1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "negotiationresult2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "negotiationresult3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "outstandingissue1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "outstandingissue2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "outstandingissue3").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "outstandingissue4").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "outstandingissue5").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "keycontactperson").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "title").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "top3cust1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "top3cust2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "top3cust3").SourceVersion = DataRowVersion.Current
                'projectspecification
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "returnrate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                'insert
                sqlstr = "doc.sp_insertdocumentdtl"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                'vendordoc

                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "headerid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                'DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
                'document
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "docdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docext").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "uploaddate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "remarks").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doctypeid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doclevelid").SourceVersion = DataRowVersion.Current
                'version
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "version").SourceVersion = DataRowVersion.Current
                'generalcontract
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paymentcode").SourceVersion = DataRowVersion.Current
                'supplychain
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "leadtime").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "sasl").SourceVersion = DataRowVersion.Current
                'quality
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nqsu").SourceVersion = DataRowVersion.Current
                'project
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectname").SourceVersion = DataRowVersion.Current
                'socialaudit
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "auditby").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "audittype").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "auditgrade").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "overallauditresult").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "auditscore").SourceVersion = DataRowVersion.Current

                'sef
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "score").SourceVersion = DataRowVersion.Current
                'sif
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myyear").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "turnovery4").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "ratioy4").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "expireddate").SourceVersion = DataRowVersion.Current
                

                'BuyerInput
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "producttype").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "lastvisit").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "keystakestopic").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "strength1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "strength2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "strength3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "weaknessess1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "weaknessess2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "weaknessess3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "opportunities1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "opportunities2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "opportunities3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "threats1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "threats2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "threats3").SourceVersion = DataRowVersion.Current                
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myyear2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "currbudget").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "totalbudget").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "productdevelopment1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "productdevelopment2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "negotiationresult1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "negotiationresult2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "negotiationresult3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "outstandingissue1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "outstandingissue2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "outstandingissue3").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "outstandingissue4").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "outstandingissue5").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "keycontactperson").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "title").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "top3cust1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "top3cust2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "top3cust3").SourceVersion = DataRowVersion.Current
                'projectspecification
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "returnrate").SourceVersion = DataRowVersion.Current

                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vdid").Direction = ParameterDirection.InputOutput

               

                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure



                sqlstr = "doc.sp_deletedocumentdtl"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                'DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vdid").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(1))


                sqlstr = "doc.sp_deletesitx"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertsitx"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "labelid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "value").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updatesitx"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "labelid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "value").SourceVersion = DataRowVersion.Current


                

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction



                mye.ra = DataAdapter.Update(mye.dataset.Tables(10))


                'Zetol
                sqlstr = "doc.sp_deletezetol"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertzetol"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "zetolid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updatezetol"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "zetolid").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(12))

                'GeneralOther
                sqlstr = "doc.sp_deletegeneralother"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertgeneralother"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paymentcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "otherdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updategeneralother"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paymentcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "otherdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(13))

                'SupplyOther
                sqlstr = "doc.sp_deletesupplychainother"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertsupplychainother"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "leadtime").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "sasl").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "otherdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updatesupplychainother"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "leadtime").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "sasl").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "otherdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(14))

                'QualityOther
                sqlstr = "doc.sp_deletequalityother"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertqualityother"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nqsu").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "otherdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updatequalityother"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nqsu").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "otherdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(15))


                'ActionPlan
                sqlstr = "doc.sp_deleteactionplan"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertactionplan"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "priority").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "situation").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "target").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "proposal").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "responsibleperson").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "startdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "enddate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "result").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "finishdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "uploaddate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updateactionplan"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "documentid").SourceVersion = DataRowVersion.Current                
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "priority").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "situation").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "target").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "proposal").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "responsibleperson").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "startdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "enddate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "result").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "finishdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "status").SourceVersion = DataRowVersion.Current
                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "actionid").SourceVersion = DataRowVersion.Current
                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "uploaddate").SourceVersion = DataRowVersion.Current
                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current


                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(16))


                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function SupplierCategoryTx(ByVal formSupplierCategory As FormSupplierCategory, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatesuppliercategory"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "supplierscategoryid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "category").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myorder").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertsuppliercategory"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "category").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myorder").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "suppliercategoryid").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletesuppliercategory"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "supplierscategoryid").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function PanelStatusTx(ByVal formPanelStatus As FormPanelStatus, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatepanelstatus"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "panelstatus").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "paneldescription").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertpanelstatus"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "panelstatus").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "paneldescription").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletepanelstatus"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function VendorStatusTx(ByVal formVendorStatus As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatevendorstatus"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "producttypeid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "rank").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remark").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = HelperClass1.UserInfo.DisplayName

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertvendorstatus"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "producttypeid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "rank").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remark").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = HelperClass1.UserInfo.DisplayName
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletevendorstatus"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function PanelHistoryTx(ByVal formPanelHistory As FormPanelHistory, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatepanelhistory"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validfrom").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "suppliercategoryid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "fp").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "cp").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "rank").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = HelperClass1.UserInfo.DisplayName
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertpanelhistory"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "validfrom").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "suppliercategoryid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "fp").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "cp").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "rank").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = HelperClass1.UserInfo.DisplayName
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletepanelhistory"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function
    Public Sub onRowUpdate(ByVal sender As Object, ByVal e As NpgsqlRowUpdatedEventArgs)
        If e.StatementType = StatementType.Insert Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If

        End If
    End Sub

    Public Sub onRowInsertUpdate(ByVal sender As Object, ByVal e As NpgsqlRowUpdatedEventArgs)
        If e.StatementType = StatementType.Insert Or e.StatementType = StatementType.Update Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If
        End If
    End Sub

    Function SupplierPanelTx(ByVal formSupplierPanel As FormSupplierPanel, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatesupplierpanel"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "supplierspanelid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current                
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "supplierscategoryid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "fp").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "cp").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "rank").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = HelperClass1.UserInfo.DisplayName
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertsupplierpanel"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current                
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "supplierscategoryid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "fp").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "cp").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "rank").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = HelperClass1.UserInfo.DisplayName
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "supplierspanelid").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletesupplierpanel"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "supplierspanelid").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function StatusTx(ByVal formMasterStatus As FormMasterStatus, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatestatus"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "statusid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "status").SourceVersion = DataRowVersion.Current
               
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertstatus"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "statusid").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletestatus"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "statusid").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function UserTx(ByVal formUser As FormUser, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updateuser"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "username").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isadmin").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isfinance").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "allowupdatedocument").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "newemail").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isguest").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertuser"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "username").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isadmin").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isfinance").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "allowupdatedocument").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isguest").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deleteuser"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function IsAdmin(ByVal p1 As String) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim Command As New NpgsqlCommand
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Command = New NpgsqlCommand("doc.sp_isadmin", conn)
                Command.CommandType = CommandType.StoredProcedure
                Command.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").Value = p1.ToLower
                Dim result As Object = Command.ExecuteScalar
                If Not IsNothing(result) Then
                    myret = result
                End If
            End Using
        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            Return False
        End Try
        Return myret
    End Function

    Function isOfficer(ByVal p1 As String) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim Command As New NpgsqlCommand
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Command = New NpgsqlCommand("doc.sp_isofficer", conn)
                Command.CommandType = CommandType.StoredProcedure
                Command.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").Value = p1.ToLower
                Dim result As Object = Command.ExecuteScalar
                If Not IsNothing(result) Then
                    myret = result
                End If
            End Using
        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            Return False
        End Try
        Return myret
    End Function


    Function getLevel(ByVal p1 As String) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim Command As New NpgsqlCommand
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Command = New NpgsqlCommand("doc.sp_getlevel", conn)
                Command.CommandType = CommandType.StoredProcedure
                Command.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").Value = p1.ToLower
                Dim result As Object = Command.ExecuteScalar
                If Not IsNothing(result) Then
                    myret = result
                End If
            End Using
        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            Return False
        End Try
        Return myret
    End Function

    Function GroupTx(ByVal formMasterGroup As FormMasterGroup, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updategroupauth"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "groupid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "groupname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "newvendorcreation").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertgroupauth"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "groupname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "newvendorcreation").Direction = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "groupid").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletegroupauth"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "groupid").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function GroupVendorTx(ByVal formGroupAssign As FormGroupVendor, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updategroupvendor"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "groupid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertgroupvendor"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "groupid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletegroupvendor"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function GroupUserTx(ByVal formGroupUser As FormGroupUser, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updategroupuser"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "groupid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertgroupuser"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "groupid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletegroupuser"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function AllowUpdateDocument(ByVal p1 As String) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim Command As New NpgsqlCommand
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Command = New NpgsqlCommand("doc.sp_allowupdatedocument", conn)
                Command.CommandType = CommandType.StoredProcedure
                Command.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").Value = p1.ToLower
                Dim result As Object = Command.ExecuteScalar
                If Not IsNothing(result) Then
                    myret = result
                End If
            End Using
        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            Return False
        End Try
        Return myret
    End Function

    Function VendorTx(ByVal formMasterVendor As FormMasterVendor, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "sp_updatevendor"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "vendorname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "ssmid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pmid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "ssmeffectivedate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "pmeffectivedate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "sp_insertvendor"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "vendorname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "ssmid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pmid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "ssmeffectivedate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "pmeffectivedate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "sp_deletevendor"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function AssetPurchaseTx(ByVal formAssetsPurchase As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction

                'ToolingProject - Create only
                AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
                'DataAdapter.ContinueUpdateOnError = True
                sqlstr = "doc.sp_inserttoolingproject"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "dept").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "ppps").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbuid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure



                sqlstr = "doc.sp_inserttoolingproject"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "dept").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "ppps").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbuid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletetoolingproject"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables("ToolingProject"))

                'AssetPurchase
                ' AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
                sqlstr = "doc.sp_insertassetpurchase"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "projectid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "applicantname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "applicantdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "assetdescription").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "categoryofasset").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "typeofinvestment").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvalname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "budgetamount").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "aeb").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "investmentorderno").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "toolingpono").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "financeassetno").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "reason").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "comments").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "exchangerate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remarks").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "amortperiod_1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "amortqty_1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "totalamortqty_1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "totalamortamount_1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "amortperunit_1").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "amortperiod_2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "amortqty_2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "totalamortqty_2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "totalamortamount_2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "amortperunit_2").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "budgetcurr").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "creator").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "sapcapdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paymentmethodid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "amortcurr").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "amortexrate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "amortremarks").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvalname2").SourceVersion = DataRowVersion.Current

                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "paymententity").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "toolingsupplier").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "expectedtoolingfinishdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "origin").SourceVersion = DataRowVersion.Current

                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "assetpurchaseid").Direction = ParameterDirection.InputOutput

                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                sqlstr = "doc.sp_updateassetpurchase"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "projectid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "applicantname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "applicantdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "assetdescription").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "categoryofasset").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "typeofinvestment").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvalname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "budgetamount").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "aeb").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "investmentorderno").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "toolingpono").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "financeassetno").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "reason").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "comments").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "exchangerate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remarks").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "amortperiod_1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "amortqty_1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "totalamortqty_1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "totalamortamount_1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "amortperunit_1").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "amortperiod_2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "amortqty_2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "totalamortqty_2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "totalamortamount_2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "amortperunit_2").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "budgetcurr").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "sapcapdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "paymentmethodid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "amortcurr").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "amortexrate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "amortremarks").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "approvalname2").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "paymententity").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "toolingsupplier").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "expectedtoolingfinishdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "origin").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "printdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "assetpurchaseid").Direction = ParameterDirection.InputOutput
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                sqlstr = "doc.sp_deletassetpurchase"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables("AssetPurchase"))

                'Tooling List
                AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
                sqlstr = "doc.sp_inserttoolinglist"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "suppliermoldno").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "toolsdescription").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "material").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cavities").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "numberoftools").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "dailycaps").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "cost").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "comments").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineno").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "assetpurchaseid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sebmodelno").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "suppliermodelreference").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "purchasedate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "location").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "toolinglistid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "typeofinvestment").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "commontool").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "originalcurrency").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "originalcost").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure



                sqlstr = "doc.sp_inserttoolinglist"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "suppliermoldno").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "toolsdescription").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "material").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cavities").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "numberoftools").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "dailycaps").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "cost").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "comments").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineno").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "assetpurchaseid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sebmodelno").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "suppliermodelreference").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "purchasedate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "location").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "toolinglistid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "typeofinvestment").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "commontool").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "originalcurrency").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "originalcost").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                sqlstr = "doc.sp_deletetoolinglist"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables("ToolingListDT"))


                'Invoice
                AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
                sqlstr = "doc.sp_inserttoolinginvoice"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "invoiceno").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "invoicedate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "proformainvoice").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "currency").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "amount").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updatetoolinginvoice"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "invoiceno").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "invoicedate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "proformainvoice").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "currency").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "amount").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                sqlstr = "doc.sp_deletetoolinginvoice"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure



                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables("ToolingInvoice"))


                'Payment
                AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
                sqlstr = "doc.sp_inserttoolingpayment"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "toolinglistid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "invoiceid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "invoiceamount").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "currency").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "exrate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updatetoolingpayment"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "toolinglistid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "invoiceid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "invoiceamount").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "currency").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "exrate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                sqlstr = "doc.sp_deletetoolingpayment"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure



                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables("ToolingPayment"))



                'Attachment
                AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
                sqlstr = "doc.sp_insertattachment"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "docdate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docext").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doctypeid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "assetpurchaseid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remarks").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updateattachment"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "docdate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docext").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "doctypeid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "assetpurchaseid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remarks").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                sqlstr = "doc.sp_deleteattachment"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables("Attachment"))


                'Tracking No
                AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
                sqlstr = "doc.sp_inserttrackingno"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "trackingno").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "assetpurchaseid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_updatetrackingno"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "trackingno").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "assetpurchaseid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                sqlstr = "doc.sp_deletetrackingno"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables("TrackingNo"))


                'AssetPurchaseAction
                AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
                sqlstr = "doc.sp_insertassetpurchaseaction"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "assetpurchaseid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remark").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure
                mye.ra = DataAdapter.Update(mye.dataset.Tables("AssetPurchaseAction"))


                'ProformaInvoice
                sqlstr = "doc.sp_insertproformainvoice"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "assetpurchaseid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "proformainvoice").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "creator").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "cretiondate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure
                mye.ra = DataAdapter.Update(mye.dataset.Tables("ProformaInvoice"))



                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function FamilyGroupSBUTx(ByVal formUpdateFamily As FormUpdateFamily, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_insertupdatefamilygroupsbu"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "groupingcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbusapid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isdeleted").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertupdatefamilygroupsbu"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "groupingcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbusapid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isdeleted").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletefamilygroupsbu"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function ToolingStatusTX(ByVal formUpdateToolingStatus As FormUpdateToolingStatus, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatetoolingstatus"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "toolinglistid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "toolingstatus").SourceVersion = DataRowVersion.Current                
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.UpdateCommand.Transaction = mytransaction
               
                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function MasterDocTypeTx(ByVal formMasterDocType As FormMasterDocType, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatedoctype"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "doctypename").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "groupid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "template").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "templateupdateby").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "reminder").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertdoctype"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "doctypename").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "groupid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "template").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "templateupdateby").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "reminder").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletedoctype"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction


                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function SPMSSMTx(ByVal formImportVendorSPMPM As FormImportVendorSPMPM, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatevendorspmpm"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "ssmid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pmid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "ssmeffectivedate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "pmeffectivedate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.UpdateCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))


                sqlstr = "doc.sp_insertvendorspmpm"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "spmid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "pmid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "spmeffectivedate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "pmeffectivedate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(5))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function isFinance(ByVal p1 As String) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim Command As New NpgsqlCommand
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Command = New NpgsqlCommand("doc.sp_isfinance", conn)
                Command.CommandType = CommandType.StoredProcedure
                Command.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").Value = p1.ToLower
                Dim result As Object = Command.ExecuteScalar
                If Not IsNothing(result) Then
                    myret = result
                End If
            End Using
        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            Return False
        End Try
        Return myret
    End Function

    Function GetGroupName(ByVal p1 As String, ByRef errormsg As String) As List(Of String)
        Dim sqlstr As String = String.Empty
        Dim myresult As New List(Of String)
        Dim Command As New NpgsqlCommand
        Dim GroupName As New System.Text.StringBuilder
        Dim remarks As New System.Text.StringBuilder
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                sqlstr = "select groupname from doc.groupuser gu left join doc.user u on u.id = gu.userid left join doc.groupauth g on g.groupid = gu.groupid where u.userid = :value1"
                Command = New NpgsqlCommand(sqlstr, conn)
                'Command.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").Value = p1.ToLower
                Command.Parameters.Add("value1", NpgsqlDbType.Varchar).Value = p1.ToLower
                'Command.Parameters.Add(New NpgsqlParameter("value1", NpgsqlDbType.Varchar))
                'Command.Parameters(0).Value = p1.ToLower
                Dim result As NpgsqlDataReader = Command.ExecuteReader
                While result.Read
                    If GroupName.Length > 0 Then
                        GroupName.Append(",")
                    End If
                    Select Case result.Item("groupname")
                        Case "Controlling Dept Tooling Amortization Approval", "Controlling Dept Tooling Amortization CC"
                            GroupName.Append("Controlling Dept")
                            If remarks.Length > 0 Then
                                remarks.Append(",")
                            End If
                            remarks.Append(String.Format("{0:d}", PaymentMethodIDEnum.Amortization))
                        Case "Controlling Dept Tooling Investment Approval", "Controlling Dept Tooling Investment CC"
                            GroupName.Append("Controlling Dept")
                            If remarks.Length > 0 Then
                                remarks.Append(",")
                            End If
                            remarks.Append(String.Format("{0:d}", PaymentMethodIDEnum.Investment))
                        Case Else
                            GroupName.Append(result.Item("groupname"))
                    End Select

                End While
                myresult.Add(GroupName.ToString)
                myresult.Add(remarks.ToString)
            End Using
        Catch ex As NpgsqlException           
            errormsg = "" & ex.Detail
            Return myresult
        End Try
        Return myresult
    End Function
    Public Function setaudit(ByVal programname As String, ByVal userid As String) As Boolean
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("doc.sp_importaudit", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = programname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid         
            result = cmd.ExecuteScalar
        End Using
        Return result
    End Function

    Function VendorFactoryContact(ByVal formFactoryAndContact As FormFactoryAndContact, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Factory HD
                sqlstr = "doc.sp_updatefactoryhd"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "customname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertfactoryhd"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "customname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletefactoryhd"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original                
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(2))


                'FactoryDtl
                sqlstr = "doc.sp_updatefactorydtl"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "factoryhdid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "chinesename").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "englishname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "chineseaddress").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "englishaddress").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "area").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "city").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "provinceid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "countryid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "main").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure
                DataAdapter.UpdateCommand.Transaction = mytransaction


                sqlstr = "doc.sp_insertfactorydtl"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "factoryhdid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "chinesename").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "englishname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "chineseaddress").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "englishaddress").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "area").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "city").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "provinceid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "countryid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "main").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletefactorydtl"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(1))


                'Vendor Factory

                sqlstr = "doc.sp_insertvendorfactory"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "factoryid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletevendorfactory"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "factoryid").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))


                'Contact             
                sqlstr = "doc.sp_updatecontact"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "contactname").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "title").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "officeph").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "factoryph").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "officemb").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "factorymb").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isecoqualitycontact").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure
                DataAdapter.UpdateCommand.Transaction = mytransaction


                sqlstr = "doc.sp_insertcontact"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)                
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "contactname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "title").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "officeph").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "factoryph").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "officemb").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "factorymb").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isecoqualitycontact").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletecontact"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                mye.ra = DataAdapter.Update(mye.dataset.Tables(4))


                'Vendor Contact
                sqlstr = "doc.sp_insertvendorcontact"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "contactid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletevendorcontact"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "contactid").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Environment.UserDomainName & "\" & Environment.UserName
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(3))
                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function CountryTx(ByVal formMasterCountry As Object, ByVal mye As Object) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatecountry"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "countryname").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertcountry"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "countryname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletecountry"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function ProvinceTx(ByVal formMasterCountry As Object, ByVal mye As Object) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updateprovince"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "provincename").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertprovince"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "provincename").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deleteprovince"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function loglogin(ByVal applicationname As String, ByVal userid As String, ByVal username As String, ByVal computername As String, ByVal time_stamp As Date)
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertlogonhistory", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = applicationname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = computername
            result = cmd.ExecuteNonQuery
        End Using
        Return result
    End Function

    Function TechnologyTx(ByVal formMasterTechnology As FormMasterTechnology, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatetechnology"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "technologyname").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_inserttechnology"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "technologyname").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletestatus"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function VendorIndirectFamilyTx(ByVal MyForm As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatevendorindirectfamily"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familycodeid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertvendorindirectfamily"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familycodeid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletevendorindirectfamily"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function VendorTechnologyTx(ByVal formVendorTechnology As FormVendorTechnology, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatevendortechnology"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "technologyid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineno").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertvendortechnology"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "technologyid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineno").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletevendortechnology"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function TaskDocumentForwardUser(ByVal formMyTaskDocument As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updateheader"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                DataAdapter.UpdateCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function SupplierPhotosTx(ByVal supplierPhotosAdapter As SupplierPhotosAdapter, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatesupplierphotos"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "phototype").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineorderfp").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineordercp").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineordergeneral").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "filename").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "fp").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "cp").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "general").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "createddate").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertsupplierphotos"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)

                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "phototype").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineorderfp").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineordercp").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineordergeneral").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "filename").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "fp").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "cp").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "general").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "createddate").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletesupplierphotos"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function DataTypeTx(ByVal dataTypeAdapter As DataTypeAdapter, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatedatatype"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "datatypename").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineorder").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "amountlineorder").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "qtylineorder").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "groupid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertdatatype"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "datatypename").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineorder").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "amountlineorder").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "qtylineorder").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "groupid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletedatatype"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function DataTypeLabelTx(ByVal dataTypeLabelAdapter As DataTypeLabelAdapter, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatedatatypelabel"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "datatypelabel").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineorder").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "datatypelabelid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertdatatypelabel"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "datatypelabel").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lineorder").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "datatypelabelid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletedatatypelabel"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Function AssetPurchaseSendEmail(ByVal roleTasks As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    'Update
                    sqlstr = "doc.sp_updateassetpurhcasesendemail"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sendcomplete").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    mytransaction = conn.BeginTransaction
                    DataAdapter.UpdateCommand.Transaction = mytransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    mytransaction.Commit()
                Catch ex As Exception
                    mytransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function FamilyCodeTx(ByVal Myform As FormIndirectFamily, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update Population
                sqlstr = "doc.sp_updateipopulation"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pop").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertipopulation"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pop").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deleteipopulation"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                'Update Family
                sqlstr = "doc.sp_updateifamily"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "key").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "family").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertifamily"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "key").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "family").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deleteifamily"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

                'Update Sub Family
                sqlstr = "doc.sp_updateisubfam"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "key").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbfam").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertisubfam"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "key").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbfam").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "description").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deleteisubfam"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(2))

                'Update FamilyCode
                sqlstr = "doc.sp_updateifamilycode"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familycode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "popid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "subfamid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertifamilycode"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familycode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "popid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "subfamid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deleteifamilycode"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(3))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret


    End Function

End Class

Public Class DbAdapterExeption
    Inherits ApplicationException
    Public Sub New(ByVal errormessage As String)
        MyBase.New(errormessage)
    End Sub
End Class