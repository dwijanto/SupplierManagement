Imports System.Threading
Imports SupplierManagement.PublicClass
Imports System.Text
Imports SupplierManagement.SharedClass
Imports System.ComponentModel

Public Class FormImportBudgetForecast
    Implements INotifyPropertyChanged


    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Public Property ds As DataSet
    Private MyYear As Integer
    Private myPeriod As Date

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
    Private Enum DataTypeEnum
        BudgetAndForecast = 1
        SupplyChainTurnover = 2

    End Enum
    Public Property DataType As Integer
        Get
            If RadioButton1.Checked Then
                Return 1
            Else
                Return 2
            End If
        End Get
        Set(ByVal value As Integer)
            If value = 1 Then
                RadioButton1.Checked = True
            Else
                RadioButton2.Checked = True
            End If
        End Set
    End Property


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not mythread.IsAlive Then
            'Get file
            MyYear = DateTimePicker1.Value.Year
            Select Case DataType
                Case DataTypeEnum.BudgetAndForecast
                    myPeriod = CDate(String.Format("{0}-1-1", DateTimePicker1.Value.Year))
                Case DataTypeEnum.SupplyChainTurnover
                    myPeriod = DateTimePicker1.Value
            End Select

            If openfiledialog1.ShowDialog = DialogResult.OK Then
                mythread = New Thread(AddressOf doWork)
                mythread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub doWork()
        Try
            Dim mystr As New StringBuilder
            Dim myInsert As New System.Text.StringBuilder
            Dim myrecord() As String

            Dim myDataTypeAdapter As New DataTypeAdapter
            Dim DataTypeBS As BindingSource = myDataTypeAdapter.getCombobBoxDataSource




            Using objTFParser = New FileIO.TextFieldParser(openfiledialog1.FileName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    ProgressReport(1, "Read Data")
                    ProgressReport(2, "Read Data")

                    Do Until .EndOfData
                        myrecord = .ReadFields

                        If count >= 1 Then
                            If IsNumeric(myrecord(8)) And IsNumeric(myrecord(2)) Then
                                Dim pos = DataTypeBS.Find("datatypename", myrecord(1))
                                If pos < 0 Then
                                    MessageBox.Show(String.Format("""{0}"" is not listed in DataType Master. Please check.", myrecord(1)))
                                    Exit Sub
                                Else

                                    DataTypeBS.Position = pos
                                    Dim drv As DataRowView = DataTypeBS.Current

                                    myInsert.Append(myPeriod & vbTab &
                                                    drv.Item("id") & vbTab &
                                                    validstr(myrecord(0)) & vbTab &
                                                    validlong(myrecord(2)) & vbTab &
                                                    validlong(myrecord(3)) & vbTab &
                                                    validstr(myrecord(4)) & vbTab &
                                                    validstr(myrecord(5)) & vbTab &
                                                    validint(myrecord(6)) & vbTab &
                                                    validstr(myrecord(7)) & vbTab &
                                                    validnumeric(myrecord(8)) & vbTab &
                                                    validnumeric(myrecord(9)) & vbCrLf)
                                End If

                            End If
                        End If
                        count += 1
                    Loop
                End With
            End Using
            'update record
            If myInsert.Length > 0 Then
                ProgressReport(1, "Start Add New Records")
                mystr.Append(String.Format("delete from doc.budgetforecast where period = '{0:yyyy-MM-dd}';", myPeriod))
                Dim sqlstr As String = "copy doc.budgetforecast(period,datatypeid,fpcp,vendorcode,cmmf,brand,sbu,comfam,range,qty,amount) from stdin with null as 'Null';"
                Dim ra As Long = 0
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False

                Try

                    ProgressReport(1, "Replace Record Please wait!")
                    ra = DbAdapter1.ExNonQuery(mystr.ToString)
                    If DbAdapter1.ExecuteScalar("select doc.getseqid('doc.budgetforecast')", ra) Then
                        If ra = 1 Then

                            DbAdapter1.ExNonQuery(String.Format("select setval('doc.budgetforecast_id_seq',{0},false);", ra))
                        Else
                            DbAdapter1.ExNonQuery(String.Format("select setval('doc.budgetforecast_id_seq',{0},true);", ra))
                        End If

                    End If


                    ProgressReport(1, "Add Record Please wait!")
                    errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                    If myret Then
                        'DbAdapter1.setaudit("CitiProgram", HelperClass1.UserInfo.userid)
                        ProgressReport(1, "Add Records Done.")
                    Else
                        ProgressReport(1, errmessage)
                    End If
                Catch ex As Exception
                    ProgressReport(1, ex.Message)

                End Try

            End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try

        ProgressReport(3, "Set Continuous Again")
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

                Case 3
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
            End Select

        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged, RadioButton1.CheckedChanged
        If DataType = DataTypeEnum.BudgetAndForecast Then
            DateTimePicker1.CustomFormat = "yyyy"
        ElseIf DataType = DataTypeEnum.SupplyChainTurnover Then
            DateTimePicker1.CustomFormat = "dd-MMM-yyyy"
        End If
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'MyYear = DateTimePicker1.Value.Year
        'Select Case DataType
        '    Case DataTypeEnum.BudgetAndForecast
        '        myPeriod = CDate(String.Format("{0}-1-1", DateTimePicker1.Value.Year))
        '    Case DataTypeEnum.SupplyChainTurnover
        '        myPeriod = DateTimePicker1.Value
        'End Select
        If MessageBox.Show(String.Format("Delete this period '{0:dd-MMM-yyyy}'?", getPeriod), "Delete Data", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
            If Not mythread.IsAlive Then
                'Get file



                mythread = New Thread(AddressOf doWork2)
                mythread.Start()

            Else
                MessageBox.Show("Process still running. Please Wait!")
            End If

        End If

    End Sub

    Private Sub doWork2()
        ProgressReport(2, "Set Marque")
        ProgressReport(1, "Deleting...")
        Dim mystr As New StringBuilder
        MyYear = DateTimePicker1.Value.Year
        'Select Case DataType
        '    Case DataTypeEnum.BudgetAndForecast
        '        myPeriod = CDate(String.Format("{0}-1-1", DateTimePicker1.Value.Year))
        '    Case DataTypeEnum.SupplyChainTurnover
        '        myPeriod = DateTimePicker1.Value
        'End Select
        mystr.Append(String.Format("delete from doc.budgetforecast where period = '{0:yyyy-MM-dd}';", getPeriod))
        Dim ra As Integer = DbAdapter1.ExNonQuery(mystr.ToString)
        ProgressReport(1, "Done.")
        ProgressReport(3, "Set Continuous Again")
    End Sub

    Private Function getPeriod() As Date
        MyYear = DateTimePicker1.Value.Year
        Select Case DataType
            Case DataTypeEnum.BudgetAndForecast
                myPeriod = CDate(String.Format("{0}-1-1", DateTimePicker1.Value.Year))
            Case DataTypeEnum.SupplyChainTurnover
                myPeriod = DateTimePicker1.Value
        End Select
        Return myPeriod
    End Function

End Class