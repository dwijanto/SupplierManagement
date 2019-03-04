Imports System.Threading
Imports SupplierManagement.PublicClass
Imports System.Text
Imports SupplierManagement.SharedClass
Public Class FormImportNQSU
    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Public Property ds As DataSet
    Dim period1 As Date
    Dim period2 As Date

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim myMessage As String = String.Empty
        If String.Format("{0:yyyyMM}", DateTimePicker1.Value) = String.Format("{0:yyyyMM}", DateTimePicker2.Value) Then
            myMessage = String.Format("Delete period {0:yyyyMM}?", DateTimePicker1.Value)
        Else
            myMessage = String.Format("Delete period from {0:yyyyMM} to {1:yyyyMM}?", DateTimePicker1.Value, DateTimePicker2.Value)
        End If
        If MessageBox.Show(myMessage, "Delete Data", Windows.Forms.MessageBoxButtons.OKCancel) = DialogResult.OK Then
            If Not mythread.IsAlive Then
                'Get file

                mythread = New Thread(AddressOf doDelete)
                mythread.Start()

            Else
                MessageBox.Show("Process still running. Please Wait!")
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not mythread.IsAlive Then
            'Get file
            If openfiledialog1.ShowDialog = DialogResult.OK Then
                period1 = DateTimePicker1.Value
                period2 = DateTimePicker2.Value
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
            Using objTFParser = New FileIO.TextFieldParser(openfiledialog1.FileName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    ProgressReport(1, "Read Data")
                    ProgressReport(2, "Read Data")
                    Dim myyear As Integer
                    Do Until .EndOfData
                        myrecord = .ReadFields

                        If count >= 1 Then
                            If IsNumeric(myrecord(0)) Then

                                If myrecord(0) >= String.Format("{0:yyyyMM}", period1) And myrecord(0) <= String.Format("{0:yyyyMM}", period2) Then

                                    myyear = myrecord(0).Substring(0, 4)
                                    'period,year,fpcp,vendorcode,inspected,rejected,criticaldefect,majordefect,samplesize,inspectedytd,rejectedytd,criticaldefectytd,majordefectytd,samplesizeytd
                                    myInsert.Append(validint(myrecord(0)) & vbTab &
                                                    myyear & vbTab &
                                                    validstr(myrecord(1)) & vbTab &
                                                    validlong(myrecord(2)) & vbTab &
                                                    validint(myrecord(4)) & vbTab &
                                                    validint(myrecord(5)) & vbTab &
                                                    validint(myrecord(6)) & vbTab &
                                                    validint(myrecord(7)) & vbTab &
                                                    validint(myrecord(8)) & vbTab &
                                                    validint(myrecord(9)) & vbTab &
                                                    validint(myrecord(10)) & vbTab &
                                                    validint(myrecord(11)) & vbTab &
                                                    validint(myrecord(12)) & vbTab &
                                                    validint(myrecord(13)) & vbCrLf)
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
                mystr.Append(String.Format("delete from doc.nqsu where period >= {0:yyyyMM} and period <= {1:yyyyMM};", period1, period2))
                Dim sqlstr As String = "copy doc.nqsu(period,year,fpcp,vendorcode,inspected,rejected,criticaldefect,majordefect,samplesize,inspectedytd,rejectedytd,criticaldefectytd,majordefectytd,samplesizeytd) from stdin with null as 'Null';"
                Dim ra As Long = 0
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False

                Try
                    If RadioButton1.Checked Then
                        ProgressReport(1, "Replace Record Please wait!")
                        ra = DbAdapter1.ExNonQuery(mystr.ToString)
                    End If
                    ProgressReport(1, "Add Record Please wait!")
                    errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                    If myret Then
                        DbAdapter1.setaudit("NQSU", HelperClass1.UserInfo.userid)
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

    Private Sub doDelete()
        Dim sqlstr = String.Format("delete from doc.nqsu where period >= {0:yyyyMM} and period <= {1:yyyyMM}", DateTimePicker1.Value, DateTimePicker2.Value)
        Dim mymessage As String = String.Empty
        ProgressReport(1, "Deleting... Please wait.")
        If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
            MessageBox.Show(mymessage)
        Else
            ProgressReport(1, "Done")
        End If

    End Sub

End Class