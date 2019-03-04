Imports System.Threading
Imports SupplierManagement.PublicClass
Imports System.Text
Imports SupplierManagement.SharedClass
Public Class FormImportVendorSPMPM

    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Public Property ds As DataSet
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Start Thread
        If Not mythread.IsAlive Then
            'Get file
            If openfiledialog1.ShowDialog = DialogResult.OK Then
                mythread = New Thread(AddressOf doWork)
                mythread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub doWork()


        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String
        Try
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
                            'find vendorcode
                            Dim vendorkey(0) As Object
                            vendorkey(0) = myrecord(0)
                            Dim result = ds.Tables(0).Rows.Find(vendorkey)
                            Dim isnew As Boolean = False
                            Dim spmid As Integer
                            Dim pmid As Integer
                            If Not IsNothing(result) Then
                                'Find ssm
                                If myrecord(2) <> "" Then
                                    Dim spmkey(0) As Object
                                    spmkey(0) = myrecord(2)
                                    Dim spm = ds.Tables(3).Rows.Find(spmkey)
                                    If Not IsNothing(spm) Then
                                        spmid = spm.Item("ofsebid")
                                        If result.Item("ssmid") <> spmid Then
                                            result.Item("ssmid") = spm.Item("ofsebid")
                                            isnew = True
                                        End If

                                        If myrecord(4) = "" Then
                                            result.Item("ssmeffectivedate") = DBNull.Value
                                        Else
                                            If IsDBNull(result.Item("ssmeffectivedate")) Then
                                                result.Item("ssmeffectivedate") = CDate(myrecord(4))
                                                isnew = True
                                            ElseIf result.Item("ssmeffectivedate") <> CDate(myrecord(4)) Then
                                                result.Item("ssmeffectivedate") = CDate(myrecord(4))
                                                isnew = True
                                            End If
                                        End If
                                    End If
                                    End If

                                    If myrecord(3) <> "" Then
                                        Dim pmkey(0) As Object
                                        pmkey(0) = myrecord(3)
                                        Dim pm = ds.Tables(4).Rows.Find(pmkey)
                                        If Not IsNothing(pm) Then
                                            pmid = pm.Item("ofsebid")
                                            If result.Item("pmid") <> pm.Item("ofsebid") Then
                                                result.Item("pmid") = pmid
                                                isnew = True
                                            End If
                                            If myrecord(5) = "" Then
                                                result.Item("pmeffectivedate") = DBNull.Value
                                            Else
                                                If IsDBNull(result.Item("pmeffectivedate")) Then
                                                    result.Item("pmeffectivedate") = CDate(myrecord(5))
                                                    isnew = True
                                                ElseIf result.Item("pmeffectivedate") <> CDate(myrecord(5)) Then
                                                    result.Item("pmeffectivedate") = CDate(myrecord(5))
                                                    isnew = True
                                                End If
                                            End If

                                        End If
                                    End If

                                    If isnew Then
                                        'create history
                                        Dim dr = ds.Tables(5).NewRow
                                        dr.Item("vendorcode") = myrecord(0)
                                        dr.Item("spmid") = spmid
                                        dr.Item("pmid") = pmid
                                        If myrecord(4) <> "" Then
                                            dr.Item("spmeffectivedate") = CDate(myrecord(4))
                                        End If
                                        If myrecord(5) <> "" Then
                                            dr.Item("pmeffectivedate") = CDate(myrecord(5))
                                        End If
                                        dr.Item("userid") = HelperClass1.UserInfo.userid
                                        ds.Tables(5).Rows.Add(dr)
                                    End If
                                End If
                            End If
                            count += 1
                    Loop
                End With
            End Using

            'update record
            Dim ds2 = ds.GetChanges
            If Not IsNothing(ds2) Then
                'save records
                Dim mymessage As String = String.Empty
                Dim ra As Integer
                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                If Not DbAdapter1.SPMSSMTx(Me, mye) Then
                    MessageBox.Show(mye.message)
                    Exit Sub
                End If
                ds.Merge(ds2)
                ds.AcceptChanges()
                MessageBox.Show("Saved.")
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
End Class