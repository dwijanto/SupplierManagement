Imports System.Threading
Imports SupplierManagement.PublicClass
Imports System.Text
Imports SupplierManagement.SharedClass

Public Class FormUpdateFamily

    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Public Property groupid As Long = 0
    Dim DS As DataSet
    Dim Mylist As New List(Of String())

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
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
        'Get Inital Table
        ProgressReport(2, "Marquee.")
        Dim sqlstr = "select * from doc.familygroupsbu;"
        DS = New DataSet
        Dim mymessage As String = String.Empty
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(1, mymessage)
            ProgressReport(3, "Continuous.")
            Exit Sub
        End If

        Dim pk1(0) As DataColumn
        pk1(0) = DS.Tables(0).Columns(0)
        DS.Tables(0).PrimaryKey = pk1



        Mylist.Clear()
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
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count >= 1 Then
                        If IsNumeric(myrecord(0)) Then                           
                            Mylist.Add(myrecord)
                        End If
                    End If
                    count += 1
                Loop
            End With
        End Using
        'update record
        'Find Existing Data, if not avail create, if avail update
        Try
            For i = 0 To Mylist.Count - 1
                Dim mykey(0) As Object
                Dim mydata = Mylist(i)
                mykey(0) = mydata(0)
                Dim dr As DataRow = DS.Tables(0).Rows.Find(mykey)
                'If mydata(0) = 36 Then
                '    Debug.Print("debug")
                'End If
                If IsNothing(dr) Then
                    dr = DS.Tables(0).NewRow
                    dr.Item(0) = mydata(0)
                    dr.Item(1) = mydata(1)
                    dr.Item(2) = mydata(2)
                    If mydata(3) <> "" Then dr.Item(3) = mydata(3)
                    DS.Tables(0).Rows.Add(dr)
                Else
                    dr.Item(1) = mydata(1)
                    dr.Item(2) = mydata(2)
                    If mydata(3) <> "" Then dr.Item(3) = mydata(3)
                End If
            Next

            Dim ds2 As DataSet = DS.GetChanges
            If Not IsNothing(ds2) Then

                Dim ra As Integer
                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)

                If Not DbAdapter1.FamilyGroupSBUTx(Me, mye) Then
                    MessageBox.Show(mye.message)
                    DS.Merge(ds2)
                    Exit Sub
                End If
                ProgressReport(1, "Done.")
            End If
            ProgressReport(3, "Continuous.")
        Catch ex As Exception
            ProgressReport(1, ex.Message)
            ProgressReport(3, "Continuous.")
        End Try
       

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

    Private Sub UpdateData()
        
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Start Thread
        If MessageBox.Show("Delete all data?", "Clear Data", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
            If Not mythread.IsAlive Then
                'Get file

                mythread = New Thread(AddressOf doDelete)
                mythread.Start()

            Else
                MessageBox.Show("Process still running. Please Wait!")
            End If
        End If

    End Sub

    Private Sub doDelete()
        ProgressReport(2, "Marquee.")
        ProgressReport(1, "Deleting...")
        Dim sqlstr = "delete from doc.familygroupsbu;"
        DbAdapter1.ExecuteScalar(sqlstr)
        ProgressReport(1, "Done.")
        ProgressReport(3, "Continuous.")
    End Sub

End Class