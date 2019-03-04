Imports System.Threading
Public Class FormImportZF0041
    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Public Property ds As DataSet

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

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
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

    Sub doWork()
        Dim myImport = New ImportZF0041(Me, openfiledialog1.FileName)
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(1, "Processing. Please wait..")
        ProgressReport(2, "Marque")
        If myImport.doImport Then
            sw.Stop()
            ProgressReport(1, "Done. Elapsed Time: " & Format(sw.Elapsed.Minutes, "00") & ":" & Format(sw.Elapsed.Seconds, "00") & "." & sw.Elapsed.Milliseconds.ToString)            
        Else
            sw.Stop()
            ProgressReport(1, String.Format("Error:{0}", myImport.ErrorMessage))
        End If
        ProgressReport(3, "Marque")

    End Sub

End Class