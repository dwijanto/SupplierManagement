Imports System.Threading
Public Class FormAutoSendAssetPurchase
    Dim myThread As New System.Threading.Thread(AddressOf doWork)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Private Sub FormAutoSendAssetPurchase_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Minimized
        LoadMe()
    End Sub


    Private Sub FormAutoSendAssetPurchase_Resize(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.WindowState = FormWindowState.Minimized Then
            Me.ShowInTaskbar = False
            Me.Hide()
            'NotifyIcon1.Visible = True
        End If
    End Sub


    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick

    End Sub

    Private Sub LoadMe()

        If Not myThread.IsAlive Then
            Try
                myThread = New System.Threading.Thread(AddressOf doWork)
                myThread.TrySetApartmentState(ApartmentState.MTA)
                myThread.Start()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Sub doWork()
        Logger.log("--------Start----------")

        'Email For Approval
        Logger.log("Send email to Purchasing Validator1")
        Dim PurchasingValidator1 = New RoleAssetPurchaseTask(RoleAssetPurchaseTask.RoleAssetPurchaseTaskEnum.PurchasingValidator1)
        If Not PurchasingValidator1.Execute Then
            Logger.log(PurchasingValidator1.errormessage)
        End If

        Logger.log("Send email to Purchasing Validator2")
        Dim PurchasingValidator2 = New RoleAssetPurchaseTask(RoleAssetPurchaseTask.RoleAssetPurchaseTaskEnum.PurchasingValidator2)
        If Not PurchasingValidator2.Execute Then
            Logger.log(PurchasingValidator2.errormessage)
        End If

        Logger.log("Send email rejected to Creator")
        Dim Creator = New RoleAssetPurchaseTask(RoleAssetPurchaseTask.RoleAssetPurchaseTaskEnum.Creator)
        If Not Creator.Execute Then
            Logger.log(Creator.errormessage)
        End If

        Logger.log("Send email to CompletedInvestmentApproval")
        Dim InvestmentApproval = New RoleAssetPurchaseTask(RoleAssetPurchaseTask.RoleAssetPurchaseTaskEnum.InvestmentApproval)
        If Not InvestmentApproval.Execute Then
            Logger.log(PurchasingValidator2.errormessage)
        End If

        Logger.log("Send email Amortization to ControllingDept")

        Dim ControllingAmortization = New RoleAssetPurchaseTask(RoleAssetPurchaseTask.RoleAssetPurchaseTaskEnum.ControllingDeptAmortization)
        If Not ControllingAmortization.Execute Then
            Logger.log(ControllingAmortization.errormessage)
        End If

        ''Logger.log("Send email Investment to ControllingDept")
        ''Dim ControllingInvestment = New RoleAssetPurchaseTask(RoleDocumentTask.RoleDocumentTaskEnum.Validator)
        ''If Not ControllingInvestment.Execute Then
        ''    Logger.log(ControllingInvestment.errormessage)
        ''End If

        Logger.log("Send email to AccountingDept")
        Dim AccountingDept = New RoleAssetPurchaseTask(RoleAssetPurchaseTask.RoleAssetPurchaseTaskEnum.AccountingDept)
        If Not AccountingDept.Execute Then
            Logger.log(AccountingDept.errormessage)
        End If

        Logger.log("Send email completed to IT")
        Dim ITDept = New RoleAssetPurchaseTask(RoleAssetPurchaseTask.RoleAssetPurchaseTaskEnum.ITDept)
        If Not ITDept.Execute Then
            Logger.log(ITDept.errormessage)
        End If

        Logger.log("Send email completed to Applicant")
        Dim Applicant = New RoleAssetPurchaseTask(RoleAssetPurchaseTask.RoleAssetPurchaseTaskEnum.Applicant)
        If Not Applicant.Execute Then
            Logger.log(ITDept.errormessage)
        End If

        Logger.log("--------End------------")
        ProgressReport(1, "Close Apps")
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.Close()
                Case 2

                Case 3

                Case 4

                Case 5

                Case 6

                Case 7

            End Select
        End If
    End Sub
End Class