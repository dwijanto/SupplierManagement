Public Class FormDialogContact
    Public errormessage As String
    Private bs As BindingSource
    Public Property VBS As BindingSource
    Public Property VCBS As BindingSource
    Enum TxAction
        New_Record = 0
        Update_Record = 1
    End Enum
    Public Property MyAction As TxAction

    Public Sub New(ByVal BS As BindingSource)
        ' This call is required by the designer.

        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.bs = BS
        MyAction = TxAction.New_Record
        UcContact1.bs = BS
        UcContact1.DataBinding()
    End Sub


    Private Sub Button_Ok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Ok.Click
        Me.Validate()
        bs.EndEdit()
        If Not UcContact1.ValidateInput(errormessage) Then
            Me.DialogResult = DialogResult.None
        Else
            If MyAction = TxAction.New_Record Then                
                If VBS.Count > 1 Then  'Shortname
                    'MessageBox.Show("Add Vendor")
                    Dim myform As New FormDialogAddVendor
                    myform.BS = VBS
                    Dim mydrv As DataRowView = bs.Current
                    'mydrv.BeginEdit()
                    'mydrv.EndEdit()
                    VCBS.Filter = String.Format("contactid = {0}", mydrv.Row.Item("id"))
                    myform.VFBS = VCBS
                    myform.FDBS = bs
                    myform.DataBinding()
                    myform.ShowDialog()
                Else
                    'By Vendorcode
                    'For Each drv As DataRowView In VCBS.List
                    '    drv.Delete()
                    'Next
                    'Add vendorfacotry for True Status
                    For Each drv As DataRowView In VBS.List
                        'BS.RemoveAt(drv.Index)
                        If drv.Row.Item("status") Then
                            Dim myrow As DataRowView = VCBS.AddNew
                            myrow.Item("vendorcode") = drv.Row.Item("vendorcode")
                            Dim mydrv As DataRowView = bs.Current
                            myrow.Item("contactid") = mydrv.Row.Item("id")
                            myrow.Item("vendorname") = drv.Row.Item("vendorname")
                            VCBS.EndEdit()
                        End If
                    Next
                End If
            Else
            End If
        End If

    End Sub

    Private Sub Button_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Cancel.Click
        Dim drv As DataRowView = bs.Current
        drv.CancelEdit()
    End Sub
End Class