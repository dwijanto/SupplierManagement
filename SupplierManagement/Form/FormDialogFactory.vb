Public Class FormDialogFactory
    Public errormessage As String
    Private bs As BindingSource
    Public Property VBS As BindingSource
    Public Property VFBS As BindingSource
    Enum TxAction
        New_Record = 0
        Update_Record = 1
    End Enum
    Public Property MyAction As TxAction
    Public Sub New(ByVal BS As BindingSource, ByVal ProvinceBS As BindingSource, ByVal CountryBS As BindingSource)
        ' This call is required by the designer.

        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.bs = BS
        UcFactory1.ProvinceBS = ProvinceBS
        UcFactory1.CountryBS = CountryBS
        UcFactory1.bs = BS
        UcFactory1.DataBinding()
    End Sub


    Private Sub Button_Ok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Ok.Click
        Me.Validate()
        bs.EndEdit()
        If Not UcFactory1.ValidateInput(errormessage) Then
            Me.DialogResult = DialogResult.None
        Else
            If MyAction = TxAction.New_Record Then
                If VBS.Count > 1 Then
                    'MessageBox.Show("Add Vendor")
                    Dim myform As New FormDialogAddVendor
                    myform.BS = VBS
                    Dim mydrv As DataRowView = bs.Current
                    'mydrv.BeginEdit()
                    'mydrv.EndEdit()
                    VFBS.Filter = String.Format("factoryid = {0}", mydrv.Row.Item("id"))
                    myform.VFBS = VFBS
                    myform.FDBS = bs
                    myform.DataBinding()
                    myform.ShowDialog()
                Else


                    'For Each drv As DataRowView In VFBS.List
                    '    drv.Delete()
                    'Next
                    ''Add vendorfacotry for True Status
                    'For Each drv As DataRowView In VBS.List
                    '    'BS.RemoveAt(drv.Index)


                    '    If drv.Row.Item("status") Then
                    '        Dim myrow As DataRowView = VFBS.AddNew
                    '        myrow.Item("vendorcode") = drv.Row.Item("vendorcode")
                    '        Dim mydrv As DataRowView = bs.Current
                    '        myrow.Item("factoryid") = mydrv.Row.Item("id")
                    '        VFBS.EndEdit()
                    '    End If
                    'Next
                    For Each drv As DataRowView In VBS.List
                        'BS.RemoveAt(drv.Index)


                        'If drv.Row.Item("status") Then
                        Dim myrow As DataRowView = VFBS.AddNew
                        myrow.Item("vendorcode") = drv.Row.Item("vendorcode")
                        Dim mydrv As DataRowView = bs.Current
                        myrow.Item("factoryid") = mydrv.Row.Item("id")
                        VFBS.EndEdit()
                        'End If
                    Next
                End If
            End If
            


        End If

    End Sub

    Private Sub Button_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Cancel.Click
        Dim drv As DataRowView = bs.Current
        drv.CancelEdit()
    End Sub
End Class