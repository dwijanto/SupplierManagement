Public Class FormDialogAddVendor

    Public Property BS As BindingSource 'Vendor list
    Public Property VFBS As BindingSource 'Vendor Factory
    Public Property FDBS As BindingSource 'Factory

    Public Sub DataBinding()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = BS

        'Remove Existing Record
        If VFBS.Count > 0 Then
            'Check only avil
            For Each drv As DataRowView In BS.List
                'BS.RemoveAt(drv.Index)
                drv.Row.Item("status") = True
                If (VFBS.Find("vendorcode", drv.Row.Item("vendorcode"))) < 0 Then
                    drv.Row.Item("status") = False
                End If

            Next
        Else
            For Each drv As DataRowView In BS.List
                drv.Row.Item("status") = True
            Next

        End If

       

    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Delete vendorfactory 
        Me.Validate()

        For Each drv As DataRowView In VFBS.List
            drv.Delete()
        Next

        'Add vendorfacotry for True Status
        For Each drv As DataRowView In BS.List
            'BS.RemoveAt(drv.Index)
            If drv.Row.Item("status") Then
                Dim myrow As DataRowView = VFBS.AddNew
                myrow.Item("vendorcode") = drv.Row.Item("vendorcode")
                myrow.Item("vendorname") = drv.Row.Item("vendorname")
                Dim mydrv As DataRowView = FDBS.Current
                myrow.Item(1) = mydrv.Row.Item("id")
                VFBS.EndEdit()
            End If            
        Next
        Me.Close()
    End Sub

End Class