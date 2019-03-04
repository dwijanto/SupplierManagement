<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormDialogContact
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button_Cancel = New System.Windows.Forms.Button()
        Me.Button_Ok = New System.Windows.Forms.Button()
        Me.UcContact1 = New SupplierManagement.UCContact()
        Me.SuspendLayout()
        '
        'Button_Cancel
        '
        Me.Button_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button_Cancel.Location = New System.Drawing.Point(397, 349)
        Me.Button_Cancel.Name = "Button_Cancel"
        Me.Button_Cancel.Size = New System.Drawing.Size(77, 38)
        Me.Button_Cancel.TabIndex = 4
        Me.Button_Cancel.Text = "Cancel"
        Me.Button_Cancel.UseVisualStyleBackColor = True
        '
        'Button_Ok
        '
        Me.Button_Ok.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button_Ok.Location = New System.Drawing.Point(314, 349)
        Me.Button_Ok.Name = "Button_Ok"
        Me.Button_Ok.Size = New System.Drawing.Size(77, 38)
        Me.Button_Ok.TabIndex = 3
        Me.Button_Ok.Text = "OK"
        Me.Button_Ok.UseVisualStyleBackColor = True
        '
        'UcContact1
        '
        Me.UcContact1.bs = Nothing
        Me.UcContact1.Location = New System.Drawing.Point(12, 12)
        Me.UcContact1.Name = "UcContact1"
        Me.UcContact1.Size = New System.Drawing.Size(462, 331)
        Me.UcContact1.TabIndex = 0
        '
        'FormDialogContact
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button_Cancel
        Me.ClientSize = New System.Drawing.Size(486, 408)
        Me.Controls.Add(Me.Button_Cancel)
        Me.Controls.Add(Me.Button_Ok)
        Me.Controls.Add(Me.UcContact1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormDialogContact"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FormDialogContact"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UcContact1 As SupplierManagement.UCContact
    Friend WithEvents Button_Cancel As System.Windows.Forms.Button
    Friend WithEvents Button_Ok As System.Windows.Forms.Button
End Class
