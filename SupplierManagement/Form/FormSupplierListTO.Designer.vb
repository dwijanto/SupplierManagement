<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormSupplierListTO
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
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.ButtonResetSBU = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.ButtonHProductType = New System.Windows.Forms.Button()
        Me.ButtonHVendorName = New System.Windows.Forms.Button()
        Me.ButtonHShortName = New System.Windows.Forms.Button()
        Me.ButtonHFamily = New System.Windows.Forms.Button()
        Me.ButtonHSBU = New System.Windows.Forms.Button()
        Me.ButtonResetFamily = New System.Windows.Forms.Button()
        Me.ButtonResetShortname = New System.Windows.Forms.Button()
        Me.ButtonResetSuppliername = New System.Windows.Forms.Button()
        Me.ButtonResetProduct = New System.Windows.Forms.Button()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2, Me.ToolStripProgressBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 286)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(598, 22)
        Me.StatusStrip1.TabIndex = 52
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 17)
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(450, 17)
        Me.ToolStripStatusLabel2.Spring = True
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
        Me.ToolStripProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(213, 205)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(169, 42)
        Me.Button1.TabIndex = 51
        Me.Button1.Text = "Generate Report"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "yyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(183, 33)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(119, 20)
        Me.DateTimePicker1.TabIndex = 53
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(140, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 55
        Me.Label1.Text = "Period"
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CustomFormat = "yyyy"
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker2.Location = New System.Drawing.Point(308, 33)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(119, 20)
        Me.DateTimePicker2.TabIndex = 56
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(148, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 13)
        Me.Label2.TabIndex = 57
        Me.Label2.Text = "SBU"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(183, 61)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(244, 20)
        Me.TextBox1.TabIndex = 58
        '
        'ButtonResetSBU
        '
        Me.ButtonResetSBU.Location = New System.Drawing.Point(433, 59)
        Me.ButtonResetSBU.Name = "ButtonResetSBU"
        Me.ButtonResetSBU.Size = New System.Drawing.Size(46, 23)
        Me.ButtonResetSBU.TabIndex = 59
        Me.ButtonResetSBU.Text = "Reset"
        Me.ButtonResetSBU.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(183, 87)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(244, 20)
        Me.TextBox2.TabIndex = 61
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(141, 90)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(36, 13)
        Me.Label3.TabIndex = 60
        Me.Label3.Text = "Family"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(183, 113)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(244, 20)
        Me.TextBox3.TabIndex = 64
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(73, 116)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 13)
        Me.Label4.TabIndex = 63
        Me.Label4.Text = "Supplier Short Name"
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(183, 139)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(244, 20)
        Me.TextBox4.TabIndex = 67
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(101, 142)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 13)
        Me.Label5.TabIndex = 66
        Me.Label5.Text = "Supplier Name"
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(183, 165)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(244, 20)
        Me.TextBox5.TabIndex = 70
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(106, 168)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(71, 13)
        Me.Label6.TabIndex = 69
        Me.Label6.Text = "Product Type"
        '
        'ButtonHProductType
        '
        Me.ButtonHProductType.Location = New System.Drawing.Point(485, 163)
        Me.ButtonHProductType.Name = "ButtonHProductType"
        Me.ButtonHProductType.Size = New System.Drawing.Size(37, 23)
        Me.ButtonHProductType.TabIndex = 76
        Me.ButtonHProductType.Text = "..."
        Me.ButtonHProductType.UseVisualStyleBackColor = True
        '
        'ButtonHVendorName
        '
        Me.ButtonHVendorName.Location = New System.Drawing.Point(485, 137)
        Me.ButtonHVendorName.Name = "ButtonHVendorName"
        Me.ButtonHVendorName.Size = New System.Drawing.Size(37, 23)
        Me.ButtonHVendorName.TabIndex = 75
        Me.ButtonHVendorName.Text = "..."
        Me.ButtonHVendorName.UseVisualStyleBackColor = True
        '
        'ButtonHShortName
        '
        Me.ButtonHShortName.Location = New System.Drawing.Point(485, 111)
        Me.ButtonHShortName.Name = "ButtonHShortName"
        Me.ButtonHShortName.Size = New System.Drawing.Size(37, 23)
        Me.ButtonHShortName.TabIndex = 74
        Me.ButtonHShortName.Text = "..."
        Me.ButtonHShortName.UseVisualStyleBackColor = True
        '
        'ButtonHFamily
        '
        Me.ButtonHFamily.Location = New System.Drawing.Point(485, 85)
        Me.ButtonHFamily.Name = "ButtonHFamily"
        Me.ButtonHFamily.Size = New System.Drawing.Size(37, 23)
        Me.ButtonHFamily.TabIndex = 73
        Me.ButtonHFamily.Text = "..."
        Me.ButtonHFamily.UseVisualStyleBackColor = True
        '
        'ButtonHSBU
        '
        Me.ButtonHSBU.Location = New System.Drawing.Point(485, 59)
        Me.ButtonHSBU.Name = "ButtonHSBU"
        Me.ButtonHSBU.Size = New System.Drawing.Size(37, 23)
        Me.ButtonHSBU.TabIndex = 72
        Me.ButtonHSBU.Text = "..."
        Me.ButtonHSBU.UseVisualStyleBackColor = True
        '
        'ButtonResetFamily
        '
        Me.ButtonResetFamily.Location = New System.Drawing.Point(433, 85)
        Me.ButtonResetFamily.Name = "ButtonResetFamily"
        Me.ButtonResetFamily.Size = New System.Drawing.Size(46, 23)
        Me.ButtonResetFamily.TabIndex = 77
        Me.ButtonResetFamily.Text = "Reset"
        Me.ButtonResetFamily.UseVisualStyleBackColor = True
        '
        'ButtonResetShortname
        '
        Me.ButtonResetShortname.Location = New System.Drawing.Point(433, 111)
        Me.ButtonResetShortname.Name = "ButtonResetShortname"
        Me.ButtonResetShortname.Size = New System.Drawing.Size(46, 23)
        Me.ButtonResetShortname.TabIndex = 78
        Me.ButtonResetShortname.Text = "Reset"
        Me.ButtonResetShortname.UseVisualStyleBackColor = True
        '
        'ButtonResetSuppliername
        '
        Me.ButtonResetSuppliername.Location = New System.Drawing.Point(433, 137)
        Me.ButtonResetSuppliername.Name = "ButtonResetSuppliername"
        Me.ButtonResetSuppliername.Size = New System.Drawing.Size(46, 23)
        Me.ButtonResetSuppliername.TabIndex = 79
        Me.ButtonResetSuppliername.Text = "Reset"
        Me.ButtonResetSuppliername.UseVisualStyleBackColor = True
        '
        'ButtonResetProduct
        '
        Me.ButtonResetProduct.Location = New System.Drawing.Point(433, 163)
        Me.ButtonResetProduct.Name = "ButtonResetProduct"
        Me.ButtonResetProduct.Size = New System.Drawing.Size(46, 23)
        Me.ButtonResetProduct.TabIndex = 80
        Me.ButtonResetProduct.Text = "Reset"
        Me.ButtonResetProduct.UseVisualStyleBackColor = True
        '
        'FormSupplierListTO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(598, 308)
        Me.Controls.Add(Me.ButtonResetProduct)
        Me.Controls.Add(Me.ButtonResetSuppliername)
        Me.Controls.Add(Me.ButtonResetShortname)
        Me.Controls.Add(Me.ButtonResetFamily)
        Me.Controls.Add(Me.ButtonHProductType)
        Me.Controls.Add(Me.ButtonHVendorName)
        Me.Controls.Add(Me.ButtonHShortName)
        Me.Controls.Add(Me.ButtonHFamily)
        Me.Controls.Add(Me.ButtonHSBU)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ButtonResetSBU)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DateTimePicker2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "FormSupplierListTO"
        Me.Text = "FormSupplierListTO"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Public WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ButtonResetSBU As System.Windows.Forms.Button
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ButtonHProductType As System.Windows.Forms.Button
    Friend WithEvents ButtonHVendorName As System.Windows.Forms.Button
    Friend WithEvents ButtonHShortName As System.Windows.Forms.Button
    Friend WithEvents ButtonHFamily As System.Windows.Forms.Button
    Friend WithEvents ButtonHSBU As System.Windows.Forms.Button
    Friend WithEvents ButtonResetFamily As System.Windows.Forms.Button
    Friend WithEvents ButtonResetShortname As System.Windows.Forms.Button
    Friend WithEvents ButtonResetSuppliername As System.Windows.Forms.Button
    Friend WithEvents ButtonResetProduct As System.Windows.Forms.Button
End Class
