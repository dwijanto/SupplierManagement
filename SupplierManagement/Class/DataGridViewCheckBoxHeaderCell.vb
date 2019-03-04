Public Delegate Sub CheckBoxClickedHandler(ByVal state As Boolean)

Public Class DataGridViewCheckBoxHeaderCellEventArgs
    Inherits EventArgs
    Dim _bChecked As Boolean
    Public Sub New(ByVal bChecked As Boolean)
        _bChecked = bChecked
    End Sub
    Public ReadOnly Property Checked As Boolean
        Get
            Return _bChecked
        End Get
    End Property
End Class

Public Class DataGridViewCheckBoxHeaderCell
    Inherits DataGridViewColumnHeaderCell

    Dim checkBoxLocation As Point
    Dim checkboxsize As Size
    Dim _checked As Boolean = False
    Dim _cellLocation As New Point()
    Dim _cbState As System.Windows.Forms.VisualStyles.CheckBoxState = System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal

    Public Event onCheckbBoxClicked As CheckBoxClickedHandler

    Public Sub New()

    End Sub

    Protected Overrides Sub Paint(ByVal graphics As System.Drawing.Graphics, ByVal clipBounds As System.Drawing.Rectangle, ByVal cellBounds As System.Drawing.Rectangle, ByVal rowIndex As Integer, ByVal dataGridViewElementState As System.Windows.Forms.DataGridViewElementStates, ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, ByVal cellStyle As System.Windows.Forms.DataGridViewCellStyle, ByVal advancedBorderStyle As System.Windows.Forms.DataGridViewAdvancedBorderStyle, ByVal paintParts As System.Windows.Forms.DataGridViewPaintParts)
        MyBase.Paint(graphics, clipBounds, cellBounds, rowIndex, dataGridViewElementState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts)
        Dim p = New Point()
        Dim s = CheckBoxRenderer.GetGlyphSize(graphics, System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal)
        p.X = cellBounds.Location.X + (cellBounds.Width / 2) - (s.Width / 2)
        p.Y = cellBounds.Location.Y + (cellBounds.Height / 2) - (s.Height / 2)
        _cellLocation = cellBounds.Location
        checkBoxLocation = p
        checkboxsize = s
        If (_checked) Then
            _cbState = System.Windows.Forms.VisualStyles.CheckBoxState.CheckedNormal
        Else
            _cbState = System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal
        End If
        CheckBoxRenderer.DrawCheckBox(graphics, checkBoxLocation, _cbState)
    End Sub

    Protected Overrides Sub OnMouseClick(ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)
        Dim p = New Point(e.X + _cellLocation.X, e.Y + _cellLocation.Y)
        If p.X >= checkBoxLocation.X AndAlso p.X <= checkBoxLocation.X + checkboxsize.Width AndAlso p.Y >= checkBoxLocation.Y AndAlso p.Y <= checkBoxLocation.Y + checkboxsize.Height Then
            _checked = Not _checked
            RaiseEvent onCheckbBoxClicked(_checked)
            Me.DataGridView.InvalidateCell(Me)
        End If
        MyBase.OnMouseClick(e)
    End Sub
End Class
