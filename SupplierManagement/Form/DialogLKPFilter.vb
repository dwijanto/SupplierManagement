Imports System.Windows.Forms
Imports System.Text

Public Class DialogLKPFilter

    Private Shared myinstance As DialogLKPFilter
    Public Shared Function GetInstance() As DialogLKPFilter
        If IsNothing(myinstance) Then
            myinstance = New DialogLKPFilter
        End If
        Return myinstance
    End Function

    Public Sub resetFilter()
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
    End Sub

    Public Function GetLKPFilter() As String
        Dim all As Integer
        Dim LKPFilter As New StringBuilder
        If CheckBox1.Checked Then
            LKPFilter.Append("LKP")
            all = all + 1
        End If
        If CheckBox2.Checked Then
            If LKPFilter.Length > 0 Then
                LKPFilter.Append(",")
            End If
            LKPFilter.Append("LKP-1")
            all = all + 1
        End If
        If CheckBox3.Checked Then
            If LKPFilter.Length > 0 Then
                LKPFilter.Append(",")
            End If
            LKPFilter.Append("LKP-2")
            all = all + 1
        End If
        If CheckBox4.Checked Then
            If LKPFilter.Length > 0 Then
                LKPFilter.Append(",")
            End If
            LKPFilter.Append("LKP-3")
            all = all + 1
        End If
        If CheckBox5.Checked Then
            If LKPFilter.Length > 0 Then
                LKPFilter.Append(",")
            End If
            LKPFilter.Append("LKP-5")
            all = all + 1
        End If
        'Return LKPFilter.ToString
        If all = 0 Then
            Return "ALL"
        Else
            Return LKPFilter.ToString
        End If
    End Function

    Private Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
