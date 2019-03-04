Imports System.Text

Public Class UCTextWithHelper
    Public Property HelperBindingSource As BindingSource
    Private _textSB As StringBuilder
    Private textArray As StringBuilder
    Public Event ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs)

    Public Property WhereSB As StringBuilder
    Public Property SearchField As String

    Public Property LeftJoinMMSB As StringBuilder
    Public Property LeftJoinStr As String

    Public Property LeftJoinProjectSB As StringBuilder
    Public Property LeftJoinProjectSTR As String

    Public Property WithSocialAuditSB As StringBuilder
    Public Property WithSocialAuditSTR As String
    Public Property WithSB As StringBuilder

    Public Property LeftJoinSocialAuditSB As StringBuilder
    Public Property LeftJoinSocialAuditSTR As String

    Public Property LeftJoinVFPSB As StringBuilder
    Public Property LeftJoinVFPSTR As String

    Public Property LeftJoinSupplierPanelSB As StringBuilder
    Public Property LeftJoinSupplierPanelSTR As String

    Public Property LeftJoinFactorySB As StringBuilder
    Public Property LeftJoinFactorySTR As String

    Public Property CheckLatest As Boolean = False
    Public Property LatestValue As Boolean
    Public Property WithSIFSB As StringBuilder
    Public Property WithIDSB As StringBuilder
    Public Property WithSIFIDSSTRLatest As String
    Public Property WithSIFIDSSTRNOTLatest As String
    Public Property WithTechnologySB As StringBuilder
    Public Property WithTechnologySTR As String

    Public Property LeftJoinSIFSB As StringBuilder
    Public Property LeftJoinIDSB As StringBuilder
    Public Property LeftJoinSIFSTR As String
    Public Property LeftJoinIDSTR As String

    Public Property LeftJoinTechnologySB As StringBuilder
    Public Property LeftJoinTechnologySTR As String

    Public Property AdditionalDisplayFieldSB As StringBuilder
    Public Property AdditionalDisplayFieldSTR As String

    Public Property LabelText As String
        Get
            Return Label1.Text
        End Get
        Set(ByVal value As String)
            Label1.Text = value
        End Set
    End Property

    Public Property LabelLocation As Point
        Get
            Return Label1.Location
        End Get
        Set(ByVal value As Point)
            Label1.Location = value
        End Set
    End Property

    Public Property LabelSize As Size
        Get
            Return Label1.Size
        End Get
        Set(ByVal value As Size)
            Label1.Size = value
        End Set
    End Property

    Public Property TextBoxLocation As Point
        Get
            Return TextBox1.Location
        End Get
        Set(ByVal value As Point)
            TextBox1.Location = value
        End Set
    End Property

    Public Property TextBoxSize As Size
        Get
            Return TextBox1.Size
        End Get
        Set(ByVal value As Size)
            TextBox1.Size = value
        End Set
    End Property

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Label1.Location = LabelLocation
        TextBox1.Location = TextBoxLocation
        _textSB = New StringBuilder
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        RaiseEvent ButtonClick(Me, e)
    End Sub

    Public Sub LoadHelper()
        If Not IsNothing(HelperBindingSource) Then
            Dim myform As New FormHelperMulti(HelperBindingSource)
            If myform.ShowDialog = DialogResult.OK Then
                If _textSB.Length = 0 Then
                Else
                    _textSB.Append(",")
                End If
                _textSB.Append(myform.SelectedResult)
                TextBox1.Text = _textSB.ToString
            End If
        Else
            MessageBox.Show("Helper not available.")
        End If
    End Sub

    Public Sub ClearText()
        _textSB.Clear()
        TextBox1.Text = String.Empty
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        _textSB.Clear()
        _textSB.Append(validateString(TextBox1.Text))
    End Sub

    Public Function getTextValue() As String
        Return _textSB.ToString
    End Function

    Public Sub AppendWithSB(ByVal sb As StringBuilder)
        If Not IsNothing(WithSB) Then
            If WithSB.Length = 0 Then
                WithSB.Append("with ")
            Else
                WithSB.Append(", ")
            End If
        End If
        WithSB.Append(sb.ToString)
    End Sub
    Public Sub DoProcess()
        If TextBox1.Text <> "" Then


            textArray = New StringBuilder
            Dim arr() As String = TextBox1.Text.Split(",")

            For i = 0 To arr.Length - 1
                If textArray.Length > 0 Then
                    textArray.Append(",")
                End If
                'textArray.Append(String.Format("'{0}'", arr(i).Replace("''", "''''").ToLower))
                textArray.Append(String.Format("'{0}'", arr(i).Replace("'", "''").ToLower))
            Next

            If WhereSB.Length = 0 Then
                WhereSB.Append(String.Format("Where {0} in ({1}) ", SearchField, textArray.ToString))
            Else
                If Not IsNothing(WithIDSB) Or Not IsNothing(WithSIFSB) Then
                    WhereSB.Append(String.Format(" or {0} in ({1}) ", SearchField, textArray.ToString))
                Else
                    WhereSB.Append(String.Format(" and {0} in ({1}) ", SearchField, textArray.ToString))
                End If

            End If


            If Not IsNothing(LeftJoinMMSB) Then
                LeftJoinMMSB.Append(String.Format("{0}", LeftJoinStr))
            End If
            If Not IsNothing(LeftJoinProjectSB) Then
                LeftJoinProjectSB.Append(String.Format("{0}", LeftJoinProjectSTR))
            End If


            If Not IsNothing(WithSocialAuditSB) Then
                WithSocialAuditSB.Append(String.Format("{0}", WithSocialAuditSTR))
                LeftJoinSocialAuditSB.Append(String.Format("{0}", LeftJoinSocialAuditSTR))
                AppendWithSB(WithSocialAuditSB)
            End If

            If Not IsNothing(LeftJoinVFPSB) Then
                LeftJoinVFPSB.Append(String.Format("{0}", LeftJoinVFPSTR))
            End If

            If Not IsNothing(LeftJoinSupplierPanelSB) Then
                If LeftJoinSupplierPanelSB.Length = 0 Then
                    LeftJoinSupplierPanelSB.Append(String.Format("{0}", LeftJoinSupplierPanelSTR))
                End If
            End If


            If Not IsNothing(LeftJoinTechnologySB) Then
                If LeftJoinTechnologySB.Length = 0 Then
                    LeftJoinTechnologySB.Append(String.Format("{0}", LeftJoinTechnologySTR))
                End If
            End If

            If Not IsNothing(WithTechnologySB) Then
                WithTechnologySB.Append(String.Format("{0}", WithTechnologySTR))
                AppendWithSB(WithTechnologySB)
            End If

            If Not IsNothing(LeftJoinFactorySB) Then
                If LeftJoinFactorySB.Length = 0 Then
                    LeftJoinFactorySB.Append(String.Format("{0}", LeftJoinFactorySTR))
                End If
            End If



            If Not IsNothing(WithSIFSB) Then

                If LatestValue Then
                    If WithSIFSB.Length = 0 Then
                        WithSIFSB.Append(String.Format("{0}", WithSIFIDSSTRLatest))
                        AppendWithSB(WithSIFSB)
                    End If
                Else
                    If WithSIFSB.Length = 0 Then
                        WithSIFSB.Append(String.Format("{0}", WithSIFIDSSTRNOTLatest))
                        AppendWithSB(WithSIFSB)
                    End If
                End If

                If LeftJoinSIFSB.Length = 0 Then
                    LeftJoinSIFSB.Append(LeftJoinSIFSTR)
                End If

            End If

            If Not IsNothing(WithIDSB) Then

                If LatestValue Then
                    If WithIDSB.Length = 0 Then
                        WithIDSB.Append(String.Format("{0}", WithSIFIDSSTRLatest))
                        AppendWithSB(WithIDSB)
                    End If
                Else
                    If WithIDSB.Length = 0 Then
                        WithIDSB.Append(String.Format("{0}", WithSIFIDSSTRNOTLatest))
                        AppendWithSB(WithIDSB)
                    End If
                End If

                If LeftJoinIDSB.Length = 0 Then
                    LeftJoinIDSB.Append(LeftJoinIDSTR)
                End If

            End If

            If Not IsNothing(AdditionalDisplayFieldSB) Then
                AdditionalDisplayFieldSB.Append(String.Format("{0}", AdditionalDisplayFieldSTR))
            End If
        End If
    End Sub

    Private Function validateString(ByVal p1 As String) As String
        Return p1.Replace("'", "''")
    End Function

End Class
