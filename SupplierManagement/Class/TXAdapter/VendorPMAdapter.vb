﻿Imports System.Text
Public Class VendorPMAdapter
    Implements IAdapter
    Implements IToolbarAction
    Dim DS As DataSet
    Dim DbAdapter1 As DbAdapter = DbAdapter.getInstance
    Public BS As BindingSource
    Dim VendorPMModel1 As VendorPMModel

    Public ReadOnly Property GetTableVendorFamilyPM As DataTable
        Get
            Return DS.Tables("VendorFamilyPM").Copy()
        End Get
    End Property

    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim BS As New BindingSource
            BS.DataSource = GetTableVendorFamilyPM
            BS.Sort = VendorPMModel1.SortField
            Return BS
        End Get
    End Property

    Public Function LoadData() As Boolean Implements IAdapter.loaddata
        Dim myret As Boolean = False
        VendorPMModel1 = New VendorPMModel
        DS = New DataSet
        If VendorPMModel1.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).PrimaryKey = pk
            DS.Tables(0).Columns("id").AutoIncrement = True
            DS.Tables(0).Columns("id").AutoIncrementSeed = -1
            DS.Tables(0).Columns("id").AutoIncrementStep = -1
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret
    End Function


    Public Function save() As Boolean Implements IAdapter.save
        Dim myret As Boolean = False
        BS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If

        Return myret
    End Function

    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format(VendorPMModel1.FilterField, value)
        End Set
    End Property

    Public Function GetCurrentRecord() As System.Data.DataRowView Implements IToolbarAction.GetCurrentRecord
        Return BS.Current
    End Function

    Public Function GetNewRecord() As System.Data.DataRowView Implements IToolbarAction.GetNewRecord
        Return BS.AddNew
    End Function

    Public Sub RemoveAt(ByVal value As Integer) Implements IToolbarAction.RemoveAt
        BS.RemoveAt(value)
    End Sub

    Public Function Save(ByVal mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Dim myret As Boolean = False
        If VendorPMModel1.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function
End Class
