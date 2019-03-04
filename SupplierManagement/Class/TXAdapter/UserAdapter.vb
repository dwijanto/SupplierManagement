Public Class UserAdapter
    Implements IAdapter
    Dim DS As DataSet
    Dim DbAdapter1 As DbAdapter = DbAdapter.getInstance
    Public bs As BindingSource
    Dim UserModel1 As New UserModel

    Public Function loaddata() As Boolean Implements IAdapter.loaddata
        Return False
    End Function

    Public Function save() As Boolean Implements IAdapter.save
        Return False
    End Function

    Public ReadOnly Property getApplicantBS As BindingSource
        Get
            Return UserModel1.getApplicantBS
        End Get        
    End Property

    Public ReadOnly Property getApprovalDeptBS As BindingSource
        Get
            Return UserModel1.getApprovalDeptBS
        End Get
    End Property

    Public ReadOnly Property getApprovalDirectorBS As BindingSource
        Get
            Return UserModel1.getApprovalDirectorBS
        End Get
    End Property

    Public ReadOnly Property getApprovalDBBS As BindingSource
        Get
            Return UserModel1.getApprovalDBBS
        End Get
    End Property

    Public ReadOnly Property getApprovalDBListBS As BindingSource
        Get
            Return UserModel1.getApprovalDBListBS
        End Get
    End Property

    Public ReadOnly Property getApprovalDBVIMBS As BindingSource
        Get
            Return UserModel1.getApprovalDBVIMBS
        End Get
    End Property
    Public ReadOnly Property getApprovalFCBS As BindingSource
        Get
            Return UserModel1.getApprovalFCBS
        End Get
    End Property

    Public ReadOnly Property getApprovalVPBS As BindingSource
        Get
            Return UserModel1.getApprovalVPBS
        End Get
    End Property

    Public ReadOnly Property getCCDBBS As BindingSource
        Get
            Return UserModel1.getCCDBBS
        End Get
    End Property
End Class
