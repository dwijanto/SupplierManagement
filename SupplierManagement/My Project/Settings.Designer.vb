﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.18444
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace My
    
    <Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "10.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase
        
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()),MySettings)
        
#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(ByVal sender As Global.System.Object, ByVal e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region
        
        Public Shared ReadOnly Property [Default]() As MySettings
            Get
                
#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
                Return defaultInstance
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("ODBC;DSN=PostgreSQLhon03nt;")>  _
        Public Property oExCon() As String
            Get
                Return CType(Me("oExCon"),String)
            End Get
            Set
                Me("oExCon") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("host=localhost;port=5432;database=LogisticDb20150120;CommandTimeout=1000;TimeOut="& _ 
            "1000;")>  _
        Public Property conTest() As String
            Get
                Return CType(Me("conTest"),String)
            End Get
            Set
                Me("conTest") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("host=hon14nt;port=5432;database=LogisticDb;commandTimeout=1000;Timeout=1000;")>  _
        Public Property conLive() As String
            Get
                Return CType(Me("conLive"),String)
            End Get
            Set
                Me("conLive") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("ODBC;DSN=PostgreSQLhon03nt;")>  _
        Public Property oExConDevLive() As String
            Get
                Return CType(Me("oExConDevLive"),String)
            End Get
            Set
                Me("oExConDevLive") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("ODBC;DSN=PostgreSQLhon14ntUnicode;")>  _
        Public Property oExConTestLive() As String
            Get
                Return CType(Me("oExConTestLive"),String)
            End Get
            Set
                Me("oExConTestLive") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("smtp.seb.com")>  _
        Public ReadOnly Property smtpclient() As String
            Get
                Return CType(Me("smtpclient"),String)
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FSD_DocumentType() As String
            Get
                Return CType(Me("FSD_DocumentType"),String)
            End Get
            Set
                Me("FSD_DocumentType") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FSD_VendorCode() As String
            Get
                Return CType(Me("FSD_VendorCode"),String)
            End Get
            Set
                Me("FSD_VendorCode") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FSD_VendorName() As String
            Get
                Return CType(Me("FSD_VendorName"),String)
            End Get
            Set
                Me("FSD_VendorName") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FSD_ShortName() As String
            Get
                Return CType(Me("FSD_ShortName"),String)
            End Get
            Set
                Me("FSD_ShortName") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FSD_ProjectName() As String
            Get
                Return CType(Me("FSD_ProjectName"),String)
            End Get
            Set
                Me("FSD_ProjectName") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property FSD_DocumentDate() As Boolean
            Get
                Return CType(Me("FSD_DocumentDate"),Boolean)
            End Get
            Set
                Me("FSD_DocumentDate") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property FSD_DTP1() As Date
            Get
                Return CType(Me("FSD_DTP1"),Date)
            End Get
            Set
                Me("FSD_DTP1") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property FSD_DTP2() As Date
            Get
                Return CType(Me("FSD_DTP2"),Date)
            End Get
            Set
                Me("FSD_DTP2") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FSD_SPM() As String
            Get
                Return CType(Me("FSD_SPM"),String)
            End Get
            Set
                Me("FSD_SPM") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FSD_PM() As String
            Get
                Return CType(Me("FSD_PM"),String)
            End Get
            Set
                Me("FSD_PM") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FSD_ProductType() As String
            Get
                Return CType(Me("FSD_ProductType"),String)
            End Get
            Set
                Me("FSD_ProductType") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FSD_VendorStatus() As String
            Get
                Return CType(Me("FSD_VendorStatus"),String)
            End Get
            Set
                Me("FSD_VendorStatus") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FSD_SBUName() As String
            Get
                Return CType(Me("FSD_SBUName"),String)
            End Get
            Set
                Me("FSD_SBUName") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property FSD_LatestDocument() As Boolean
            Get
                Return CType(Me("FSD_LatestDocument"),Boolean)
            End Get
            Set
                Me("FSD_LatestDocument") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_ProductType() As String
            Get
                Return CType(Me("FASS_ProductType"),String)
            End Get
            Set
                Me("FASS_ProductType") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_SEBBrand() As String
            Get
                Return CType(Me("FASS_SEBBrand"),String)
            End Get
            Set
                Me("FASS_SEBBrand") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_RangeCode() As String
            Get
                Return CType(Me("FASS_RangeCode"),String)
            End Get
            Set
                Me("FASS_RangeCode") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_CMMF() As String
            Get
                Return CType(Me("FASS_CMMF"),String)
            End Get
            Set
                Me("FASS_CMMF") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_CommercialCode() As String
            Get
                Return CType(Me("FASS_CommercialCode"),String)
            End Get
            Set
                Me("FASS_CommercialCode") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_Project() As String
            Get
                Return CType(Me("FASS_Project"),String)
            End Get
            Set
                Me("FASS_Project") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_LatestSocialAudit() As String
            Get
                Return CType(Me("FASS_LatestSocialAudit"),String)
            End Get
            Set
                Me("FASS_LatestSocialAudit") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_VendorFinanceProgram() As String
            Get
                Return CType(Me("FASS_VendorFinanceProgram"),String)
            End Get
            Set
                Me("FASS_VendorFinanceProgram") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_SupplierStatus() As String
            Get
                Return CType(Me("FASS_SupplierStatus"),String)
            End Get
            Set
                Me("FASS_SupplierStatus") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_SBU() As String
            Get
                Return CType(Me("FASS_SBU"),String)
            End Get
            Set
                Me("FASS_SBU") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_FamilyCode() As String
            Get
                Return CType(Me("FASS_FamilyCode"),String)
            End Get
            Set
                Me("FASS_FamilyCode") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_SupplierCategory() As String
            Get
                Return CType(Me("FASS_SupplierCategory"),String)
            End Get
            Set
                Me("FASS_SupplierCategory") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_FPPanel() As String
            Get
                Return CType(Me("FASS_FPPanel"),String)
            End Get
            Set
                Me("FASS_FPPanel") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_CPPanel() As String
            Get
                Return CType(Me("FASS_CPPanel"),String)
            End Get
            Set
                Me("FASS_CPPanel") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_SPM() As String
            Get
                Return CType(Me("FASS_SPM"),String)
            End Get
            Set
                Me("FASS_SPM") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_PM() As String
            Get
                Return CType(Me("FASS_PM"),String)
            End Get
            Set
                Me("FASS_PM") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_Province() As String
            Get
                Return CType(Me("FASS_Province"),String)
            End Get
            Set
                Me("FASS_Province") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_Country() As String
            Get
                Return CType(Me("FASS_Country"),String)
            End Get
            Set
                Me("FASS_Country") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_MainCustomers() As String
            Get
                Return CType(Me("FASS_MainCustomers"),String)
            End Get
            Set
                Me("FASS_MainCustomers") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_SuppliersProductBrand() As String
            Get
                Return CType(Me("FASS_SuppliersProductBrand"),String)
            End Get
            Set
                Me("FASS_SuppliersProductBrand") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_MainProductSold() As String
            Get
                Return CType(Me("FASS_MainProductSold"),String)
            End Get
            Set
                Me("FASS_MainProductSold") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_SBUIdentitySheet() As String
            Get
                Return CType(Me("FASS_SBUIdentitySheet"),String)
            End Get
            Set
                Me("FASS_SBUIdentitySheet") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_Family() As String
            Get
                Return CType(Me("FASS_Family"),String)
            End Get
            Set
                Me("FASS_Family") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_OEMODM() As String
            Get
                Return CType(Me("FASS_OEMODM"),String)
            End Get
            Set
                Me("FASS_OEMODM") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_Material() As String
            Get
                Return CType(Me("FASS_Material"),String)
            End Get
            Set
                Me("FASS_Material") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property FASS_LatestOnly() As Boolean
            Get
                Return CType(Me("FASS_LatestOnly"),Boolean)
            End Get
            Set
                Me("FASS_LatestOnly") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property FASS_Technology() As String
            Get
                Return CType(Me("FASS_Technology"),String)
            End Get
            Set
                Me("FASS_Technology") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.ConnectionString),  _
         Global.System.Configuration.DefaultSettingValueAttribute("host=hon14nt;port=5432;database=LogisticDb;commandTimeout=1000;Timeout=1000;")>  _
        Public ReadOnly Property Connectionstring1() As String
            Get
                Return CType(Me("Connectionstring1"),String)
            End Get
        End Property
    End Class
End Namespace

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.SupplierManagement.My.MySettings
            Get
                Return Global.SupplierManagement.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace
