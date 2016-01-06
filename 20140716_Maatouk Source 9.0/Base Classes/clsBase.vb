Public MustInherit Class clsBase
    Inherits Object

    Protected _Object As Object
    Protected _FormUID As String
    Protected oForm As SAPbouiCOM.Form
    Protected LookUpOpen As Boolean
    Protected LookUpFrmUID As String

#Region "New"
    Public Sub New()
        MyBase.New()
    End Sub
#End Region
    
#Region "Overridable Functions"
    Public Overridable Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
    End Sub

    Public Overridable Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    End Sub
#End Region
   
#Region "Protected Functions"
    Protected Sub OpenLookUpForm(ByVal xmlFileName As String)
        Try
            oApplication.Utilities.LoadForm(Me._Object, xmlFileName)
            LookUpOpen = True
            LookUpFrmUID = _Object.FrmUID
            oApplication.LookUpCollection.Add(LookUpFrmUID, _FormUID)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Public Properties"
    Public Property FrmUID() As String
        Get
            Return _FormUID
        End Get
        Set(ByVal Value As String)
            _FormUID = Value
        End Set
    End Property

    Public Property IsLookUpOpen() As Boolean
        Get
            Return LookUpOpen
        End Get
        Set(ByVal Value As Boolean)
            LookUpOpen = Value
        End Set
    End Property

    Public Property Form() As SAPbouiCOM.Form
        Get
            Return oForm
        End Get
        Set(ByVal Value As SAPbouiCOM.Form)
            oForm = Value
        End Set
    End Property

    Public ReadOnly Property LookUpFormUID() As String
        Get
            Return LookUpFrmUID
        End Get
    End Property
#End Region

End Class

