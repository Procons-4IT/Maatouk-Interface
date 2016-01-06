Public Class clsUtilities

#Region "Declartion"
    Private objCompany As SAPbobsCOM.Company
    Private objclsSBO As ClsSBO

#End Region

#Region "Methods"

#Region "Constructor"

    Public Sub New(ByVal objSBO As ClsSBO)
        objclsSBO = objSBO
    End Sub

#End Region

#End Region

End Class
