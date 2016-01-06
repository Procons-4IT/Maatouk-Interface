Imports System.IO
Imports System.Text
Public Class clsMain

#Region "Declaration"
    Public strDbServerName, strDbUserName, strDbPassword As String
    Public strCompanyName, strCompanyUserName, strCopmanyPassword As String
    Public objSBOAPI As ClsSBO
    Private objutity As clsUtilities
#End Region

#Region "Methods"

    Public Sub New()

    End Sub

#Region "Initialise"
    '*****************************************************************
    'Type               :  Function    
    'Name               : Initialise
    'Parameter          :
    'Return Value       : Boolean
    'Author             : DEV-3
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Initialise the Application and Create Table
    '******************************************************************

    Public Function Initialise(ByVal SBOCompany As SAPbobsCOM.Company) As Boolean
        objSBOAPI = New ClsSBO
        objSBOAPI.oCompany = SBOCompany
        objutity = New clsUtilities(objSBOAPI)
        objSBOAPI.CreateObjectTemp()
        Dim objLogin As frmLoginScreen
        objLogin = New frmLoginScreen
        Return True
    End Function
#End Region

#End Region

    
End Class
