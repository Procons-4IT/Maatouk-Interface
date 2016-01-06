Imports System.Xml
Imports System.IO
Module SubMain

#Region "Declaration"
    Private objMain As ClsSBO
    Public objMainCompany, objRemoteCompany As SAPbobsCOM.Company
    Public strCustomerDSNName, strVendorDSNName, strConnectionString As String
    Public lngServiceTimer As Long = 0
    Public lngBPMasterDataTimer As Long = 0
    Public strBPList As String
    Public strServiceCallList As String
    Public intDays As Integer
    Public strCust, strSOHead, strSOLines, strFilePath, strExportFilePaty, strSuccessFolder, strFileStart, strErrorLogPath, strErrorFileName As String
    Public strLocalServertype, strMainServertype, strSQLServer, strSQLUID, strSQLPWD, strSAPCompany, strSAPUID, strSAPPWD, strMainSQLServer, strMainSQLUID, strMainSQLPWD, strMainSAPCompany, strMainSAPUID, strMainSAPPWD, strMainLicenseServer, strLocalLicensserver As String
    Public strTmp As String
    '  Public oCompany As SAPbobsCOM.Company

  
#End Region

#Region "Main Method"
    '*****************************************************************************
    'Type               : Procedure   
    'Name               : main
    'Parameter          : 
    'Return Value       : 
    'Author             : DEV-9
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Create Instance to MainClass and Initialize Applicaiton 
    '******************************************************************************
    Public Sub main()
        Dim objForm As New frmLoginScreen
        System.Windows.Forms.Application.Run(objForm)
    End Sub
#End Region

#Region "Close Application"

    Public Sub CloseApp()
        System.Windows.Forms.Application.Exit()
    End Sub

#End Region

    Public Function GetConInfo(ByVal ConType As String) As String
        Dim objXMLDoc As New System.Xml.XmlDocument
        Dim objNode As Xml.XmlNode
        Dim intLoop As Integer

        objXMLDoc.Load(System.Windows.Forms.Application.StartupPath & "\ConnectionInfo.xml")
        With objXMLDoc.GetElementsByTagName("Parameter")
            For intLoop = 0 To .Count - 1
                objNode = .Item(intLoop)
                If objNode.Attributes("Type").Value = ConType Then
                    Return objNode.Attributes("Value").Value
                End If
            Next
        End With
        Return ""
    End Function

    Public Sub WriteConInfo(ByVal strType As String, ByVal strValue As String)
        Dim objXMLDoc As New System.Xml.XmlDocument
        Dim objNode As Xml.XmlNode
        Dim intLoop As Integer
        Try
            objXMLDoc.Load(System.Windows.Forms.Application.StartupPath & "\ConnectionInfo.xml")
            With objXMLDoc.GetElementsByTagName("Parameter")
                For intLoop = 0 To .Count - 1
                    objNode = .Item(intLoop)
                    If objNode.Attributes("Type").Value = strType Then
                        objNode.Attributes("Value").Value = strValue
                    End If
                Next
            End With
            objXMLDoc.Save(System.Windows.Forms.Application.StartupPath & "\ConnectionInfo.xml")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Module