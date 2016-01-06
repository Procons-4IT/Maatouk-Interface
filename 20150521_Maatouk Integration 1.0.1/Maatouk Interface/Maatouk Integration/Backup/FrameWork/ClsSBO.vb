Imports System.Data
Imports System.Data.Odbc
Imports System.IO
Imports System.Data.SqlClient
Imports System.Drawing
Public Class ClsSBO

#Region "Declaration"
    ' Private objPackage As frmPackageProcessMonitoring
    Private objutility As clsUtilities
    Public oCompany As SAPbobsCOM.Company
    Public LastErrorDescription As String
    Public LastErrorCode As Integer
    Public intRow As Integer
    Private strFormUID As String
    Public strItemcode, strSourcecolumn, strFormtype, strPono As String
    Public bolScanboxidchoice As Boolean = False
    Dim strConstr As String
    Private dbName, server, provider, uid, pwd, tc, DNSname As String
    Public strConnectString, strConnectODBC As String

#End Region

#Region "Connect Company"
    '*****************************************************************
    'Type               : Procedure    
    'Name               : ConnectCompany
    'Parameter          : 
    'Return Value       : 
    'Author             : DEV-3
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Connect SBO Company
    '******************************************************************
    Public Sub New()

    End Sub
    Public Sub CreateObjectTemp()
        createobjects()
    End Sub
#End Region
   
#Region "Update Error Details"
    Private Sub updateLastErrorDetails(ByVal ErrorCode As Integer)
        LastErrorCode = ErrorCode
        LastErrorDescription = oCompany.GetLastErrorCode() & ":" & oCompany.GetLastErrorDescription()
    End Sub
#End Region

#Region "Create Objects"
    '*****************************************************************
    'Type               : Procedure    
    'Name               : createObjects
    'Parameter          : 
    'Return Value       : 
    'Author             : DEV-9
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Create instants for the Class
    '*****************************************************************
    Private Sub createobjects()
        objutility = New clsUtilities(Me)
    End Sub

#End Region

#Region "Database Function"

#Region "Get Code"
    '*****************************************************************
    'Type               : Function    
    'Name               : GetCode
    'Parameter          : Tablename
    'Return Value       : Maximum Code value in String Format
    'Author             : DEV-9
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Get Maximum Code field value for given Table
    '*****************************************************************
    Public Function GetCode(ByVal sTableName As String) As String
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim sQuery As String
        oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        sQuery = "SELECT Top 1 Code FROM " & sTableName + " ORDER BY Convert(Int,Code) desc"
        oRecSet.DoQuery(sQuery)
        If Not oRecSet.EoF Then
            GetCode = Convert.ToInt32(oRecSet.Fields.Item(0).Value.ToString()) + 1
        Else
            GetCode = "1"
        End If
    End Function

#End Region

#Region "DI Methods"

#Region "GetDateTime"

    '*****************************************************************
    'Type               : Function   
    'Name               : GetDateTimeValue
    'Parameter          : DateString
    'Return Value       : DateFormate
    'Author             : DEV-9
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Convert given string into dateTime Format
    '*****************************************************************

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : GetSBODateString
    'Parameter          : DateTime
    'Return Value       : String
    'Author             : DEV-9
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Convert given  dateTime Format into string format
    '*****************************************************************
    Public Function GetSBODateString(ByVal DateVal As DateTime) As String
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_DateToString(DateVal).Fields.Item(0).Value
    End Function

#End Region

#Region "Business Objects"

    '*****************************************************************
    'Type               : Function   
    'Name               : GetBusinessObject
    'Parameter          : BOobjectTypes
    'Return Value       : Object
    'Author             : DEV-9
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Create instance to the give object
    '*****************************************************************
    Public Function GetBusinessObject(ByVal ObjectType As SAPbobsCOM.BoObjectTypes) As Object
        Return oCompany.GetBusinessObject(ObjectType)
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : CreateUIObject
    'Parameter          : BOCreatableobjectType
    'Return Value       : Object
    'Author             : DEV-9
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Create instance to the give UIObject
    '*****************************************************************
    'Public Function CreateUIObject(ByVal Type As SAPbouiCOM.BoCreatableObjectType) As Object
    '    Return objApplication.CreateObject(Type)
    'End Function

#End Region

#Region "Get Tax Rate"
    '*****************************************************************
    'Type               : Function   
    'Name               : GetTaxRate
    'Parameter          : StrCode
    'Return Value       : string
    'Author             : DEV-9
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Get Tax Value for give Item code
    '*****************************************************************
    Public Function GetTaxRate(ByVal strCode As String) As String
        Dim rsCurr As SAPbobsCOM.Recordset
        Dim strsql, GetTaxRate1 As String
        strsql = ""
        GetTaxRate1 = ""
        strsql = "Select rate from OVTG where code='" & strCode + "'"
        rsCurr = GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rsCurr.DoQuery(strsql)
        GetTaxRate1 = rsCurr.Fields.Item(0).Value
        Return GetTaxRate1
    End Function

#End Region

#End Region

#End Region

#Region "Get Price"
    Public Function GetPrice1(ByVal sPrice As String) As Double
        Dim strlen As Integer
        Dim i As Integer
        Dim strPrice As String
        Dim dblPrice As Double
        strlen = sPrice.Length
        strPrice = ""
        i = 0
        Try
            While i < strlen
                If (Asc(sPrice.Chars(i)) >= 48 And Asc(sPrice.Chars(i)) <= 57) Or Asc(sPrice.Chars(i)) = 46 Or Asc(sPrice.Chars(i)) = 44 Then
                    strPrice = strPrice & sPrice.Chars(i)
                End If
                i = i + 1
            End While
            dblPrice = Convert.ToDouble(strPrice)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        GetPrice1 = dblPrice
    End Function
    Public Function GetPrice(ByVal sPrice As String) As String
        Dim strlen As Integer
        Dim i As Integer
        Dim strPrice As String
        Dim dblPrice As Double
        strlen = sPrice.Length
        strPrice = ""
        i = 0
        Try
            While i < strlen
                If (Asc(sPrice.Chars(i)) >= 48 And Asc(sPrice.Chars(i)) <= 57) Or Asc(sPrice.Chars(i)) = 46 Or Asc(sPrice.Chars(i)) = 44 Then
                    strPrice = strPrice & sPrice.Chars(i)
                End If
                i = i + 1
            End While
            ' dblPrice = Convert.ToDouble(strPrice)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        GetPrice = strPrice
    End Function
#End Region

End Class
