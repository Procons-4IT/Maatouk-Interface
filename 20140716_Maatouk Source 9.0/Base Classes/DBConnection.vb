'**************************************************************
'Name			        :   DBConnection
'Purpose		        :   Create DB Server Login
'Created Date       	:	03/03/06
'Last Modified By		:	Manu
'Modified Date        	:
'**************************************************************
Imports System.IO
Public Class DBConnection
    Inherits clsBase

#Region "Declarations"
    '---------- Declaration of Controls ------------
    Private objFrm As SAPbouiCOM.Form
    '-------------------------------------------------

    '---------- Declaration of Variable used -------
    Private strPath, strErrMsg, strSql As String
    Private lnErrCode As Long
    Dim blnFlag As Boolean = False
    '-------------------------------------------------

    '----------- Declaration of Classes and Objects Used ---------
    'Private oComFunc As New ComFunc
    'Private oCreateFunc As New CreateFunction
    Private oWrite As System.IO.StreamWriter
    'Private Shared ThreadColse As New Threading.Thread(AddressOf CloseApp)
    '-------------------------------------------------
#End Region

#Region "Public Functions"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    '******** On Click of OK button
                    '******** Check Connection & Create Function
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True Then
                        If oForm.Items.Item("etUser").Specific.String = "" Then
                            oApplication.SBO_Application.SetStatusBarMessage("Invalid Login", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                            Exit Sub
                        Else
                            oWrite = File.CreateText(Application.StartupPath & "\DBLogin.ini")
                            oWrite.WriteLine(oForm.Items.Item("etUser").Specific.Value)
                            oWrite.WriteLine(oForm.Items.Item("etPwd").Specific.Value)
                            oWrite.Close()
                        End If

                    End If
                    '******** On Click of Cancel button
                    '******** Terminate AddOn Payroll
                    If pVal.ItemUID = "2" And pVal.BeforeAction = True Then

                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region
End Class
