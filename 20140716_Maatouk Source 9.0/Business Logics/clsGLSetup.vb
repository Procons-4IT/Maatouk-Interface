Public Class clsGLSetup
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        oApplication.Utilities.LoadForm(xml_GLSetup, frm_GLSetup)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oGrid = oForm.Items.Item("1").Specific
        oGrid.DataTable.ExecuteQuery("Select * from [@Z_OACT]")
        oGrid.Columns.Item("Code").Visible = False
        oGrid.Columns.Item("Name").Visible = False
        oGrid.Columns.Item("U_Z_GrLossCredit").TitleObject.Caption = "Green Loss Credit Account"
        oEditTextColumn = oGrid.Columns.Item("U_Z_GrLossCredit")
        oEditTextColumn.ChooseFromListUID = "CFL_2"
        oEditTextColumn.ChooseFromListAlias = "FormatCode"
        oEditTextColumn.LinkedObjectType = "1"
        oGrid.Columns.Item("U_Z_GrLossDebit").TitleObject.Caption = "Green Loss Debit Account"
        oEditTextColumn = oGrid.Columns.Item("U_Z_GrLossDebit")
        oEditTextColumn.ChooseFromListUID = "CFL_3"
        oEditTextColumn.ChooseFromListAlias = "FormatCode"
        oEditTextColumn.LinkedObjectType = "1"
        oGrid.Columns.Item("U_Z_RoLossCredit").TitleObject.Caption = "Rosted Loss Credit Account"
        oEditTextColumn = oGrid.Columns.Item("U_Z_RoLossCredit")
        oEditTextColumn.ChooseFromListUID = "CFL_4"
        oEditTextColumn.ChooseFromListAlias = "FormatCode"
        oEditTextColumn.LinkedObjectType = "1"
        oGrid.Columns.Item("U_Z_RoLossDebit").TitleObject.Caption = "Rosted Loss Debit Account"
        oEditTextColumn = oGrid.Columns.Item("U_Z_RoLossDebit")
        oEditTextColumn.ChooseFromListUID = "CFL_5"
        oEditTextColumn.ChooseFromListAlias = "FormatCode"
        oEditTextColumn.LinkedObjectType = "1"
        oGrid.AutoResizeColumns()
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
        oForm.Freeze(False)
    End Sub
    Public Function AddtoExportUDT(ByVal aform As SAPbouiCOM.Form) As Boolean
        Try
            Dim oUsertable As SAPbobsCOM.UserTable
            Dim strsql, sCode, strUpdateQuery As String
            Dim oRec, oRecUpdate As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oGrid = aform.Items.Item("1").Specific
            oRec.DoQuery("Delete from [@Z_OACT] ")
            If oRec.RecordCount <= 0 Then
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    strsql = oApplication.Utilities.getMaxCode("@Z_OACT", "CODE")
                    If oGrid.DataTable.GetValue("U_Z_GrLossCredit", intRow) <> "" Then
                        oUsertable = oApplication.Company.UserTables.Item("Z_OACT")
                        oUsertable.Code = strsql
                        oUsertable.Name = strsql & "M"
                        oUsertable.UserFields.Fields.Item("U_Z_GrLossCredit").Value = oGrid.DataTable.GetValue("U_Z_GrLossCredit", intRow)
                        oUsertable.UserFields.Fields.Item("U_Z_GrLossDebit").Value = oGrid.DataTable.GetValue("U_Z_GrLossDebit", intRow)
                        oUsertable.UserFields.Fields.Item("U_Z_RoLossCredit").Value = oGrid.DataTable.GetValue("U_Z_RoLossCredit", intRow)
                        oUsertable.UserFields.Fields.Item("U_Z_RoLossDebit").Value = oGrid.DataTable.GetValue("U_Z_RoLossDebit", intRow) 'strAction '"A"
                        If oUsertable.Add <> 0 Then
                            MsgBox(oApplication.Company.GetLastErrorDescription)
                            Return False
                        End If
                    End If
                Next
            End If
            Return True
        Catch ex As Exception
            oApplication.SBO_Application.MessageBox(ex.Message)
        End Try
    End Function
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_GLSetup Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" Then
                                    If AddtoExportUDT(oForm) = True Then
                                        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        oForm.Close()
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1, val2, val3 As String
                                Dim sCHFL_ID, val, val4, val5 As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        If pVal.ItemUID = "1" Then
                                            val = oDataTable.GetValue("FormatCode", 0)
                                            oGrid = oForm.Items.Item("1").Specific
                                            Try
                                                oGrid.DataTable.SetValue(pVal.ColUID, pVal.Row, val)
                                            Catch ex As Exception
                                            End Try
                                        End If
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    oForm.Freeze(False)
                                End Try

                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_GLSetup
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
