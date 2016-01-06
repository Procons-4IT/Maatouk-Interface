Imports System.IO
Public Class clsStockRequest
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
        oForm = oApplication.Utilities.LoadForm(xml_InvSO, frm_InvSO)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        AddChooseFromList(oForm)
        Databind(oForm)
    End Sub
#Region "AddCFL"
    Private Sub AddChooseFromList(ByVal aform As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = aform.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding Header GL CFL, one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "4"
            oCFLCreationParams.UniqueID = "CFL2"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "64"
            oCFLCreationParams.UniqueID = "CFL3"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "4"
            oCFLCreationParams.UniqueID = "CFL4"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "4"
            oCFLCreationParams.UniqueID = "CFL5"
            oCFL = oCFLs.Add(oCFLCreationParams)


        Catch
            oApplication.Utilities.Message(Err.Description, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub DeleteRow(ByVal aGrid As SAPbouiCOM.Grid)
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            If aGrid.Rows.IsSelected(intRow) Then
                aGrid.DataTable.Rows.Remove(intRow)
                Exit Sub
            End If
        Next
    End Sub
#End Region


#Region "DataBind"
    Private Sub Databind(ByVal aForm As SAPbouiCOM.Form)

        aForm.DataSources.UserDataSources.Add("BPCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        aForm.DataSources.UserDataSources.Add("BPName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        aForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE)
        aForm.DataSources.UserDataSources.Add("NumAtCard", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oEditText = aForm.Items.Item("4").Specific
        oEditText.DataBind.SetBound(True, "", "BPCode")
        oEditText.ChooseFromListUID = "CFL1"
        oEditText.ChooseFromListAlias = "CardCode"
        oEditText = aForm.Items.Item("6").Specific
        oEditText.DataBind.SetBound(True, "", "BPName")
        oEditText = aForm.Items.Item("8").Specific
        oEditText.DataBind.SetBound(True, "", "DocDate")
        oEditText = aForm.Items.Item("10").Specific
        oEditText.DataBind.SetBound(True, "", "NumAtCard")


        oGrid = aForm.Items.Item("11").Specific
        dtTemp = oGrid.DataTable

        oEditTextColumn = oGrid.Columns.Item(0)
        oEditTextColumn.ChooseFromListUID = "CFL5"
        oEditTextColumn.ChooseFromListAlias = "CodeBars"
        oEditTextColumn.LinkedObjectType = "4"


        oEditTextColumn = oGrid.Columns.Item(1)
        oEditTextColumn.ChooseFromListUID = "CFL2"
        oEditTextColumn.ChooseFromListAlias = "ItemCode"
        oEditTextColumn.LinkedObjectType = "4"


        oEditTextColumn = oGrid.Columns.Item(2)
        oEditTextColumn.ChooseFromListUID = "CFL4"
        oEditTextColumn.ChooseFromListAlias = "ItemName"
        oEditTextColumn.LinkedObjectType = "4"

        oEditTextColumn = oGrid.Columns.Item(3)
        oEditTextColumn.ChooseFromListUID = "CFL3"
        oEditTextColumn.ChooseFromListAlias = "WhsCode"
        oEditTextColumn.LinkedObjectType = "64"


        dtTemp.ExecuteQuery("SELECT T0.[CodeBars],T0.[ItemCode], T0.[Dscription], T0.[WhsCode],  T0.[OpenQty] FROM INV1 T0 where 1=2")
        oGrid.DataTable = dtTemp
        FormatGrid(oGrid)

    End Sub
#End Region

    Private Sub FormatGrid(ByVal aGrid As SAPbouiCOM.Grid)
        aGrid.Columns.Item(0).TitleObject.Caption = "BarCode"
        aGrid.Columns.Item(0).Editable = True
        oEditTextColumn = aGrid.Columns.Item(0)
        oEditTextColumn.LinkedObjectType = "4"

        aGrid.Columns.Item(1).TitleObject.Caption = "Item Code"
        aGrid.Columns.Item(1).Editable = True
        oEditTextColumn = aGrid.Columns.Item(1)
        oEditTextColumn.LinkedObjectType = "4"

        aGrid.Columns.Item(2).TitleObject.Caption = "Description"
        aGrid.Columns.Item(2).Editable = True
        'oEditTextColumn.ChooseFromListUID = "CFL4"
        'oEditTextColumn.ChooseFromListAlias = "Code"

        aGrid.Columns.Item(3).TitleObject.Caption = "Warehouse"
        aGrid.Columns.Item(3).Editable = True
        oEditTextColumn = aGrid.Columns.Item(3)
        oEditTextColumn.LinkedObjectType = "64"

        'aGrid.Columns.Item(3).TitleObject.Caption = "TaxCode"
        'aGrid.Columns.Item(3).Editable = True
        'oEditTextColumn = aGrid.Columns.Item(3)
        'oEditTextColumn = aGrid.Columns.Item(3)
        'oEditTextColumn.ChooseFromListUID = "CFL4"
        'oEditTextColumn.ChooseFromListAlias = "Code"
        'oEditTextColumn.LinkedObjectType = "128"

        aGrid.Columns.Item(4).TitleObject.Caption = "Quantity"
        aGrid.Columns.Item(4).Editable = True
        aGrid.AutoResizeColumns()
        aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub

#Region "Validate Grid Details"
    Private Function ValidateGrid(ByVal aGrid As SAPbouiCOM.Grid) As Boolean
        Dim strItem, strWhs, strTax, spath As String
        Dim dblQty As Double
        Dim strMessage As String
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            strItem = aGrid.DataTable.GetValue("ItemCode", intRow)
            If strItem <> "" Then
                strWhs = aGrid.DataTable.GetValue("WhsCode", intRow)
                ' strTax = aGrid.DataTable.GetValue("TaxCode", intRow)
                dblQty = aGrid.DataTable.GetValue("OpenQty", intRow)
                spath = System.Windows.Forms.Application.StartupPath & "\ErrorLog.txt"
                strMessage = "CardCode : " & strCardCode & "--> Item Code :" & strItem & "-->Warehouse : " & strWhs
                If strWhs = "" Then
                    WriteErrorlog("Error in Matching Sales Order -->Line Number : " & intRow + 1 & "  Warehouse Missing... ", spath)
                    blnError = True
                End If
                'If strTax = "" Then
                '    WriteErrorlog("Error in Matching Sales Order -->Line Number : " & intRow + 1 & "  TaxCode Missing... ", spath)
                '    blnError = True
                'End If
                If dblQty <= 0 Then
                    WriteErrorlog("Error in Matching Sales Order -->Line Number : " & intRow + 1 & "  Quantity Should be greater than Zero... ", spath)
                    blnError = True
                End If
            End If
        Next
        If blnError = True Then
            Return False
        Else
            Return True
        End If

    End Function
#End Region

#Region "Check the Item Availablitity"
    Private Function CheckItem(ByVal strItem As String, ByVal dblQty As Double, ByVal strWhs As String, ByVal intLineNo As Integer) As Boolean
        Dim otemprs As SAPbobsCOM.Recordset
        Dim strSQL, spath As String
        Dim strMessage As String
        spath = System.Windows.Forms.Application.StartupPath & "\ErrorLog.txt"
        strMessage = "CardCode : " & strCardCode & "--> Item Code :" & strItem & "-->Warehouse : " & strWhs
        otemprs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ' strSQL = "Select * from RDR1  where Quantity-isnull(U_RemQty,0)>0 and  docentry in (Select docentry from ORDR where Cardcode='" & strCardCode & "' and DocStatus='O') and  linestatus='O' and  Itemcode='" & strItem & "' and WhsCode='" & strWhs & "'"
        strSQL = "Select * from RDR1  where isnull(OpenCreQty,0)>0 and  docentry in (Select docentry from ORDR where Cardcode='" & strCardCode & "' and DocStatus='O') and  linestatus='O' and  Itemcode='" & strItem & "' and WhsCode='" & strWhs & "'"
        otemprs.DoQuery(strSQL)
        If otemprs.RecordCount <= 0 Then
            WriteErrorlog("Error in Matching Sales Order -->Line Number : " & intLineNo & "  No Open Sales Order for the " & strMessage, spath)
            blnError = True
        End If
        strSQL = "Select  isnull(sum(OpenCreQty),0) from RDR1  where  OpenCreQty>0 and docentry in (Select docentry from ORDR where Cardcode='" & strCardCode & "' and DocStatus='O') and linestatus='O' and  Itemcode='" & strItem & "' and WhsCode='" & strWhs & "'"
        otemprs.DoQuery(strSQL)
        If dblQty > otemprs.Fields.Item(0).Value Then
            WriteErrorlog("Error in matching Sales order --> Line Number : " & intLineNo & "   Invoice quantity greater than the open sales order quantity for :" & strMessage, spath)
            blnError = True
        End If
        If blnError = True Then
            Return False
        Else
            Return True
        End If
    End Function
#End Region

#Region "Validations"
    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strWhs, stritem, spath As String
        Dim dblQty As Double
        Dim oTemp As SAPbobsCOM.Recordset
        Try
            spath = System.Windows.Forms.Application.StartupPath & "\ErrorLog.txt"
            If File.Exists(spath) Then
                File.Delete(spath)
            End If
            blnError = False

            Dim strCardCode, strDocDate, NumAtCard As String
            strCardCode = oApplication.Utilities.getEdittextvalue(aForm, "4")
            If strCardCode = "" Then
                oApplication.Utilities.Message("Customer code missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            strDocDate = oApplication.Utilities.getEdittextvalue(aForm, "8")
            If strDocDate = "" Then
                oApplication.Utilities.Message("Document date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            NumAtCard = oApplication.Utilities.getEdittextvalue(aForm, "10")
            If NumAtCard = "" Then
                oApplication.Utilities.Message("Customer reference no is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If


            oGrid = aForm.Items.Item("11").Specific
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oTemp.DoQuery("Select * from OINV where CardCode='" & strCardCode & "' and NumAtCard='" & NumAtCard & "'")
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("AR Invoice already created for this Customer and Customer ref.No ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            oTemp.DoQuery("Select * from ORDR where CardCode='" & strCardCode & "' and DocStatus='O'")
            If oTemp.RecordCount <= 0 Then
                oApplication.SBO_Application.MessageBox("No open sales orders for the customer :" & strCardCode)
                WriteErrorlog(" Error in matching Sales order -->No open sales orders for the customer :" & strCardCode, spath)
                blnError = True
                Return False
            End If

            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                stritem = oGrid.DataTable.GetValue("ItemCode", intRow)
                If oGrid.DataTable.GetValue("ItemCode", intRow) <> "" Then
                    stritem = oGrid.DataTable.GetValue("ItemCode", intRow)
                    If 1 = 1 Then
                        dblQty = oGrid.DataTable.GetValue("OpenQty", intRow)
                    Else
                        dblQty = 0
                    End If
                    strWhs = oGrid.DataTable.GetValue("WhsCode", intRow)
                    If CheckItem(stritem, dblQty, strWhs, intRow) = False Then
                    End If
                End If
            Next

            If ValidateGrid(oGrid) = False Then
                blnError = True
            Else
                blnError = False
            End If
            If blnError = True Then
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Return True
    End Function
#End Region


#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_InvSO
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_CloseOrderLines
                    If oApplication.SBO_Application.MessageBox("Do you want to close the open sales order lines?.", , "Yes", "No") = 2 Then
                        Exit Sub
                    Else
                        oApplication.Utilities.CloseOpenSOLines()
                    End If
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
    Private Sub WriteErrorlog(ByVal aMessage As String, ByVal aPath As String)
        Dim aSw As System.IO.StreamWriter
        Try
            If File.Exists(aPath) Then
            End If
            aSw = New StreamWriter(aPath, True)
            aSw.WriteLine(aMessage)
            aSw.Flush()
            aSw.Close()
            aSw.Dispose()
        Catch ex As Exception
            MsgBox("test")
        End Try
    End Sub

#Region "Update Stock Qty to Sales Order Lines"

    Private Function CreateInvoices(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oTempRs, oTempSO, oUpdateRS As SAPbobsCOM.Recordset
        Dim strWhs, strItem, strdate, strcondition, strcardcode, strNumAtCard, strTaxcode As String
        Dim dblTransQty, dblSOQty As Double
        Dim dblAssignQty, intRowCount As Double
        Dim dtDate As Date
        Dim st As String
        Dim dtResult As SAPbouiCOM.DataTable
        oTempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempSO = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oUpdateRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If 1 = 1 Then
                strcardcode = oApplication.Utilities.getEdittextvalue(aForm, "4")
                strNumAtCard = oApplication.Utilities.getEdittextvalue(aForm, "10")
                strdate = oApplication.Utilities.getEdittextvalue(aForm, "8")
                If strdate <> "" Then
                    dtDate = oApplication.Utilities.GetDateTimeValue(strdate)
                End If
                oGrid = aForm.Items.Item("11").Specific
                dtResult = aForm.DataSources.DataTables.Item("dtResult")
                dtResult.Rows.Clear()
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    strItem = oGrid.DataTable.GetValue("ItemCode", intRow)
                    If strItem <> "" Then
                        dblTransQty = oGrid.DataTable.GetValue("OpenQty", intRow)
                        ' strTaxcode = oGrid.DataTable.GetValue("TaxCode", intRow)
                        strWhs = oGrid.DataTable.GetValue("WhsCode", intRow)
                        strcondition = "CardCode='" & strcardcode & "'"
                        st = "Select LineNum,isnull(OpenCreQty,0),Quantity,DocEntry,TaxCode from RDR1 where ItemCode='" & strItem & "' and  whscode='" & strWhs & "' and OpenCreQty >0 and DocEntry in (Select DocEntry from ORDR where " & strcondition & ") and Linestatus='O' order by DocEntry"
                        oTempSO.DoQuery(st)
                        For intLoop As Integer = 0 To oTempSO.RecordCount - 1
                            dblAssignQty = 0
                            If dblTransQty <= 0 Then
                                Exit For
                            End If
                            dblSOQty = oTempSO.Fields.Item(1).Value
                            If dblSOQty >= dblTransQty Then
                                dblAssignQty = dblTransQty
                            Else
                                dblAssignQty = dblSOQty
                            End If
                            strTaxcode = oTempSO.Fields.Item(4).Value
                            dtResult.Rows.Add()
                            intRowCount = dtResult.Rows.Count - 1
                            dtResult.SetValue("ItemCode", intRowCount, strItem)
                            dtResult.SetValue("BaseEntry", intRowCount, oTempSO.Fields.Item("DocEntry").Value)
                            dtResult.SetValue("BaseLine", intRowCount, oTempSO.Fields.Item("LineNum").Value)
                            dtResult.SetValue("TaxCode", intRowCount, strTaxcode)
                            dtResult.SetValue("Quantity", intRowCount, dblAssignQty)
                            dblTransQty = dblTransQty - dblAssignQty
                            oTempSO.MoveNext()
                        Next
                    End If

                Next
                If dtResult.Rows.Count > 0 Then
                    Dim oDoc As SAPbobsCOM.Documents
                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    oDoc.CardCode = strcardcode
                    oDoc.NumAtCard = strNumAtCard
                    oDoc.DocDate = dtDate
                    For introw As Integer = 0 To dtResult.Rows.Count - 1
                        If introw > 0 Then
                            oDoc.Lines.Add()
                        End If
                        oDoc.Lines.SetCurrentLine(introw)
                        oDoc.Lines.BaseType = 17
                        oDoc.Lines.BaseEntry = dtResult.GetValue("BaseEntry", introw)
                        oDoc.Lines.BaseLine = dtResult.GetValue("BaseLine", introw)
                        oDoc.Lines.Quantity = dtResult.GetValue("Quantity", introw)
                        ' oDoc.Lines.VatGroup = "0"
                        oDoc.Lines.TaxCode = dtResult.GetValue("TaxCode", introw)
                    Next
                    If oDoc.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    Else
                        Dim strDocNum As String
                        oApplication.Company.GetNewObjectCode(strDocNum)
                        oTempRs.DoQuery("Select * from OINV where DocEntry=" & strDocNum)
                        oApplication.SBO_Application.MessageBox("AR Invoice :  " & oTempRs.Fields.Item("DocNum").Value & " Created successfully")
                        Return True
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region



#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_InvSO Then
                Select Case pVal.BeforeAction
                    Case True
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "3" Then
                            oMode = pVal.FormMode
                            If 1 = 1 Then
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                blnError = False
                                strCardCode = oApplication.Utilities.getEdittextvalue(oForm, "4")
                                If Validation(oForm) = False Then
                                    Dim sPath As String
                                    blnError = True
                                    oApplication.SBO_Application.MessageBox("Error Occured...")
                                    sPath = System.Windows.Forms.Application.StartupPath & "\ErrorLog.txt"
                                    If File.Exists(sPath) = False Then
                                        Exit Sub
                                    End If
                                    Dim x As System.Diagnostics.ProcessStartInfo
                                    x = New System.Diagnostics.ProcessStartInfo
                                    x.UseShellExecute = True
                                    sPath = System.Windows.Forms.Application.StartupPath & "\ErrorLog.txt"
                                    x.FileName = sPath
                                    System.Diagnostics.Process.Start(x)
                                    x = Nothing
                                    BubbleEvent = False
                                    Exit Sub
                                Else
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    If CreateInvoices(oForm) Then
                                        oForm.Close()
                                    End If
                                End If
                            End If
                        End If
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "14" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oGrid = oForm.Items.Item("11").Specific
                                    DeleteRow(oGrid)
                                ElseIf pVal.ItemUID = "13" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oGrid = oForm.Items.Item("11").Specific
                                    If oGrid.DataTable.GetValue("ItemCode", oGrid.DataTable.Rows.Count - 1) <> "" Then
                                        oGrid.DataTable.Rows.Add()
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val As String
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
                                        oGrid = oForm.Items.Item("11").Specific
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If ((pVal.ItemUID = "11" And (pVal.ColUID = "ItemCode" Or pVal.ColUID = "Dscription" Or pVal.ColUID = "CodeBars"))) Then

                                            codebar = oDataTable.GetValue("CodeBars", 0)
                                            val = oDataTable.GetValue("ItemCode", 0)
                                            val1 = oDataTable.GetValue("ItemName", 0)
                                            oGrid.DataTable.SetValue("CodeBars", pVal.Row, codebar)
                                            oGrid.DataTable.SetValue("Dscription", pVal.Row, val1)
                                            oGrid.DataTable.SetValue("ItemCode", pVal.Row, val)
                                            If pVal.Row = oGrid.DataTable.Rows.Count - 1 Then
                                                oGrid.DataTable.Rows.Add()
                                            End If
                                        ElseIf pVal.ItemUID = "11" And pVal.ColUID = "WhsCode" Then
                                            val = oDataTable.GetValue("WhsCode", 0)
                                            oGrid.DataTable.SetValue("WhsCode", pVal.Row, val)
                                        ElseIf pVal.ItemUID = "4" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            val1 = oDataTable.GetValue("CardName", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "6", val1)
                                            oApplication.Utilities.setEdittextvalue(oForm, "4", val)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
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

End Class
