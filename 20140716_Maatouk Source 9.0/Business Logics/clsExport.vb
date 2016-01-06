Imports System.IO
Public Class clsExport
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
    Dim strFileName As String
    Dim strSelectedFilepath, sPath, strSelectedFolderPath As String
    Dim dtDatatable As SAPbouiCOM.DataTable
    Dim blnErrorflag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
#Region "Methods"
    Private Sub LoadForm()
        oApplication.Utilities.LoadForm(xml_Export, frm_Export)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.DataSources.UserDataSources.Add("path", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("from", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("to", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oCombobox = oForm.Items.Item("4").Specific
        oCombobox.ValidValues.Add("All", "All")
        oCombobox.ValidValues.Add("SKU", "SKU")
        ' oCombobox.ValidValues.Add("BP", "Business Partner")
        oCombobox.ValidValues.Add("SO", "Sales Order")
        oCombobox.ValidValues.Add("ARCR", "Supplier Returns")
        oCombobox.ValidValues.Add("PO", "Purchase Order")
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oForm.Items.Item("4").DisplayDesc = True
        oEditText = oForm.Items.Item("6").Specific
        oEditText.DataBind.SetBound(True, "", "path")
        oEditText = oForm.Items.Item("16").Specific
        oEditText.DataBind.SetBound(True, "", "from")
        oEditText = oForm.Items.Item("18").Specific
        oEditText.DataBind.SetBound(True, "", "to")
        AddChooseFromList(oForm)
        oForm.Items.Item("15").Visible = False
        oForm.Items.Item("16").Visible = False
        oForm.Items.Item("17").Visible = False
        oForm.Items.Item("18").Visible = False
    End Sub

#Region "Add Choose From List"
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "2"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "3"
            oCFL = oCFLs.Add(oCFLCreationParams)


            oCFLCreationParams.ObjectType = "4"
            oCFLCreationParams.UniqueID = "4"
            oCFL = oCFLs.Add(oCFLCreationParams)



            oCFLCreationParams.ObjectType = "4"
            oCFLCreationParams.UniqueID = "5"
            oCFL = oCFLs.Add(oCFLCreationParams)



            oCFLCreationParams.ObjectType = "17"
            oCFLCreationParams.UniqueID = "6"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL2

            '' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


            oCFLCreationParams.ObjectType = "17"
            oCFLCreationParams.UniqueID = "7"
            oCFL = oCFLs.Add(oCFLCreationParams)
            ' Adding Conditions to CFL2
            '' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "22"
            oCFLCreationParams.UniqueID = "8"
            oCFL = oCFLs.Add(oCFLCreationParams)
            '' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "22"
            oCFLCreationParams.UniqueID = "9"
            oCFL = oCFLs.Add(oCFLCreationParams)
            '' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()



            oCFLCreationParams.ObjectType = "112"
            oCFLCreationParams.UniqueID = "10"
            oCFL = oCFLs.Add(oCFLCreationParams)
            '' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "112"
            oCFLCreationParams.UniqueID = "11"
            oCFL = oCFLs.Add(oCFLCreationParams)
            '' Adding Conditions to CFL2
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "DocStatus"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "O"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()
        Catch
            MsgBox(Err.Description)
        End Try
    End Sub

#End Region

#Region "Export"
    Private Sub export(ByVal aForm As SAPbouiCOM.Form)
        Dim strvalue As String
        Dim stpath As String
        Try



            oCombobox = aForm.Items.Item("4").Specific
            strvalue = oCombobox.Selected.Value
            stpath = oApplication.Utilities.getEdittextvalue(oForm, "6")
            If stpath = "" Then
                oApplication.Utilities.Message("Folder path missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            If Directory.Exists(stpath) = False Then
                oApplication.Utilities.Message("Folder does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If


            If oApplication.SBO_Application.MessageBox("Do you want to export the selected documents?", , "Yes", "No") = 2 Then
                Exit Sub
            End If
            Try
                AddToExportUDT(aForm)
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End Try
            'companyStorekey = oApplication.Utilities.getStoreKey()
            'If companyStorekey = "" Then
            '    oApplication.Utilities.Message("Define the storekey in the company details", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Exit Sub
            'End If

            Select Case strvalue
                Case "SKU"
                    oApplication.Utilities.ExportSKU(stpath, "SKU")
                Case "SO"
                    oApplication.Utilities.ExportSalesOrder(stpath, "SO")
                Case "ARCR"
                    oApplication.Utilities.ExportARCreditMemo(stpath, "ARCR")
                Case "PO"
                    oApplication.Utilities.ExportPurchaseOrder(stpath, "PO")
                Case "BP"
                    'oApplication.Utilities.ExportSalesOrder(stpath, "BP")
                Case "All"
                    oApplication.Utilities.ExportSKU(stpath, "SKU")
                    oApplication.Utilities.ExportSalesOrder(stpath, "SO")
                    oApplication.Utilities.ExportARCreditMemo(stpath, "ARCR")
                    oApplication.Utilities.ExportPurchaseOrder(stpath, "PO")
            End Select
            oApplication.Utilities.Message("Export process completed successfully.....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

    Private Sub EnableControls(ByVal aForm As SAPbouiCOM.Form)
        oCombobox = aForm.Items.Item("4").Specific
        Dim ostatic, ostatic1 As SAPbouiCOM.StaticText
        oEditText = aForm.Items.Item("16").Specific

        Select Case oCombobox.Selected.Value
            Case "All"
                aForm.Items.Item("15").Visible = False
                aForm.Items.Item("16").Visible = False
                aForm.Items.Item("17").Visible = False
                aForm.Items.Item("18").Visible = False


            Case "SKU"
                ostatic = aForm.Items.Item("15").Specific
                ostatic1 = aForm.Items.Item("17").Specific
                oEditText = aForm.Items.Item("16").Specific
                oEditText.ChooseFromListUID = "4"
                oEditText.ChooseFromListAlias = "ItemCode"
                oEditText.String = ""
                oEditText = aForm.Items.Item("18").Specific
                oEditText.ChooseFromListUID = "5"
                oEditText.ChooseFromListAlias = "ItemCode"
                oEditText.String = ""
                aForm.Items.Item("15").Visible = True
                aForm.Items.Item("16").Visible = True
                aForm.Items.Item("17").Visible = True
                aForm.Items.Item("18").Visible = True


                ostatic.Caption = "Item Code   From"
                ostatic1.Caption = "To"
            Case "BP"
                oEditText = aForm.Items.Item("16").Specific
                oEditText.ChooseFromListUID = "2"
                oEditText.ChooseFromListAlias = "CardCode"
                oEditText.String = ""
                oEditText = aForm.Items.Item("18").Specific
                oEditText.ChooseFromListUID = "3"
                oEditText.ChooseFromListAlias = "CardCode"
                oEditText.String = ""
                ostatic = aForm.Items.Item("15").Specific
                ostatic1 = aForm.Items.Item("17").Specific
                aForm.Items.Item("15").Visible = True
                aForm.Items.Item("16").Visible = True
                aForm.Items.Item("17").Visible = True
                aForm.Items.Item("18").Visible = True

                ostatic.Caption = "Card Code   From"
                ostatic1.Caption = "To"
            Case "SO"
                ostatic = aForm.Items.Item("15").Specific
                ostatic1 = aForm.Items.Item("17").Specific
                oEditText = aForm.Items.Item("16").Specific
                oEditText.ChooseFromListUID = "6"
                oEditText.ChooseFromListAlias = "DocNum"
                oEditText.String = ""
                oEditText = aForm.Items.Item("18").Specific
                oEditText.ChooseFromListUID = "7"
                oEditText.ChooseFromListAlias = "DocNum"
                oEditText.String = ""
                aForm.Items.Item("15").Visible = True
                aForm.Items.Item("16").Visible = True
                aForm.Items.Item("17").Visible = True
                aForm.Items.Item("18").Visible = True
                ostatic.Caption = "Doc.Number  From "
                ostatic1.Caption = "To"
            Case "PO"
                ostatic = aForm.Items.Item("15").Specific
                ostatic1 = aForm.Items.Item("17").Specific
                oEditText = aForm.Items.Item("16").Specific
                oEditText.ChooseFromListUID = "8"
                oEditText.ChooseFromListAlias = "DocNum"
                oEditText.String = ""
                oEditText = aForm.Items.Item("18").Specific
                oEditText.ChooseFromListUID = "9"
                oEditText.ChooseFromListAlias = "DocNum"
                oEditText.String = ""
                aForm.Items.Item("15").Visible = True
                aForm.Items.Item("16").Visible = True
                aForm.Items.Item("17").Visible = True
                aForm.Items.Item("18").Visible = True
                ostatic.Caption = "Doc.Number  From "
                ostatic1.Caption = "To"

            Case "ARCR"
                ostatic = aForm.Items.Item("15").Specific
                ostatic1 = aForm.Items.Item("17").Specific
                oEditText = aForm.Items.Item("16").Specific
                oEditText.ChooseFromListUID = "10"
                oEditText.ChooseFromListAlias = "DocNum"
                oEditText.String = ""
                oEditText = aForm.Items.Item("18").Specific
                oEditText.ChooseFromListUID = "11"
                oEditText.ChooseFromListAlias = "DocNum"
                oEditText.String = ""
                aForm.Items.Item("15").Visible = True
                aForm.Items.Item("16").Visible = True
                aForm.Items.Item("17").Visible = True
                aForm.Items.Item("18").Visible = True
                ostatic.Caption = "Doc.Number  From "
                ostatic1.Caption = "To"

        End Select
    End Sub

#Region "Browse File"

    '*****************************************************************
    'Type               : Procedure    
    'Name               : BrowseFile
    'Parameter          : Form
    'Return Value       : 
    'Author             :  Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Browse a  File
    '******************************************************************
    Private Sub BrowseFile(ByVal Form As SAPbouiCOM.Form)
        'ShowFileDialog(Form)
    End Sub
#End Region

#Region "ShowFileDialog"

    '*****************************************************************
    'Type               : Procedure
    'Name               : ShowFileDialog
    'Parameter          :
    'Return Value       :
    'Author             : Senthil Kumar B 
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To open a File Browser
    '******************************************************************

    Private Sub fillopen()
        Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
        mythr.SetApartmentState(Threading.ApartmentState.STA)
        mythr.Start()
        mythr.Join()

    End Sub

    Private Sub ShowFileDialog()
        Dim oDialogBox As New FolderBrowserDialog
        Dim strFileName, strMdbFilePath As String
        Dim oEdit As SAPbouiCOM.EditText
        Dim oProcesses() As Process
        Try
            oProcesses = Process.GetProcessesByName("SAP Business One")
            If oProcesses.Length <> 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    Dim MyWindow As New WindowWrapper(oProcesses(i).MainWindowHandle)
                    If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                        strMdbFilePath = oDialogBox.SelectedPath
                        strSelectedFilepath = oDialogBox.SelectedPath
                        strFileName = strSelectedFilepath
                        strSelectedFolderPath = strFileName
                        If strSelectedFolderPath.EndsWith("\") Then
                            strSelectedFolderPath = strSelectedFilepath.Substring(0, strSelectedFolderPath.Length - 1)
                        End If
                    Else
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
#End Region

#Region "AddToExportUDT"
    Private Sub AddToExportUDT(ByVal aForm As SAPbouiCOM.Form)
        Dim strChoice, strFrom, strTo, strQuery, strTable, strField, strCondition, strMasterField As String
        Dim oRS As SAPbobsCOM.Recordset
        oCombobox = aForm.Items.Item("4").Specific
        strChoice = oCombobox.Selected.Value
        If strChoice <> "All" Then
            strFrom = oApplication.Utilities.getEdittextvalue(aForm, "16")
            strTo = oApplication.Utilities.getEdittextvalue(aForm, "18")
            strField = ""
            strTable = ""
            strMasterField = ""
            Select Case strChoice
                Case "SKU"
                    strTable = "OITM"
                    strField = "ItemCode"
                    strMasterField = "ItemCode"
                Case "SO"
                    strTable = "ORDR"
                    strField = "DocNum"
                    strMasterField = "DocEntry"
                Case "PO"
                    strTable = "OPOR"
                    strField = "DocNum"
                    strMasterField = "DocEntry"
                Case "BP"
                    strTable = "OCRD"
                    strField = "CardCode"
                    strMasterField = "CardCode"
                Case "ARCR"
                    strTable = "ODRF"
                    strField = "DocNum"
                    strMasterField = "DocEntry"
            End Select
            If strFrom <> "" And strTo <> "" Then
                strCondition = strField & ">='" & strFrom & "' and " & strField & "<='" & strTo & "'"
            ElseIf strFrom <> "" And strTo = "" Then
                strCondition = strField & ">='" & strFrom & "'"
            ElseIf strFrom = "" And strTo <> "" Then
                strCondition = strField & "<='" & strFrom & "'"
            Else
                strCondition = "1=1"
            End If
            strQuery = "Select " & strField & "," & strMasterField & " from " & strTable & " where " & strCondition
            oRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery(strQuery)
            For introw As Integer = 0 To oRS.RecordCount - 1
                oApplication.Utilities.Message("Exporting in process....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oApplication.Utilities.AddtoExportUDT(oRS.Fields.Item(1).Value, oRS.Fields.Item(0).Value, strChoice, "A")
                oRS.MoveNext()
            Next
        End If
    End Sub
#End Region
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Export Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                EnableControls(oForm)

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "12" Then
                                    fillopen()
                                    oEditText = oForm.Items.Item("6").Specific
                                    oEditText.String = strSelectedFilepath
                                ElseIf pVal.ItemUID = "3" Then
                                    export(oForm)

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
                                        ' oGrid = oForm.Items.Item("11").Specific
                                        intChoice = 0
                                        oForm.Freeze(True)

                                        If pVal.ItemUID = "16" Then
                                            If (oCFL.UniqueID = "6" Or oCFL.UniqueID = "7" Or oCFL.UniqueID = "8" Or oCFL.UniqueID = "9" Or oCFL.UniqueID = "10" Or oCFL.UniqueID = "11") Then
                                                val = oDataTable.GetValue("DocNum", 0)
                                            Else
                                                val = oDataTable.GetValue(0, 0)
                                            End If
                                            oApplication.Utilities.setEdittextvalue(oForm, "16", val)
                                        ElseIf pVal.ItemUID = "18" Then
                                            If (oCFL.UniqueID = "6" Or oCFL.UniqueID = "7" Or oCFL.UniqueID = "8" Or oCFL.UniqueID = "9" Or oCFL.UniqueID = "10" Or oCFL.UniqueID = "11") Then
                                                val = oDataTable.GetValue("DocNum", 0)
                                            Else
                                                val = oDataTable.GetValue(0, 0)
                                            End If
                                            oApplication.Utilities.setEdittextvalue(oForm, "18", val)
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

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Export
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
