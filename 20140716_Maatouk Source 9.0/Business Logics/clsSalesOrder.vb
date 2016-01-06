Public Class clsSalesOrder
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
    Private oBP As SAPbobsCOM.BusinessPartners
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_SalesOrder Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED


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
                Case mnu_InvSO
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
                If oForm.TypeEx = frm_itemmaster Or oForm.TypeEx = "-150" Then
                    oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    If oItem.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        oApplication.Utilities.AddtoExportUDT(oItem.ItemCode, oItem.ItemCode, "SKU", "A")
                    End If
                ElseIf oForm.TypeEx = frm_BPMaster Or oForm.TypeEx = "-134" Then
                    oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                    If oBP.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        oApplication.Utilities.AddtoExportUDT(oBP.CardCode, oBP.CardCode, "BP", "A")
                    End If
                ElseIf oForm.TypeEx = frm_SalesOrder Or oForm.TypeEx = "-139" Then
                    oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                    If oInvoice.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If oInvoice.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oApplication.Utilities.AddtoExportUDT(oInvoice.DocEntry, oInvoice.DocNum, "SO", "A")
                        End If
                    End If
                ElseIf oForm.TypeEx = frm_ARCreditMemo Or oForm.TypeEx = "-179" Then
                    oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                    If oInvoice.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If oInvoice.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oApplication.Utilities.AddtoExportUDT(oInvoice.DocEntry, oInvoice.DocNum, "ARCR", "A")
                        End If
                    End If
                ElseIf oForm.TypeEx = frm_PurchaseOrder Or oForm.TypeEx = "-142" Then
                    oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                    If oInvoice.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If oInvoice.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oApplication.Utilities.AddtoExportUDT(oInvoice.DocEntry, oInvoice.DocNum, "PO", "A")
                        End If
                    End If
                End If
            End If
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm
                If oForm.TypeEx = frm_itemmaster Or oForm.TypeEx = "-150" Then
                    oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    If oItem.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        oApplication.Utilities.AddtoExportUDT(oItem.ItemCode, oItem.ItemCode, "SKU", "U")
                    End If
                ElseIf oForm.TypeEx = frm_BPMaster Or oForm.TypeEx = "-134" Then
                    oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                    If oBP.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        oApplication.Utilities.AddtoExportUDT(oBP.CardCode, oBP.CardCode, "BP", "U")
                    End If
                ElseIf oForm.TypeEx = frm_SalesOrder Or oForm.TypeEx = "-139" Then
                    oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                    If oInvoice.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If oInvoice.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oApplication.Utilities.AddtoExportUDT(oInvoice.DocEntry, oInvoice.DocNum, "SO", "U")
                        End If
                    End If
                ElseIf oForm.TypeEx = frm_ARCreditMemo Or oForm.TypeEx = "-179" Then
                    oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                    If oInvoice.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If oInvoice.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oApplication.Utilities.AddtoExportUDT(oInvoice.DocEntry, oInvoice.DocNum, "ARCR", "U")
                        End If
                    End If
                ElseIf oForm.TypeEx = frm_PurchaseOrder Or oForm.TypeEx = "-142" Then
                    oInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                    If oInvoice.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If oInvoice.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oApplication.Utilities.AddtoExportUDT(oInvoice.DocEntry, oInvoice.DocNum, "PO", "U")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
End Class
