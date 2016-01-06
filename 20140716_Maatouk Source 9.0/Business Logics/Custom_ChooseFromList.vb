Public Class Custom_ChooseFromList
    Inherits clsBase

    Private oGrid As SAPbouiCOM.Grid
    Private oMatrix As SAPbouiCOM.Matrix
    Private oCFL_Type As CFL_Type
    Private NumToSearch As Integer
    Private SelectedRow As Integer
    Private OrderNum As String

    Public IsCFL_Items As Boolean
    Public BaseFrmUID As String

    Public Sub New()
        MyBase.New()
        IsCFL_Items = False
    End Sub

    Public Enum CFL_Type As Integer
        cfl_QUOTATION = 1
        cfl_ORDER_RETURN
        cfl_ORDER_INVOICE
        sfl_ORDER_ALERT
    End Enum

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        If Not pVal.BeforeAction Then
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_CLICK

                    If pVal.ItemUID = "5" Then
                        If pVal.Row > 0 And pVal.Row <= oMatrix.RowCount Then
                            oMatrix.SelectRow(pVal.Row, True, False)
                            SelectedRow = pVal.Row
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "6" Then
                        If Not Me.IsCFL_Items Then
                            FillBaseDocument()
                        Else
                            Me._Object = oApplication.Collection.Item(BaseFrmUID)
                            'Me._Object.FillPartialReturn(OrderNum, getSelectedItems())
                            oForm.Close()
                        End If

                        'oForm.Close()
                    End If

                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                    If Not Me.IsCFL_Items Then
                        If pVal.Row > 0 And pVal.Row <= oMatrix.RowCount Then
                            oMatrix.SelectRow(pVal.Row, True, False)
                            SelectedRow = pVal.Row
                            FillBaseDocument()
                            'oForm.Close()
                        End If
                    End If


                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                    If pVal.ItemUID = "4" And ((pVal.CharPressed >= 48 And pVal.CharPressed <= 57) Or pVal.CharPressed = 8 Or pVal.CharPressed = 36) And oMatrix.RowCount > 0 Then
                        If oForm.Items.Item("4").Specific.String <> "" Then
                            NumToSearch = CType(oForm.Items.Item("4").Specific.Value, Integer)
                            If NumToSearch > 0 Then
                                Dim NumToCompare As Integer
                                Dim blFound As Boolean = False
                                Dim First As Integer = 1
                                Dim Mid As Integer
                                Dim Last As Integer = oMatrix.RowCount

                                While First <= Last
                                    Mid = Math.Floor((First + Last) / 2)
                                    NumToCompare = CType(oMatrix.Columns.Item("1").Cells.Item(Mid).Specific.Value, Integer)

                                    If NumToSearch < NumToCompare Then
                                        Last = Mid - 1

                                    ElseIf NumToSearch > NumToCompare Then
                                        First = Mid + 1

                                    Else
                                        oMatrix.SelectRow(Mid, True, False)
                                        SelectedRow = Mid
                                        blFound = True
                                        Exit While
                                    End If
                                End While

                                If Not blFound Then
                                    oMatrix.SelectRow(1, True, False)
                                    SelectedRow = 1
                                End If

                            End If
                        Else
                            oMatrix.SelectRow(1, True, False)
                            SelectedRow = 1
                        End If
                    End If

            End Select
        End If
    End Sub
#End Region

#Region "Show Selected Items"
    Public Sub ShowChoosedDocList(ByVal Type As CFL_Type, Optional ByVal CardCode As String = "")
        Dim oParentForm As SAPbouiCOM.Form
        Dim oConditions As SAPbouiCOM.Conditions
        Dim oCondition As SAPbouiCOM.Condition
        Dim sTableName As String

        Try
            oCFL_Type = Type

            oForm.EnableMenu(mnu_ADD, False)
            oForm.EnableMenu(mnu_FIND, False)

            oMatrix = oForm.Items.Item("5").Specific
            oConditions = New SAPbouiCOM.Conditions
            oCondition = oConditions.Add

            Select Case Type

                Case CFL_Type.cfl_QUOTATION

                    sTableName = "@REN_OQUT"
                    oForm.Title = "List of Rental Quotations"
                    With oCondition
                        .BracketOpenNum = 1
                        .Alias = "U_Status"
                        .CondVal = "C"
                        .Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                        .BracketCloseNum = 1
                    End With
                    oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCondition = oConditions.Add
                    With oCondition
                        .BracketOpenNum = 1
                        .Alias = "Status"
                        .CondVal = "O"
                        .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        .BracketCloseNum = 1
                    End With

                Case CFL_Type.sfl_ORDER_ALERT

                    sTableName = "@REN_ORDR"
                    oForm.Title = "List of Rental Orders"
                    With oCondition
                        .BracketOpenNum = 1
                        .Alias = "U_Status"
                        .CondVal = "R"
                        .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        .BracketCloseNum = 1
                    End With

                Case CFL_Type.cfl_ORDER_INVOICE

                    sTableName = "@REN_ORDR"
                    oForm.Title = "List of Rental Orders"
                    With oCondition
                        .BracketOpenNum = 1
                        .Alias = "U_Status"
                        .CondVal = "R"
                        .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        .BracketCloseNum = 1
                    End With
                    oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCondition = oConditions.Add
                    With oCondition
                        .BracketOpenNum = 1
                        .Alias = "U_PInvoice"
                        .CondVal = "Y"
                        .Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                        .BracketCloseNum = 1
                    End With

            End Select

            If CardCode <> "" Then
                oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCondition = oConditions.Add
                With oCondition
                    .BracketOpenNum = 1
                    .Alias = "U_CardCode"
                    .CondVal = CardCode.Trim
                    .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    .BracketCloseNum = 1
                End With
            End If

            oForm.DataSources.DBDataSources.Add(sTableName)
            oForm.DataSources.DBDataSources.Item(sTableName).Query(oConditions)

            With oMatrix.Columns
                .Item("0").Visible = False
                .Item("1").DataBind.SetBound(True, sTableName, "DocNum")
                .Item("2").DataBind.SetBound(True, sTableName, "U_PostDt")
                .Item("3").DataBind.SetBound(True, sTableName, "U_CardName")
                .Item("4").DataBind.SetBound(True, sTableName, "U_Remarks")
            End With

            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            If oMatrix.RowCount > 0 Then
                oMatrix.SelectRow(1, True, False)
                SelectedRow = 1
                oForm.Items.Item("6").Enabled = True
            End If

            oParentForm = oApplication.SBO_Application.Forms.Item(oApplication.LookUpCollection.Item(_FormUID))
            oForm.Left = oParentForm.Left + oParentForm.Width / 7
            oForm.Top = oParentForm.Top + oParentForm.Height / 6
            oForm.Visible = True

        Catch ex As Exception
            Throw ex
        Finally
            If Not oParentForm Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oParentForm)
                oParentForm = Nothing
            End If
            If Not oConditions Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oConditions)
                oConditions = Nothing
            End If
            If Not oCondition Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCondition)
                oCondition = Nothing
            End If
        End Try
    End Sub
#End Region

#Region "Function to be called when opened from Return"

#Region "Show Valid Return Documents"
    Public Sub ShowValidReturnDocuments(ByVal CardCode As String)
        Dim rsOrders As SAPbobsCOM.Recordset
        Dim oDataSrc As SAPbouiCOM.DBDataSource
        Dim InsertRecordAt As Integer

        Try
            oDataSrc = oForm.DataSources.DBDataSources.Add("@REN_ORDR")
            oMatrix = oForm.Items.Item("5").Specific
            oMatrix.Columns.Item("1").DataBind.SetBound(True, "@REN_ORDR", "DocNum")
            oMatrix.Columns.Item("2").Visible = False
            oMatrix.Columns.Item("3").DataBind.SetBound(True, "@REN_ORDR", "U_CardName")
            oMatrix.Columns.Item("4").DataBind.SetBound(True, "@REN_ORDR", "U_Remarks")

            strSQL = "Select T1.DocEntry, U_PostDt, U_CardName, U_Remarks " & _
                        "From [@REN_ORDR] T1 " & _
                        "Inner Join [@REN_RDR1] T2 On T2.DocEntry = T1.DocEntry " & _
                        "Where U_Status = 'R' And T2.U_ItemCode Is Not NULL And T2.U_ReqdQty > T2.U_RetQty And U_CardCode = '" & CardCode & "' " & _
                        "Group By T1.DocEntry, U_PostDt, U_CardName, U_Remarks " & _
                        "Order By T1.DocEntry "
            oApplication.Utilities.ExecuteSQL(rsOrders, strSQL)

            While Not rsOrders.EoF
                InsertRecordAt = oDataSrc.Size - 1

                oDataSrc.InsertRecord(InsertRecordAt)

                oDataSrc.SetValue("DocNum", InsertRecordAt, rsOrders.Fields.Item(0).Value)
                oDataSrc.SetValue("U_CardName", InsertRecordAt, rsOrders.Fields.Item(2).Value)
                oDataSrc.SetValue("U_Remarks", InsertRecordAt, rsOrders.Fields.Item(3).Value)

                rsOrders.MoveNext()
            End While

            oDataSrc.RemoveRecord(oDataSrc.Size - 1)
            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            If oMatrix.RowCount > 0 Then
                oMatrix.SelectRow(1, True, False)
                SelectedRow = 1
                oForm.Items.Item("6").Enabled = True
            End If
            oForm.Title = "List of Rental Orders"
            oCFL_Type = CFL_Type.cfl_ORDER_RETURN

        Catch ex As Exception
            Throw ex
        Finally
            oForm.Visible = True
        End Try
    End Sub
#End Region

#End Region

#Region "Fill Base Document"
    Private Sub FillBaseDocument()
        Dim DocNum As Integer

        DocNum = oMatrix.Columns.Item("1").Cells.Item(SelectedRow).Specific.String

        Select Case oCFL_Type
            Case CFL_Type.cfl_QUOTATION
                Me._Object = oApplication.Collection.Item(oApplication.LookUpCollection.Item(_FormUID))
                Me._Object.FillBaseDocFromQuotation(DocNum)
                Me.oForm.Close()

            Case CFL_Type.cfl_ORDER_INVOICE
                Me._Object = oApplication.Collection.Item(oApplication.LookUpCollection.Item(_FormUID))
                oApplication.Utilities.Message("Preparing Invoice... please wait.", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Me.oForm.Close()
                Me._Object.Form.Freeze(True)
                Me._Object.FillBaseDocFromOrder(DocNum)

            Case CFL_Type.cfl_ORDER_RETURN
                Dim sDocList As String
                For Count As Integer = 1 To oMatrix.RowCount
                    If oMatrix.IsRowSelected(Count) Then
                        oMatrix.GetLineData(Count)

                        sDocList += oMatrix.Columns.Item("1").Cells.Item(Count).Specific.String & ","
                    End If
                Next

                sDocList = sDocList.Remove(sDocList.LastIndexOf(","), 1)

                'Me._Object = New clsItemsInOrder
                'oApplication.Utilities.LoadForm(Me._Object, "ItemsInOrder.xml")
                'CType(Me._Object, clsItemsInOrder).ShowItemsInOrders(sDocList)
                'CType(Me._Object, clsItemsInOrder).BaseForm = oApplication.LookUpCollection.Item(_FormUID)
                oForm.Close()

            Case CFL_Type.sfl_ORDER_ALERT
                Me._Object = oApplication.Collection.Item(oApplication.LookUpCollection.Item(_FormUID))
                Me._Object.FillDocNum(DocNum)
                Me.oForm.Close()

        End Select
    End Sub
#End Region

#Region "Finalize"
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        Try
            If Not oMatrix Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix)
                oMatrix = Nothing
            End If
        Catch ex As Exception
            Throw ex
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

End Class
